#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '..', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
const PASSWORD = '99';

// All possible input fields in column H
const ALL_INPUT_FIELDS = [
    // Customer Details
    { id: 'customer_name', cell: 'H12', label: 'Customer Name', type: 'text', required: true, section: 'Customer Details' },
    { id: 'address', cell: 'H13', label: 'Address', type: 'text', required: true, section: 'Customer Details' },
    { id: 'postcode', cell: 'H14', label: 'Postcode', type: 'text', required: false, section: 'Customer Details' },
    
    // ENERGY USE - CURRENT ELECTRICITY TARIFF
    { id: 'single_day_rate', cell: 'H20', label: 'Single / Day Rate (pence per kWh)', type: 'number', required: true, section: 'Energy Use' },
    { id: 'night_rate', cell: 'H21', label: 'Night Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'off_peak_hours', cell: 'H22', label: 'No. of Off-Peak Hours', type: 'number', required: false, section: 'Energy Use' },
    
    // ENERGY USE - NEW ELECTRICITY TARIFF
    { id: 'new_day_rate', cell: 'H24', label: 'Day Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'new_night_rate', cell: 'H25', label: 'Night Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    
    // ENERGY USE - ELECTRICITY CONSUMPTION
    { id: 'annual_usage', cell: 'H27', label: 'Estimated Annual Usage (kWh)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'standing_charge', cell: 'H28', label: 'Standing Charge (pence per day)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'annual_spend', cell: 'H29', label: 'Annual Spend (£)', type: 'number', required: false, section: 'Energy Use' },
    
    // ENERGY USE - EXPORT TARIFF
    { id: 'export_tariff_rate', cell: 'H31', label: 'Export Tariff Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    
    // EXISTING SYSTEM
    { id: 'commissioning_date', cell: 'H36', label: 'Approximate Commissioning Date', type: 'text', required: false, section: 'Existing System' },
    { id: 'sem_percentage', cell: 'H37', label: 'Percentage of above SEM used to quote self-consumption savings', type: 'number', required: false, section: 'Existing System' },
    
    // NEW SYSTEM - SOLAR
    { id: 'panel_manufacturer', cell: 'H42', label: 'Panel Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'panel_model', cell: 'H43', label: 'Panel Model', type: 'text', required: false, section: 'New System' },
    { id: 'no_of_arrays', cell: 'H44', label: 'No. of Arrays', type: 'number', required: false, section: 'New System' },
    
    // NEW SYSTEM - BATTERY
    { id: 'battery_manufacturer', cell: 'H48', label: 'Battery Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'battery_model', cell: 'H49', label: 'Battery Model', type: 'text', required: false, section: 'New System' },
    { id: 'battery_extended_warranty', cell: 'H50', label: 'Battery Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
    { id: 'battery_replacement_cost', cell: 'H51', label: 'Battery Replacement Cost (£)', type: 'number', required: false, section: 'New System' },
    
    // NEW SYSTEM - SOLAR/HYBRID INVERTER
    { id: 'solar_inverter_manufacturer', cell: 'H55', label: 'Solar/Hybrid Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'solar_inverter_model', cell: 'H56', label: 'Solar/Hybrid Inverter Model', type: 'text', required: false, section: 'New System' },
    { id: 'solar_inverter_extended_warranty', cell: 'H59', label: 'Solar Inverter Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
    { id: 'solar_inverter_replacement_cost', cell: 'H60', label: 'Solar Inverter Replacement Cost (£)', type: 'number', required: false, section: 'New System' },
    
    // NEW SYSTEM - BATTERY INVERTER
    { id: 'battery_inverter_manufacturer', cell: 'H62', label: 'Battery Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'battery_inverter_model', cell: 'H63', label: 'Battery Inverter Model', type: 'text', required: false, section: 'New System' },
    { id: 'battery_inverter_extended_warranty', cell: 'H66', label: 'Battery Inverter Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
    { id: 'battery_inverter_replacement_cost', cell: 'H67', label: 'Battery Inverter Replacement Cost (£)', type: 'number', required: false, section: 'New System' }
];

function analyzeCellInDetail(cell, cellAddress) {
    const analysis = {
        exists: false,
        hasValue: false,
        hasFormula: false,
        isLocked: false,
        isHidden: false,
        isProtected: false,
        hasValidation: false,
        cellType: null,
        value: null,
        formula: null,
        protection: null,
        style: null
    };

    if (!cell) {
        return analysis;
    }

    analysis.exists = true;
    analysis.hasValue = cell.v !== undefined && cell.v !== null;
    analysis.value = cell.v;
    analysis.cellType = cell.t; // s=string, n=number, b=boolean, etc.
    analysis.hasFormula = !!cell.f;
    analysis.formula = cell.f;

    // Check cell protection/locking
    if (cell.l) {
        analysis.isLocked = cell.l.locked;
        analysis.isHidden = cell.l.hidden;
        analysis.protection = cell.l;
    }

    // Check cell style (might indicate if cell is disabled)
    if (cell.s) {
        analysis.style = cell.s;
    }

    // Check for data validation
    if (cell.v) {
        analysis.hasValidation = !!cell.v;
    }

    return analysis;
}

function isCellActuallyEnabled(cell, cellAddress) {
    if (!cell) {
        return { enabled: false, reason: 'Cell not found' };
    }

    // Check if cell has formula (calculated field - should be disabled)
    if (cell.f) {
        return { enabled: false, reason: 'Has formula (calculated field)' };
    }

    // Check if cell is locked
    if (cell.l && cell.l.locked) {
        return { enabled: false, reason: 'Cell is locked' };
    }

    // Check if cell is hidden
    if (cell.l && cell.l.hidden) {
        return { enabled: false, reason: 'Cell is hidden' };
    }

    // Check cell style - look for disabled/grayed out indicators
    if (cell.s) {
        const style = cell.s;
        
        // Check for grayed out text (disabled appearance)
        if (style.font && style.font.color && style.font.color.rgb === '808080') {
            return { enabled: false, reason: 'Cell appears disabled (grayed out)' };
        }
        
        // Check for disabled background
        if (style.fill && style.fill.fgColor && style.fill.fgColor.rgb === 'F2F2F2') {
            return { enabled: false, reason: 'Cell appears disabled (gray background)' };
        }
    }

    // Check if cell value indicates it's disabled
    if (cell.v === '#N/A' || cell.v === '#VALUE!' || cell.v === '#REF!') {
        return { enabled: false, reason: 'Cell has error value' };
    }

    // If we get here, the cell appears to be enabled
    return { enabled: true, reason: 'Cell appears enabled' };
}

function getAccurateDynamicInputs(excelFilePath = EXCEL_FILE_PATH) {
    console.log('🔍 Getting ACCURATE Dynamic Inputs from Excel');
    console.log('=============================================');
    
    try {
        // Check if file exists
        if (!fs.existsSync(excelFilePath)) {
            throw new Error(`Excel file not found: ${excelFilePath}`);
        }
        
        console.log(`📁 Found Excel file: ${excelFilePath}`);
        
        // Read the Excel file with all possible options
        console.log('📖 Reading Excel file with detailed analysis...');
        const workbook = XLSX.readFile(excelFilePath, { 
            password: PASSWORD,
            cellStyles: true,
            cellDates: true,
            cellFormula: true,
            cellNF: true,
            cellHTML: true
        });
        
        console.log('✅ Excel file read successfully');
        
        // Get the Inputs worksheet
        const inputsSheet = workbook.Sheets['Inputs'];
        if (!inputsSheet) {
            throw new Error('Inputs sheet not found');
        }
        
        console.log('✅ Found Inputs sheet');
        
        // Analyze all input fields
        console.log(`\n🔍 Analyzing ${ALL_INPUT_FIELDS.length} input fields in detail...`);
        console.log('==============================================================');
        
        const results = [];
        let enabledCount = 0;
        let disabledCount = 0;
        
        for (const field of ALL_INPUT_FIELDS) {
            const cell = inputsSheet[field.cell];
            const analysis = analyzeCellInDetail(cell, field.cell);
            const enabledStatus = isCellActuallyEnabled(cell, field.cell);
            
            const result = {
                ...field,
                analysis,
                enabled: enabledStatus.enabled,
                reason: enabledStatus.reason,
                value: analysis.value || ''
            };
            
            results.push(result);
            
            if (result.enabled) {
                enabledCount++;
            } else {
                disabledCount++;
            }
        }
        
        // Filter to only enabled fields for the response
        const enabledFields = results.filter(field => field.enabled).map(field => ({
            id: field.id,
            label: field.label,
            type: field.type,
            required: field.required,
            value: field.value,
            cellReference: field.cell,
            section: field.section
        }));
        
        // Summary
        console.log('\n📊 ACCURATE ANALYSIS SUMMARY');
        console.log('=============================');
        console.log(`Total fields analyzed: ${ALL_INPUT_FIELDS.length}`);
        console.log(`✅ Enabled fields: ${enabledCount}`);
        console.log(`❌ Disabled fields: ${disabledCount}`);
        console.log(`📈 Success rate: ${((enabledCount / ALL_INPUT_FIELDS.length) * 100).toFixed(1)}%`);
        
        // Show enabled fields
        console.log('\n✅ ENABLED FIELDS:');
        console.log('==================');
        enabledFields.forEach(field => {
            console.log(`${field.cellReference}: ${field.label} = "${field.value}"`);
        });
        
        // Show disabled fields with detailed reasons
        console.log('\n❌ DISABLED FIELDS:');
        console.log('===================');
        const disabledFields = results.filter(field => !field.enabled);
        disabledFields.forEach(field => {
            console.log(`${field.cell}: ${field.label}`);
            console.log(`   └─ Reason: ${field.reason}`);
            if (field.analysis.hasFormula) {
                console.log(`   └─ Formula: ${field.analysis.formula}`);
            }
            if (field.analysis.isLocked) {
                console.log(`   └─ Locked: ${field.analysis.isLocked}`);
            }
            if (field.analysis.isHidden) {
                console.log(`   └─ Hidden: ${field.analysis.isHidden}`);
            }
        });
        
        // Group by section for better analysis
        console.log('\n📋 RESULTS BY SECTION:');
        console.log('======================');
        
        const sections = {};
        results.forEach(result => {
            if (!sections[result.section]) {
                sections[result.section] = [];
            }
            sections[result.section].push(result);
        });
        
        Object.keys(sections).forEach(sectionName => {
            const sectionFields = sections[sectionName];
            const enabledInSection = sectionFields.filter(f => f.enabled).length;
            
            console.log(`\n${sectionName} (${enabledInSection}/${sectionFields.length} enabled):`);
            console.log('-'.repeat(sectionName.length + 20));
            
            sectionFields.forEach(field => {
                const status = field.enabled ? '✅' : '❌';
                console.log(`${status} ${field.cell}: ${field.label}`);
                if (!field.enabled) {
                    console.log(`   └─ ${field.reason}`);
                }
            });
        });
        
        const response = {
            success: true,
            message: `Found ${enabledCount} enabled input fields`,
            inputFields: enabledFields,
            summary: {
                total: ALL_INPUT_FIELDS.length,
                enabled: enabledCount,
                disabled: disabledCount,
                successRate: (enabledCount / ALL_INPUT_FIELDS.length) * 100
            }
        };
        
        // Export detailed results for debugging
        const outputPath = path.join(__dirname, 'accurate-excel-inputs-results.json');
        fs.writeFileSync(outputPath, JSON.stringify({
            timestamp: new Date().toISOString(),
            excelFile: excelFilePath,
            response,
            detailedResults: results,
            sections: sections
        }, null, 2));
        
        console.log(`\n💾 Detailed results saved to: ${outputPath}`);
        
        return response;
        
    } catch (error) {
        console.error('❌ Error getting accurate dynamic inputs:', error);
        return {
            success: false,
            message: 'Failed to get dynamic inputs',
            error: error.message
        };
    }
}

// Run the function if called directly
if (require.main === module) {
    getAccurateDynamicInputs();
}

module.exports = { getAccurateDynamicInputs };
