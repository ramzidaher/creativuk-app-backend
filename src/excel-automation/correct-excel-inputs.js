#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '..', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
const PASSWORD = '99';

// All possible input fields in column H
const ALL_INPUT_FIELDS = [
    // Customer Details (always enabled)
    { id: 'customer_name', cell: 'H12', label: 'Customer Name', type: 'text', required: true, section: 'Customer Details' },
    { id: 'address', cell: 'H13', label: 'Address', type: 'text', required: true, section: 'Customer Details' },
    { id: 'postcode', cell: 'H14', label: 'Postcode', type: 'text', required: false, section: 'Customer Details' },
    
    // ENERGY USE - CURRENT ELECTRICITY TARIFF
    { id: 'single_day_rate', cell: 'H19', label: 'Single / Day Rate (pence per kWh)', type: 'number', required: true, section: 'Energy Use' },
    { id: 'night_rate', cell: 'H20', label: 'Night Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'off_peak_hours', cell: 'H21', label: 'No. of Off-Peak Hours', type: 'number', required: false, section: 'Energy Use' },
    
    // ENERGY USE - NEW ELECTRICITY TARIFF
    { id: 'new_day_rate', cell: 'H23', label: 'Day Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'new_night_rate', cell: 'H24', label: 'Night Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    
    // ENERGY USE - ELECTRICITY CONSUMPTION
    { id: 'annual_usage', cell: 'H26', label: 'Estimated Annual Usage (kWh)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'standing_charge', cell: 'H27', label: 'Standing Charge (pence per day)', type: 'number', required: false, section: 'Energy Use' },
    { id: 'annual_spend', cell: 'H28', label: 'Annual Spend (£)', type: 'number', required: false, section: 'Energy Use' },
    
    // ENERGY USE - EXPORT TARIFF
    { id: 'export_tariff_rate', cell: 'H30', label: 'Export Tariff Rate (pence per kWh)', type: 'number', required: false, section: 'Energy Use' },
    
    // EXISTING SYSTEM
    { id: 'existing_sem', cell: 'H34', label: 'Existing SEM', type: 'number', required: false, section: 'Existing System' },
    { id: 'commissioning_date', cell: 'H35', label: 'Approximate Commissioning Date', type: 'text', required: false, section: 'Existing System' },
    { id: 'sem_percentage', cell: 'H36', label: 'Percentage of above SEM used to quote self-consumption savings', type: 'number', required: false, section: 'Existing System' },
    
    // NEW SYSTEM - SOLAR
    { id: 'panel_manufacturer', cell: 'H41', label: 'Panel Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'panel_model', cell: 'H42', label: 'Panel Model', type: 'text', required: false, section: 'New System' },
    { id: 'no_of_arrays', cell: 'H43', label: 'No. of Arrays', type: 'number', required: false, section: 'New System' },
    
    // NEW SYSTEM - BATTERY
    { id: 'battery_manufacturer', cell: 'H45', label: 'Battery Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'battery_model', cell: 'H46', label: 'Battery Model', type: 'text', required: false, section: 'New System' },
    { id: 'battery_extended_warranty', cell: 'H48', label: 'Battery Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
    { id: 'battery_replacement_cost', cell: 'H49', label: 'Battery Replacement Cost (£)', type: 'number', required: false, section: 'New System' },
    
    // NEW SYSTEM - SOLAR/HYBRID INVERTER
    { id: 'solar_inverter_manufacturer', cell: 'H51', label: 'Solar/Hybrid Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'solar_inverter_model', cell: 'H52', label: 'Solar/Hybrid Inverter Model', type: 'text', required: false, section: 'New System' },
    { id: 'solar_inverter_extended_warranty', cell: 'H55', label: 'Solar Inverter Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
    { id: 'solar_inverter_replacement_cost', cell: 'H56', label: 'Solar Inverter Replacement Cost (£)', type: 'number', required: false, section: 'New System' },
    
    // NEW SYSTEM - BATTERY INVERTER
    { id: 'battery_inverter_manufacturer', cell: 'H58', label: 'Battery Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'battery_inverter_model', cell: 'H59', label: 'Battery Inverter Model', type: 'text', required: false, section: 'New System' },
    { id: 'battery_inverter_extended_warranty', cell: 'H62', label: 'Battery Inverter Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
    { id: 'battery_inverter_replacement_cost', cell: 'H63', label: 'Battery Inverter Replacement Cost (£)', type: 'number', required: false, section: 'New System' }
];

function isCellCorrectlyEnabled(cell) {
    if (!cell) {
        return { enabled: false, reason: 'Cell not found' };
    }

    // Check if cell has formula (calculated field - should be disabled)
    if (cell.f) {
        return { enabled: false, reason: 'Has formula (calculated field)' };
    }

    // Check cell style for enabled indicators FIRST (CORRECT LOCATION)
    if (cell.s && cell.s.fgColor && cell.s.fgColor.rgb) {
        const bgColor = cell.s.fgColor.rgb;
        
        // Light gray background indicates enabled input field
        if (bgColor === 'E8E8E8') {
            return { enabled: true, reason: 'Light gray background (enabled input)' };
        }
        
        // Light green background indicates enabled input field
        if (bgColor === 'E3F1CB' || bgColor === 'E7F3D1') {
            return { enabled: true, reason: 'Light green background (enabled input)' };
        }
        
        // Dark gray background indicates disabled
        if (bgColor === '595959') {
            return { enabled: false, reason: 'Dark gray background (disabled)' };
        }
    }

    // Check if cell value is "undefined" (this indicates disabled state)
    if (cell.v === 'undefined') {
        return { enabled: false, reason: 'Cell value is undefined (disabled)' };
    }

    // Check if cell type is 'z' (error/undefined type) - but only if no background color was found
    if (cell.t === 'z') {
        return { enabled: false, reason: 'Cell type is error/undefined' };
    }

    // Check if cell is locked
    if (cell.l && cell.l.locked) {
        return { enabled: false, reason: 'Cell is locked' };
    }

    // Check if cell is hidden
    if (cell.l && cell.l.hidden) {
        return { enabled: false, reason: 'Cell is hidden' };
    }

    // If we get here, the cell appears to be enabled
    return { enabled: true, reason: 'Cell appears enabled' };
}

function getCorrectDynamicInputs(excelFilePath = EXCEL_FILE_PATH) {
    console.log('🔍 Getting CORRECT Dynamic Inputs from Excel');
    console.log('===========================================');
    
    try {
        // Check if file exists
        if (!fs.existsSync(excelFilePath)) {
            throw new Error(`Excel file not found: ${excelFilePath}`);
        }
        
        console.log(`📁 Found Excel file: ${excelFilePath}`);
        
        // Read the Excel file
        console.log('📖 Reading Excel file...');
        const workbook = XLSX.readFile(excelFilePath, { 
            password: PASSWORD,
            cellStyles: true,
            cellDates: true,
            cellFormula: true
        });
        
        console.log('✅ Excel file read successfully');
        
        // Get the Inputs worksheet
        const inputsSheet = workbook.Sheets['Inputs'];
        if (!inputsSheet) {
            throw new Error('Inputs sheet not found');
        }
        
        console.log('✅ Found Inputs sheet');
        
        // Analyze all input fields based on correct background color reading
        console.log(`\n🔍 Analyzing ${ALL_INPUT_FIELDS.length} input fields with correct color detection...`);
        console.log('================================================================');
        
        const results = [];
        let enabledCount = 0;
        let disabledCount = 0;
        
        for (const field of ALL_INPUT_FIELDS) {
            const cell = inputsSheet[field.cell];
            const enabledStatus = isCellCorrectlyEnabled(cell, field.cell);
            
            const result = {
                ...field,
                enabled: enabledStatus.enabled,
                reason: enabledStatus.reason,
                value: cell ? (cell.v === 'undefined' ? '' : cell.v) : '',
                cellType: cell ? cell.t : null,
                hasFormula: cell ? !!cell.f : false,
                backgroundColor: cell && cell.s && cell.s.fgColor ? cell.s.fgColor.rgb : null
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
        console.log('\n📊 CORRECT ANALYSIS SUMMARY');
        console.log('===========================');
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
            console.log(`   └─ Value: "${field.value}"`);
            console.log(`   └─ Background: ${field.backgroundColor || 'None'}`);
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
        const outputPath = path.join(__dirname, 'correct-excel-inputs-results.json');
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
        console.error('❌ Error getting correct dynamic inputs:', error);
        return {
            success: false,
            message: 'Failed to get dynamic inputs',
            error: error.message
        };
    }
}

// Run the function if called directly
if (require.main === module) {
    getCorrectDynamicInputs();
}

module.exports = { getCorrectDynamicInputs };
