#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '..', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
const PASSWORD = '99';

function analyzeCellProperties(cell, sheet, cellAddress) {
    const analysis = {
        hasValue: false,
        hasFormula: false,
        isLocked: false,
        isHidden: false,
        hasValidation: false,
        dependsOn: [],
        controls: [],
        cellType: null,
        value: null,
        formula: null
    };

    if (!cell) return analysis;

    // Basic properties
    analysis.hasValue = cell.v !== undefined && cell.v !== null;
    analysis.value = cell.v;
    analysis.cellType = cell.t; // s=string, n=number, b=boolean, etc.
    analysis.hasFormula = !!cell.f;
    analysis.formula = cell.f;

    // Check for cell protection/locking
    if (cell.l) {
        analysis.isLocked = cell.l.locked;
        analysis.isHidden = cell.l.hidden;
    }

    // Check for data validation
    if (cell.v) {
        analysis.hasValidation = !!cell.v;
    }

    // Analyze formula dependencies
    if (cell.f) {
        analysis.dependsOn = extractFormulaDependencies(cell.f);
    }

    return analysis;
}

function extractFormulaDependencies(formula) {
    const dependencies = [];
    
    // Extract cell references from formula (simplified)
    const cellRefs = formula.match(/[A-Z]+\d+/g) || [];
    dependencies.push(...cellRefs);
    
    return dependencies;
}

function checkFieldAvailability(field, sheet) {
    const cell = sheet[field.cell];
    const analysis = analyzeCellProperties(cell, sheet, field.cell);
    
    // Check if this field depends on radio button selections
    const radioButtonDependencies = getRadioButtonDependencies(field.cell);
    
    // Check if dependent cells are selected
    const dependenciesMet = checkDependenciesMet(radioButtonDependencies, sheet);
    
    // A field is truly enabled if:
    // 1. Cell exists and has no formula (not calculated)
    // 2. Cell is not locked
    // 3. All dependencies are met (radio buttons selected appropriately)
    const isTrulyEnabled = analysis.hasValue && 
                          !analysis.hasFormula && 
                          !analysis.isLocked && 
                          dependenciesMet;
    
    return {
        ...field,
        analysis,
        radioButtonDependencies,
        dependenciesMet,
        isTrulyEnabled,
        reason: getReasonForStatus(analysis, dependenciesMet)
    };
}

function getRadioButtonDependencies(cellAddress) {
    // Map of cell addresses to their radio button dependencies
    const dependencyMap = {
        // ENERGY USE - CURRENT ELECTRICITY TARIFF
        'H20': { dependsOn: 'B20', requires: 'SingleRate' }, // Single/Day Rate
        'H21': { dependsOn: 'B20', requires: 'DualRate' },   // Night Rate
        'H22': { dependsOn: 'B20', requires: 'DualRate' },   // Off-Peak Hours
        
        // ENERGY USE - NEW ELECTRICITY TARIFF
        'H24': { dependsOn: 'B23', requires: 'Yes' },        // New Day Rate
        'H25': { dependsOn: 'B23', requires: 'Yes' },        // New Night Rate
        
        // ENERGY USE - ELECTRICITY CONSUMPTION
        'H27': { dependsOn: 'B27', requires: 'Yes' },        // Annual Usage
        'H28': { dependsOn: 'B27', requires: 'Yes' },        // Standing Charge
        'H29': { dependsOn: 'B27', requires: 'Yes' },        // Annual Spend
        
        // ENERGY USE - EXPORT TARIFF
        'H31': { dependsOn: 'B30', requires: 'Yes' },        // Export Tariff Rate
        
        // EXISTING SYSTEM
        'H36': { dependsOn: 'B35', requires: 'Yes' },        // Commissioning Date
        'H37': { dependsOn: 'B35', requires: 'Yes' },        // SEM Percentage
        
        // NEW SYSTEM - BATTERY
        'H50': { dependsOn: 'B48', requires: 'Yes' },        // Battery Extended Warranty
        'H51': { dependsOn: 'B48', requires: 'Yes' },        // Battery Replacement Cost
        
        // NEW SYSTEM - SOLAR/HYBRID INVERTER
        'H59': { dependsOn: 'B55', requires: 'Yes' },        // Inverter Extended Warranty
        'H60': { dependsOn: 'B55', requires: 'Yes' },        // Inverter Replacement Cost
        
        // NEW SYSTEM - BATTERY INVERTER
        'H66': { dependsOn: 'B62', requires: 'Yes' },        // Battery Inverter Extended Warranty
        'H67': { dependsOn: 'B62', requires: 'Yes' }         // Battery Inverter Replacement Cost
    };
    
    return dependencyMap[cellAddress] || null;
}

function checkDependenciesMet(dependencies, sheet) {
    if (!dependencies) return true; // No dependencies means always enabled
    
    const controlCell = sheet[dependencies.dependsOn];
    if (!controlCell) return false;
    
    const controlValue = controlCell.v;
    return controlValue === dependencies.requires;
}

function getReasonForStatus(analysis, dependenciesMet) {
    if (analysis.hasFormula) return 'Has formula (calculated field)';
    if (analysis.isLocked) return 'Cell is locked';
    if (!dependenciesMet) return 'Dependencies not met (radio button not selected)';
    if (!analysis.hasValue) return 'No value';
    return 'Enabled';
}

function enhancedExcelInputs() {
    console.log('🔍 Enhanced Excel Inputs Analysis');
    console.log('==================================');
    
    try {
        // Check if file exists
        if (!fs.existsSync(EXCEL_FILE_PATH)) {
            console.error(`❌ Excel file not found: ${EXCEL_FILE_PATH}`);
            process.exit(1);
        }
        
        console.log(`📁 Found Excel file: ${EXCEL_FILE_PATH}`);
        
        // Read the Excel file
        console.log('📖 Reading Excel file...');
        const workbook = XLSX.readFile(EXCEL_FILE_PATH, { 
            password: PASSWORD,
            cellStyles: true,
            cellDates: true,
            cellFormula: true
        });
        
        console.log('✅ Excel file read successfully');
        
        // Get the Inputs worksheet
        const inputsSheet = workbook.Sheets['Inputs'];
        if (!inputsSheet) {
            console.error('❌ Inputs sheet not found');
            process.exit(1);
        }
        
        console.log('✅ Found Inputs sheet');
        
        // Define input fields to check
        const inputFields = [
            // Customer Details (always enabled)
            { id: 'customer_name', label: 'Customer Name', cell: 'H12', type: 'text', required: true, section: 'Customer Details' },
            { id: 'address', label: 'Address', cell: 'H13', type: 'text', required: true, section: 'Customer Details' },
            { id: 'postcode', label: 'Postcode', cell: 'H14', type: 'text', required: false, section: 'Customer Details' },
            
            // ENERGY USE - CURRENT ELECTRICITY TARIFF
            { id: 'single_day_rate', label: 'Single / Day Rate (pence per kWh)', cell: 'H20', type: 'number', required: true, section: 'Energy Use' },
            { id: 'night_rate', label: 'Night Rate (pence per kWh)', cell: 'H21', type: 'number', required: false, section: 'Energy Use' },
            { id: 'off_peak_hours', label: 'No. of Off-Peak Hours', cell: 'H22', type: 'number', required: false, section: 'Energy Use' },
            
            // ENERGY USE - NEW ELECTRICITY TARIFF
            { id: 'new_day_rate', label: 'Day Rate (pence per kWh)', cell: 'H24', type: 'number', required: false, section: 'Energy Use' },
            { id: 'new_night_rate', label: 'Night Rate (pence per kWh)', cell: 'H25', type: 'number', required: false, section: 'Energy Use' },
            
            // ENERGY USE - ELECTRICITY CONSUMPTION
            { id: 'annual_usage', label: 'Estimated Annual Usage (kWh)', cell: 'H27', type: 'number', required: false, section: 'Energy Use' },
            { id: 'standing_charge', label: 'Standing Charge (pence per day)', cell: 'H28', type: 'number', required: false, section: 'Energy Use' },
            { id: 'annual_spend', label: 'Annual Spend (£)', cell: 'H29', type: 'number', required: false, section: 'Energy Use' },
            
            // ENERGY USE - EXPORT TARIFF
            { id: 'export_tariff_rate', label: 'Export Tariff Rate (pence per kWh)', cell: 'H31', type: 'number', required: false, section: 'Energy Use' },
            
            // EXISTING SYSTEM
            { id: 'commissioning_date', label: 'Approximate Commissioning Date', cell: 'H36', type: 'text', required: false, section: 'Existing System' },
            { id: 'sem_percentage', label: 'Percentage of above SEM used to quote self-consumption savings', cell: 'H37', type: 'number', required: false, section: 'Existing System' },
            
            // NEW SYSTEM - SOLAR
            { id: 'panel_manufacturer', label: 'Panel Manufacturer', cell: 'H42', type: 'text', required: false, section: 'New System' },
            { id: 'panel_model', label: 'Panel Model', cell: 'H43', type: 'text', required: false, section: 'New System' },
            { id: 'no_of_arrays', label: 'No. of Arrays', cell: 'H44', type: 'number', required: false, section: 'New System' },
            
            // NEW SYSTEM - BATTERY
            { id: 'battery_manufacturer', label: 'Battery Manufacturer', cell: 'H48', type: 'text', required: false, section: 'New System' },
            { id: 'battery_model', label: 'Battery Model', cell: 'H49', type: 'text', required: false, section: 'New System' },
            { id: 'battery_extended_warranty', label: 'Battery Extended Warranty Period (years)', cell: 'H50', type: 'number', required: false, section: 'New System' },
            { id: 'battery_replacement_cost', label: 'Battery Replacement Cost (£)', cell: 'H51', type: 'number', required: false, section: 'New System' },
            
            // NEW SYSTEM - SOLAR/HYBRID INVERTER
            { id: 'solar_inverter_manufacturer', label: 'Solar/Hybrid Inverter Manufacturer', cell: 'H55', type: 'text', required: false, section: 'New System' },
            { id: 'solar_inverter_model', label: 'Solar/Hybrid Inverter Model', cell: 'H56', type: 'text', required: false, section: 'New System' },
            { id: 'solar_inverter_extended_warranty', label: 'Solar Inverter Extended Warranty Period (years)', cell: 'H59', type: 'number', required: false, section: 'New System' },
            { id: 'solar_inverter_replacement_cost', label: 'Solar Inverter Replacement Cost (£)', cell: 'H60', type: 'number', required: false, section: 'New System' },
            
            // NEW SYSTEM - BATTERY INVERTER
            { id: 'battery_inverter_manufacturer', label: 'Battery Inverter Manufacturer', cell: 'H62', type: 'text', required: false, section: 'New System' },
            { id: 'battery_inverter_model', label: 'Battery Inverter Model', cell: 'H63', type: 'text', required: false, section: 'New System' },
            { id: 'battery_inverter_extended_warranty', label: 'Battery Inverter Extended Warranty Period (years)', cell: 'H66', type: 'number', required: false, section: 'New System' },
            { id: 'battery_inverter_replacement_cost', label: 'Battery Inverter Replacement Cost (£)', cell: 'H67', type: 'number', required: false, section: 'New System' }
        ];
        
        console.log(`\n🔍 Analyzing ${inputFields.length} input fields with enhanced logic...`);
        console.log('================================================================');
        
        const results = [];
        let trulyEnabledCount = 0;
        let disabledCount = 0;
        
        // Group by section for better analysis
        const sections = {};
        
        for (const field of inputFields) {
            const result = checkFieldAvailability(field, inputsSheet);
            results.push(result);
            
            if (result.isTrulyEnabled) {
                trulyEnabledCount++;
            } else {
                disabledCount++;
            }
            
            // Group by section
            if (!sections[field.section]) {
                sections[field.section] = [];
            }
            sections[field.section].push(result);
        }
        
        // Summary
        console.log('\n📊 ENHANCED ANALYSIS SUMMARY');
        console.log('============================');
        console.log(`Total fields: ${inputFields.length}`);
        console.log(`✅ Truly Enabled: ${trulyEnabledCount}`);
        console.log(`❌ Disabled: ${disabledCount}`);
        console.log(`📈 Success rate: ${((trulyEnabledCount / inputFields.length) * 100).toFixed(1)}%`);
        
        // Show results by section
        console.log('\n📋 RESULTS BY SECTION:');
        console.log('======================');
        
        Object.keys(sections).forEach(sectionName => {
            const sectionFields = sections[sectionName];
            const enabledInSection = sectionFields.filter(f => f.isTrulyEnabled).length;
            
            console.log(`\n${sectionName} (${enabledInSection}/${sectionFields.length} enabled):`);
            console.log('-'.repeat(sectionName.length + 20));
            
            sectionFields.forEach(field => {
                const status = field.isTrulyEnabled ? '✅' : '❌';
                console.log(`${status} ${field.cell}: ${field.label}`);
                if (!field.isTrulyEnabled) {
                    console.log(`   └─ Reason: ${field.reason}`);
                }
            });
        });
        
        // Show truly enabled fields
        console.log('\n✅ TRULY ENABLED FIELDS:');
        console.log('========================');
        const trulyEnabledFields = results.filter(r => r.isTrulyEnabled);
        trulyEnabledFields.forEach(field => {
            console.log(`${field.cell}: ${field.label} = "${field.value}"`);
        });
        
        // Show disabled fields with reasons
        console.log('\n❌ DISABLED FIELDS:');
        console.log('===================');
        const disabledFields = results.filter(r => !r.isTrulyEnabled);
        disabledFields.forEach(field => {
            console.log(`${field.cell}: ${field.label}`);
            console.log(`   └─ Reason: ${field.reason}`);
            if (field.radioButtonDependencies) {
                console.log(`   └─ Depends on: ${field.radioButtonDependencies.dependsOn} = "${field.radioButtonDependencies.requires}"`);
            }
        });
        
        // Export results
        const outputPath = path.join(__dirname, 'enhanced-excel-inputs-results.json');
        const outputData = {
            timestamp: new Date().toISOString(),
            excelFile: EXCEL_FILE_PATH,
            summary: {
                total: inputFields.length,
                trulyEnabled: trulyEnabledCount,
                disabled: disabledCount,
                successRate: (trulyEnabledCount / inputFields.length) * 100
            },
            results: results,
            sections: sections,
            trulyEnabledFields: trulyEnabledFields.map(f => ({
                id: f.id,
                label: f.label,
                cell: f.cell,
                type: f.type,
                value: f.value,
                section: f.section
            }))
        };
        
        fs.writeFileSync(outputPath, JSON.stringify(outputData, null, 2));
        console.log(`\n💾 Enhanced results saved to: ${outputPath}`);
        
        return outputData;
        
    } catch (error) {
        console.error('❌ Error in enhanced Excel inputs analysis:', error);
        process.exit(1);
    }
}

// Run the enhanced analysis
if (require.main === module) {
    enhancedExcelInputs();
}

module.exports = { enhancedExcelInputs };
