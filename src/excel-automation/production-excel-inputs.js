#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '..', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
const PASSWORD = '99';

// Radio button control cells and their possible values
const RADIO_BUTTON_CONTROLS = {
    'B20': { name: 'Electricity Tariff Type', options: ['SingleRate', 'DualRate'] },
    'B23': { name: 'New Electricity Tariff', options: ['Yes', 'No'] },
    'B27': { name: 'Know Annual Usage', options: ['Yes', 'No'] },
    'B30': { name: 'Export Tariff', options: ['Yes', 'No'] },
    'B35': { name: 'Existing Solar Installation', options: ['Yes', 'No'] },
    'B48': { name: 'Battery Extended Warranty', options: ['Yes', 'No'] },
    'B55': { name: 'Solar Inverter Extended Warranty', options: ['Yes', 'No'] },
    'B62': { name: 'Battery Inverter Extended Warranty', options: ['Yes', 'No'] }
};

// Field dependencies based on radio button selections
const FIELD_DEPENDENCIES = {
    // ENERGY USE - CURRENT ELECTRICITY TARIFF
    'H20': { dependsOn: 'B20', requires: 'SingleRate', label: 'Single / Day Rate (pence per kWh)' },
    'H21': { dependsOn: 'B20', requires: 'DualRate', label: 'Night Rate (pence per kWh)' },
    'H22': { dependsOn: 'B20', requires: 'DualRate', label: 'No. of Off-Peak Hours' },
    
    // ENERGY USE - NEW ELECTRICITY TARIFF
    'H24': { dependsOn: 'B23', requires: 'Yes', label: 'Day Rate (pence per kWh)' },
    'H25': { dependsOn: 'B23', requires: 'Yes', label: 'Night Rate (pence per kWh)' },
    
    // ENERGY USE - ELECTRICITY CONSUMPTION
    'H27': { dependsOn: 'B27', requires: 'Yes', label: 'Estimated Annual Usage (kWh)' },
    'H28': { dependsOn: 'B27', requires: 'Yes', label: 'Standing Charge (pence per day)' },
    'H29': { dependsOn: 'B27', requires: 'Yes', label: 'Annual Spend (£)' },
    
    // ENERGY USE - EXPORT TARIFF
    'H31': { dependsOn: 'B30', requires: 'Yes', label: 'Export Tariff Rate (pence per kWh)' },
    
    // EXISTING SYSTEM
    'H36': { dependsOn: 'B35', requires: 'Yes', label: 'Approximate Commissioning Date' },
    'H37': { dependsOn: 'B35', requires: 'Yes', label: 'Percentage of above SEM used to quote self-consumption savings' },
    
    // NEW SYSTEM - BATTERY
    'H50': { dependsOn: 'B48', requires: 'Yes', label: 'Battery Extended Warranty Period (years)' },
    'H51': { dependsOn: 'B48', requires: 'Yes', label: 'Battery Replacement Cost (£)' },
    
    // NEW SYSTEM - SOLAR/HYBRID INVERTER
    'H59': { dependsOn: 'B55', requires: 'Yes', label: 'Solar Inverter Extended Warranty Period (years)' },
    'H60': { dependsOn: 'B55', requires: 'Yes', label: 'Solar Inverter Replacement Cost (£)' },
    
    // NEW SYSTEM - BATTERY INVERTER
    'H66': { dependsOn: 'B62', requires: 'Yes', label: 'Battery Inverter Extended Warranty Period (years)' },
    'H67': { dependsOn: 'B62', requires: 'Yes', label: 'Battery Inverter Replacement Cost (£)' }
};

// Always enabled fields (no dependencies)
const ALWAYS_ENABLED_FIELDS = [
    { id: 'customer_name', cell: 'H12', label: 'Customer Name', type: 'text', required: true, section: 'Customer Details' },
    { id: 'address', cell: 'H13', label: 'Address', type: 'text', required: true, section: 'Customer Details' },
    { id: 'postcode', cell: 'H14', label: 'Postcode', type: 'text', required: false, section: 'Customer Details' },
    
    // NEW SYSTEM - SOLAR (always enabled)
    { id: 'panel_manufacturer', cell: 'H42', label: 'Panel Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'panel_model', cell: 'H43', label: 'Panel Model', type: 'text', required: false, section: 'New System' },
    { id: 'no_of_arrays', cell: 'H44', label: 'No. of Arrays', type: 'number', required: false, section: 'New System' },
    
    // NEW SYSTEM - BATTERY (always enabled)
    { id: 'battery_manufacturer', cell: 'H48', label: 'Battery Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'battery_model', cell: 'H49', label: 'Battery Model', type: 'text', required: false, section: 'New System' },
    
    // NEW SYSTEM - SOLAR/HYBRID INVERTER (always enabled)
    { id: 'solar_inverter_manufacturer', cell: 'H55', label: 'Solar/Hybrid Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'solar_inverter_model', cell: 'H56', label: 'Solar/Hybrid Inverter Model', type: 'text', required: false, section: 'New System' },
    
    // NEW SYSTEM - BATTERY INVERTER (always enabled)
    { id: 'battery_inverter_manufacturer', cell: 'H62', label: 'Battery Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
    { id: 'battery_inverter_model', cell: 'H63', label: 'Battery Inverter Model', type: 'text', required: false, section: 'New System' }
];

function getRadioButtonStates(sheet) {
    const states = {};
    
    for (const [cellAddress, control] of Object.entries(RADIO_BUTTON_CONTROLS)) {
        const cell = sheet[cellAddress];
        if (cell && cell.v !== undefined && cell.v !== null) {
            states[cellAddress] = cell.v;
        } else {
            states[cellAddress] = null; // No selection made
        }
    }
    
    return states;
}

function checkFieldAvailability(field, sheet, radioButtonStates) {
    const cell = sheet[field.cell];
    
    // Check if cell exists (regardless of value)
    if (!cell) {
        return {
            ...field,
            enabled: false,
            reason: 'Cell not found',
            value: null
        };
    }
    
    // Check if cell has formula (calculated field - should be disabled)
    if (cell.f) {
        return {
            ...field,
            enabled: false,
            reason: 'Has formula (calculated field)',
            value: cell.v
        };
    }
    
    // Check if cell is locked
    if (cell.l && cell.l.locked) {
        return {
            ...field,
            enabled: false,
            reason: 'Cell is locked',
            value: cell.v
        };
    }
    
    // Check dependencies
    const dependency = FIELD_DEPENDENCIES[field.cell];
    if (dependency) {
        const controlValue = radioButtonStates[dependency.dependsOn];
        const dependenciesMet = controlValue === dependency.requires;
        
        return {
            ...field,
            enabled: dependenciesMet,
            reason: dependenciesMet ? 'Enabled' : `Depends on ${dependency.dependsOn} = "${dependency.requires}"`,
            value: cell.v || '', // Use empty string if no value
            dependency: dependency
        };
    }
    
    // No dependencies - always enabled (regardless of current value)
    return {
        ...field,
        enabled: true,
        reason: 'No dependencies',
        value: cell.v || '' // Use empty string if no value
    };
}

function getDynamicInputs(excelFilePath = EXCEL_FILE_PATH) {
    console.log('🔍 Getting Dynamic Inputs from Excel');
    console.log('====================================');
    
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
        
        // Get current radio button states
        const radioButtonStates = getRadioButtonStates(inputsSheet);
        console.log('📻 Radio button states:', radioButtonStates);
        
        // Check all fields
        const allFields = [...ALWAYS_ENABLED_FIELDS];
        
        // Add dependent fields
        for (const [cellAddress, dependency] of Object.entries(FIELD_DEPENDENCIES)) {
            allFields.push({
                id: cellAddress.toLowerCase().replace('h', 'field_'),
                cell: cellAddress,
                label: dependency.label,
                type: 'number',
                required: false,
                section: 'Dependent Fields'
            });
        }
        
        console.log(`\n🔍 Analyzing ${allFields.length} input fields...`);
        
        const results = [];
        let enabledCount = 0;
        let disabledCount = 0;
        
        for (const field of allFields) {
            const result = checkFieldAvailability(field, inputsSheet, radioButtonStates);
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
        console.log('\n📊 SUMMARY');
        console.log('==========');
        console.log(`Total fields analyzed: ${allFields.length}`);
        console.log(`✅ Enabled fields: ${enabledCount}`);
        console.log(`❌ Disabled fields: ${disabledCount}`);
        console.log(`📈 Success rate: ${((enabledCount / allFields.length) * 100).toFixed(1)}%`);
        
        // Show enabled fields
        console.log('\n✅ ENABLED FIELDS:');
        console.log('==================');
        enabledFields.forEach(field => {
            console.log(`${field.cellReference}: ${field.label} = "${field.value}"`);
        });
        
        // Show disabled fields with reasons
        console.log('\n❌ DISABLED FIELDS:');
        console.log('===================');
        const disabledFields = results.filter(field => !field.enabled);
        disabledFields.forEach(field => {
            console.log(`${field.cell}: ${field.label}`);
            console.log(`   └─ Reason: ${field.reason}`);
        });
        
        const response = {
            success: true,
            message: `Found ${enabledCount} enabled input fields`,
            radioButtonStates,
            inputFields: enabledFields,
            summary: {
                total: allFields.length,
                enabled: enabledCount,
                disabled: disabledCount,
                successRate: (enabledCount / allFields.length) * 100
            }
        };
        
        // Export detailed results for debugging
        const outputPath = path.join(__dirname, 'production-excel-inputs-results.json');
        fs.writeFileSync(outputPath, JSON.stringify({
            timestamp: new Date().toISOString(),
            excelFile: excelFilePath,
            response,
            detailedResults: results
        }, null, 2));
        
        console.log(`\n💾 Detailed results saved to: ${outputPath}`);
        
        return response;
        
    } catch (error) {
        console.error('❌ Error getting dynamic inputs:', error);
        return {
            success: false,
            message: 'Failed to get dynamic inputs',
            error: error.message
        };
    }
}

// Run the function if called directly
if (require.main === module) {
    getDynamicInputs();
}

module.exports = { getDynamicInputs };
