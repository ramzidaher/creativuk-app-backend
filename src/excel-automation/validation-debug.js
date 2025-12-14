#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '..', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
const PASSWORD = '99';

function debugValidationAndErrors() {
    console.log('🔍 DEBUGGING Data Validation & Error Popups');
    console.log('===========================================');
    
    try {
        // Check if file exists
        if (!fs.existsSync(EXCEL_FILE_PATH)) {
            throw new Error(`Excel file not found: ${EXCEL_FILE_PATH}`);
        }
        
        console.log(`📁 Found Excel file: ${EXCEL_FILE_PATH}`);
        
        // Read the Excel file with ALL possible options to get validation data
        console.log('📖 Reading Excel file with validation data...');
        const workbook = XLSX.readFile(EXCEL_FILE_PATH, { 
            password: PASSWORD,
            cellStyles: true,
            cellDates: true,
            cellFormula: true,
            cellNF: true,
            cellHTML: true,
            cellText: true,
            cellFormula: true,
            cellStyles: true,
            cellDates: true,
            cellNF: true,
            cellHTML: true,
            cellText: true,
            cellFormula: true,
            cellStyles: true,
            cellDates: true,
            cellNF: true,
            cellHTML: true,
            cellText: true,
            // Try to get validation data
            cellFormula: true,
            cellStyles: true,
            cellDates: true,
            cellNF: true,
            cellHTML: true,
            cellText: true
        });
        
        console.log('✅ Excel file read successfully');
        console.log(`📋 Sheets found: ${workbook.SheetNames.join(', ')}`);
        
        // Get the Inputs worksheet
        const inputsSheet = workbook.Sheets['Inputs'];
        if (!inputsSheet) {
            throw new Error('Inputs sheet not found');
        }
        
        console.log('\n✅ Found Inputs sheet');
        console.log('Sheet properties:', inputsSheet['!ref']);
        
        // Check for data validation rules
        console.log('\n✅ DATA VALIDATION RULES:');
        console.log('==========================');
        if (inputsSheet['!dataValidation']) {
            console.log('Data validation rules found:', inputsSheet['!dataValidation']);
            
            // Parse validation rules
            inputsSheet['!dataValidation'].forEach(rule => {
                console.log(`\nValidation Rule:`);
                console.log(`  └─ Range: ${rule.sqref}`);
                console.log(`  └─ Type: ${rule.type}`);
                console.log(`  └─ Operator: ${rule.operator}`);
                console.log(`  └─ Formula1: ${rule.formula1}`);
                console.log(`  └─ Formula2: ${rule.formula2}`);
                console.log(`  └─ ShowErrorMessage: ${rule.showErrorMessage}`);
                console.log(`  └─ ErrorTitle: ${rule.errorTitle}`);
                console.log(`  └─ Error: ${rule.error}`);
                console.log(`  └─ ShowInputMessage: ${rule.showInputMessage}`);
                console.log(`  └─ PromptTitle: ${rule.promptTitle}`);
                console.log(`  └─ Prompt: ${rule.prompt}`);
            });
        } else {
            console.log('No data validation rules found in sheet');
        }
        
        // Check for workbook-level validation
        console.log('\n📋 WORKBOOK VALIDATION:');
        console.log('========================');
        if (workbook.Workbook && workbook.Workbook.Names) {
            console.log('Named ranges (might be used for validation):');
            workbook.Workbook.Names.forEach(name => {
                if (name.Name.includes('Validation') || name.Name.includes('List') || name.Name.includes('Dropdown')) {
                    console.log(`  └─ ${name.Name}: ${name.Ref}`);
                }
            });
        }
        
        // Check for conditional formatting that might disable cells
        console.log('\n🎨 CONDITIONAL FORMATTING:');
        console.log('==========================');
        if (inputsSheet['!cf']) {
            console.log('Conditional formatting rules found:', inputsSheet['!cf']);
            
            inputsSheet['!cf'].forEach(cf => {
                console.log(`\nConditional Format Rule:`);
                console.log(`  └─ Range: ${cf.ref}`);
                console.log(`  └─ Type: ${cf.type}`);
                console.log(`  └─ Priority: ${cf.priority}`);
                console.log(`  └─ StopIfTrue: ${cf.stopIfTrue}`);
                if (cf.cfvo) {
                    cf.cfvo.forEach(vo => {
                        console.log(`  └─ CFVO: ${vo.type} = ${vo.val}`);
                    });
                }
            });
        } else {
            console.log('No conditional formatting found');
        }
        
        // Check for sheet protection
        console.log('\n🔒 SHEET PROTECTION:');
        console.log('=====================');
        if (inputsSheet['!protect']) {
            console.log('Sheet protection found:', inputsSheet['!protect']);
        } else {
            console.log('No sheet protection found');
        }
        
        // Check for workbook protection
        console.log('\n🔒 WORKBOOK PROTECTION:');
        console.log('=======================');
        if (workbook.Workbook && workbook.Workbook.WorkbookPr) {
            console.log('Workbook properties:', workbook.Workbook.WorkbookPr);
        }
        
        // Sample cells to check for validation
        const sampleCells = ['H12', 'H20', 'H21', 'H22', 'H24', 'H25', 'H27', 'H28', 'H29', 'H31', 'H36', 'H37', 'H42', 'H43', 'H44', 'H48', 'H49', 'H50', 'H51', 'H55', 'H56', 'H59', 'H60', 'H62', 'H63', 'H66', 'H67'];
        
        console.log('\n🔍 CELL-BY-CELL VALIDATION CHECK:');
        console.log('===================================');
        
        sampleCells.forEach(cellAddress => {
            const cell = inputsSheet[cellAddress];
            console.log(`\n${cellAddress}:`);
            
            if (!cell) {
                console.log('  └─ Cell not found');
                return;
            }
            
            console.log(`  └─ Value: "${cell.v}"`);
            console.log(`  └─ Type: ${cell.t}`);
            console.log(`  └─ Formula: ${cell.f || 'None'}`);
            
            // Check for validation-related properties
            if (cell.v) {
                console.log(`  └─ Has validation value: ${!!cell.v}`);
            }
            
            // Check for any validation-related properties
            const validationProps = Object.keys(cell).filter(key => 
                key.includes('validation') || 
                key.includes('list') || 
                key.includes('dropdown') ||
                key.includes('error') ||
                key.includes('prompt')
            );
            
            if (validationProps.length > 0) {
                console.log(`  └─ Validation properties:`, validationProps);
            }
            
            // Check for any other unusual properties
            const unusualProps = Object.keys(cell).filter(key => 
                !['v', 't', 'f', 'l', 's', 'r', 'h', 'z', 'w'].includes(key)
            );
            
            if (unusualProps.length > 0) {
                console.log(`  └─ Unusual properties:`, unusualProps);
            }
        });
        
        // Look for cells that might have validation errors
        console.log('\n🚫 POTENTIAL VALIDATION ERROR CELLS:');
        console.log('=====================================');
        
        sampleCells.forEach(cellAddress => {
            const cell = inputsSheet[cellAddress];
            if (cell) {
                // Check for cells that might cause validation errors
                if (cell.v === 'undefined' || cell.v === '#N/A' || cell.v === '#VALUE!' || cell.v === '#REF!') {
                    console.log(`${cellAddress}: Potential validation error - "${cell.v}"`);
                }
                
                // Check for cells with specific patterns that might indicate validation issues
                if (typeof cell.v === 'string' && (
                    cell.v.includes('error') || 
                    cell.v.includes('invalid') || 
                    cell.v.includes('not allowed') ||
                    cell.v.includes('validation')
                )) {
                    console.log(`${cellAddress}: Validation-related text - "${cell.v}"`);
                }
            }
        });
        
        // Check for any VBA or macro references that might cause popups
        console.log('\n🔧 VBA/MACRO REFERENCES:');
        console.log('=========================');
        if (workbook.Workbook && workbook.Workbook.Names) {
            const vbaNames = workbook.Workbook.Names.filter(name => 
                name.Name.includes('VBA') || 
                name.Name.includes('Macro') || 
                name.Name.includes('Function') ||
                name.Name.includes('Sub')
            );
            
            if (vbaNames.length > 0) {
                console.log('VBA/Macro references found:');
                vbaNames.forEach(name => {
                    console.log(`  └─ ${name.Name}: ${name.Ref}`);
                });
            } else {
                console.log('No VBA/Macro references found');
            }
        }
        
        console.log('\n💾 Validation debug complete!');
        
    } catch (error) {
        console.error('❌ Error debugging validation:', error);
    }
}

// Run the debug function
if (require.main === module) {
    debugValidationAndErrors();
}

module.exports = { debugValidationAndErrors };
