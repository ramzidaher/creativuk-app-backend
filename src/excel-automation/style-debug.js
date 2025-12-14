#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '..', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
const PASSWORD = '99';

function debugCellStyles() {
    console.log('🔍 DEBUGGING Cell Styles & Background Colors');
    console.log('===========================================');
    
    try {
        // Check if file exists
        if (!fs.existsSync(EXCEL_FILE_PATH)) {
            throw new Error(`Excel file not found: ${EXCEL_FILE_PATH}`);
        }
        
        console.log(`📁 Found Excel file: ${EXCEL_FILE_PATH}`);
        
        // Read the Excel file with ALL style options
        console.log('📖 Reading Excel file with style debugging...');
        const workbook = XLSX.readFile(EXCEL_FILE_PATH, { 
            password: PASSWORD,
            cellStyles: true,
            cellDates: true,
            cellFormula: true,
            cellNF: true,
            cellHTML: true,
            cellText: true
        });
        
        console.log('✅ Excel file read successfully');
        
        // Get the Inputs worksheet
        const inputsSheet = workbook.Sheets['Inputs'];
        if (!inputsSheet) {
            throw new Error('Inputs sheet not found');
        }
        
        console.log('✅ Found Inputs sheet');
        
        // Sample cells to check (based on the Excel images)
        const sampleCells = [
            'H12', 'H13', 'H14',  // Customer Details (should be yellow)
            'H19', 'H20', 'H21',  // Energy Use (should be yellow)
            'H23', 'H24',         // New Electricity Tariff (should be yellow)
            'H26', 'H27', 'H28',  // Electricity Consumption (should be yellow)
            'H30',                // Export Tariff (should be yellow)
            'H34', 'H35', 'H36',  // Existing System (should be light green)
            'H41', 'H42', 'H43',  // New System - Solar (should be light green)
            'H45', 'H46',         // New System - Battery (should be light green)
            'H48', 'H49',         // Battery Warranty (should be light green)
            'H51', 'H52',         // Solar/Hybrid Inverter (should be light green)
            'H55', 'H56',         // Solar Inverter Warranty (should be light green)
            'H58', 'H59',         // Battery Inverter (should be light green)
            'H62', 'H63'          // Battery Inverter Warranty (should be light green)
        ];
        
        console.log('\n🔍 DETAILED CELL STYLE ANALYSIS:');
        console.log('==================================');
        
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
            
            // Check if cell has style information
            if (cell.s) {
                console.log(`  └─ Has style: YES`);
                console.log(`  └─ Style object keys: ${Object.keys(cell.s).join(', ')}`);
                
                // Check fill/background
                if (cell.s.fill) {
                    console.log(`  └─ Fill object keys: ${Object.keys(cell.s.fill).join(', ')}`);
                    
                    if (cell.s.fill.fgColor) {
                        console.log(`  └─ Foreground color keys: ${Object.keys(cell.s.fill.fgColor).join(', ')}`);
                        console.log(`  └─ Foreground color RGB: ${cell.s.fill.fgColor.rgb || 'None'}`);
                        console.log(`  └─ Foreground color theme: ${cell.s.fill.fgColor.theme || 'None'}`);
                        console.log(`  └─ Foreground color tint: ${cell.s.fill.fgColor.tint || 'None'}`);
                    }
                    
                    if (cell.s.fill.bgColor) {
                        console.log(`  └─ Background color keys: ${Object.keys(cell.s.fill.bgColor).join(', ')}`);
                        console.log(`  └─ Background color RGB: ${cell.s.fill.bgColor.rgb || 'None'}`);
                        console.log(`  └─ Background color theme: ${cell.s.fill.bgColor.theme || 'None'}`);
                        console.log(`  └─ Background color tint: ${cell.s.fill.bgColor.tint || 'None'}`);
                    }
                } else {
                    console.log(`  └─ No fill information`);
                }
                
                // Check font
                if (cell.s.font) {
                    console.log(`  └─ Font object keys: ${Object.keys(cell.s.font).join(', ')}`);
                    
                    if (cell.s.font.color) {
                        console.log(`  └─ Font color RGB: ${cell.s.font.color.rgb || 'None'}`);
                    }
                }
                
                // Check alignment
                if (cell.s.alignment) {
                    console.log(`  └─ Alignment object keys: ${Object.keys(cell.s.alignment).join(', ')}`);
                }
                
                // Check border
                if (cell.s.border) {
                    console.log(`  └─ Border object keys: ${Object.keys(cell.s.border).join(', ')}`);
                }
                
                // Check protection
                if (cell.s.protection) {
                    console.log(`  └─ Protection object keys: ${Object.keys(cell.s.protection).join(', ')}`);
                }
                
                // Dump the entire style object for debugging
                console.log(`  └─ Full style object:`, JSON.stringify(cell.s, null, 4));
                
            } else {
                console.log(`  └─ No style information`);
            }
            
            // Check for any other properties
            const otherProps = Object.keys(cell).filter(key => !['v', 't', 'f', 'l', 's'].includes(key));
            if (otherProps.length > 0) {
                console.log(`  └─ Other properties:`, otherProps);
            }
        });
        
        // Check for any cells that might have different style structures
        console.log('\n🔍 LOOKING FOR CELLS WITH STYLES:');
        console.log('==================================');
        
        const sheetRange = XLSX.utils.decode_range(inputsSheet['!ref']);
        let cellsWithStyles = 0;
        let cellsWithFill = 0;
        let cellsWithColor = 0;
        
        for (let row = sheetRange.s.r; row <= Math.min(sheetRange.e.r, sheetRange.s.r + 100); row++) {
            for (let col = sheetRange.s.c; col <= Math.min(sheetRange.e.c, sheetRange.s.c + 20); col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = inputsSheet[cellAddress];
                
                if (cell && cell.s) {
                    cellsWithStyles++;
                    
                    if (cell.s.fill) {
                        cellsWithFill++;
                        
                        if (cell.s.fill.fgColor && cell.s.fill.fgColor.rgb) {
                            cellsWithColor++;
                            console.log(`${cellAddress}: RGB ${cell.s.fill.fgColor.rgb}`);
                        }
                    }
                }
            }
        }
        
        console.log(`\n📊 STYLE STATISTICS:`);
        console.log(`Cells with styles: ${cellsWithStyles}`);
        console.log(`Cells with fill: ${cellsWithFill}`);
        console.log(`Cells with color: ${cellsWithColor}`);
        
        console.log('\n💾 Style debug complete!');
        
    } catch (error) {
        console.error('❌ Error debugging cell styles:', error);
    }
}

// Run the debug function
if (require.main === module) {
    debugCellStyles();
}

module.exports = { debugCellStyles };
