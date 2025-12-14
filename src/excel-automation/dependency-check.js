#!/usr/bin/env node

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '..', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
const PASSWORD = '99';

function checkCellDependencies() {
    console.log('🔍 CHECKING Cell Dependencies & Formulas');
    console.log('========================================');
    
    try {
        // Check if file exists
        if (!fs.existsSync(EXCEL_FILE_PATH)) {
            throw new Error(`Excel file not found: ${EXCEL_FILE_PATH}`);
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
            throw new Error('Inputs sheet not found');
        }
        
        console.log('✅ Found Inputs sheet');
        
        // Sample cells to check
        const sampleCells = ['H12', 'H20', 'H21', 'H22', 'H24', 'H25', 'H27', 'H28', 'H29', 'H31', 'H36', 'H37', 'H42', 'H43', 'H44', 'H48', 'H49', 'H50', 'H51', 'H55', 'H56', 'H59', 'H60', 'H62', 'H63', 'H66', 'H67'];
        
        console.log('\n🔍 CHECKING CELL DEPENDENCIES:');
        console.log('===============================');
        
        // Check which cells reference our input cells
        const cellReferences = {};
        
        // Scan all cells in the sheet to find formulas that reference our input cells
        const sheetRange = XLSX.utils.decode_range(inputsSheet['!ref']);
        
        for (let row = sheetRange.s.r; row <= sheetRange.e.r; row++) {
            for (let col = sheetRange.s.c; col <= sheetRange.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = inputsSheet[cellAddress];
                
                if (cell && cell.f) {
                    // This cell has a formula, check if it references any of our input cells
                    const formula = cell.f;
                    
                    sampleCells.forEach(inputCell => {
                        if (formula.includes(inputCell)) {
                            if (!cellReferences[inputCell]) {
                                cellReferences[inputCell] = [];
                            }
                            cellReferences[inputCell].push({
                                cell: cellAddress,
                                formula: formula
                            });
                        }
                    });
                }
            }
        }
        
        // Now check each input cell
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
            
            // Check if this cell is referenced by other formulas
            if (cellReferences[cellAddress]) {
                console.log(`  └─ Referenced by ${cellReferences[cellAddress].length} other cells:`);
                cellReferences[cellAddress].forEach(ref => {
                    console.log(`     └─ ${ref.cell}: ${ref.formula}`);
                });
            } else {
                console.log(`  └─ Not referenced by any formulas`);
            }
            
            // Check if this cell has dependencies (formulas that reference other cells)
            if (cell.f) {
                console.log(`  └─ Has formula that depends on other cells`);
            }
            
            // Check if this cell is a calculated field
            if (cell.t === 'z' && cell.v === 'undefined') {
                console.log(`  └─ APPEARS TO BE A CALCULATED/DEPENDENT FIELD`);
            }
        });
        
        // Check for cells that might be controlled by radio buttons or other inputs
        console.log('\n🎛️ CHECKING FOR CONTROL CELLS:');
        console.log('==============================');
        
        // Look for cells in column B (where radio buttons usually are)
        const controlCells = ['B20', 'B23', 'B27', 'B30', 'B35', 'B48', 'B55', 'B62'];
        
        controlCells.forEach(cellAddress => {
            const cell = inputsSheet[cellAddress];
            console.log(`\n${cellAddress}:`);
            
            if (!cell) {
                console.log('  └─ Cell not found');
                return;
            }
            
            console.log(`  └─ Value: "${cell.v}"`);
            console.log(`  └─ Type: ${cell.t}`);
            console.log(`  └─ Formula: ${cell.f || 'None'}`);
            
            // Check if this control cell is referenced by formulas
            if (cellReferences[cellAddress]) {
                console.log(`  └─ Controls ${cellReferences[cellAddress].length} other cells:`);
                cellReferences[cellAddress].forEach(ref => {
                    console.log(`     └─ ${ref.cell}: ${ref.formula}`);
                });
            }
        });
        
        // Check for any cells that might have validation or error handling
        console.log('\n🚫 CHECKING FOR ERROR HANDLING:');
        console.log('===============================');
        
        // Look for cells with error values or special handling
        sampleCells.forEach(cellAddress => {
            const cell = inputsSheet[cellAddress];
            if (cell && cell.v === 'undefined') {
                console.log(`${cellAddress}: Has "undefined" value - likely controlled by other cells`);
                
                // Check if this cell is in a range that might be controlled
                const row = parseInt(cellAddress.substring(1));
                const col = cellAddress.charCodeAt(0) - 65; // Convert A=0, B=1, etc.
                
                // Check if there are any control cells in the same row
                for (let c = 0; c < col; c++) {
                    const controlCell = XLSX.utils.encode_cell({ r: row - 1, c: c }); // Excel is 1-indexed
                    const controlCellData = inputsSheet[controlCell];
                    if (controlCellData && controlCellData.v) {
                        console.log(`  └─ Possible control cell ${controlCell}: "${controlCellData.v}"`);
                    }
                }
            }
        });
        
        console.log('\n💾 Dependency check complete!');
        
    } catch (error) {
        console.error('❌ Error checking dependencies:', error);
    }
}

// Run the function
if (require.main === module) {
    checkCellDependencies();
}

module.exports = { checkCellDependencies };
