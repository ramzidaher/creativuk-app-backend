import { Injectable, Logger } from '@nestjs/common';
import { spawn } from 'child_process';
import * as path from 'path';
import * as fs from 'fs';

@Injectable()
export class ExcelCellDetectorService {
  private readonly logger = new Logger(ExcelCellDetectorService.name);

  private readonly TEMPLATE_FILE_PATH = path.join(process.cwd(), 'src', 'excel-file-calculator', 'Off peak V2.1 Eon SEG-cleared.xlsm');
  private readonly OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');

  /**
   * Find the latest version of an existing opportunity file
   */
  private findLatestOpportunityFile(opportunityId: string): string | null {
    if (!fs.existsSync(this.OPPORTUNITIES_FOLDER)) {
      return null;
    }

    const files = fs.readdirSync(this.OPPORTUNITIES_FOLDER);
    const baseFileName = `Off peak V2.1 Eon SEG-${opportunityId}`;
    const basePattern = baseFileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const versionRegex = new RegExp(`^${basePattern}-v(\\d+)\\.xlsm$`);
    
    let latestFile: string | null = null;
    let maxVersion = 0;
    
    for (const file of files) {
      const match = file.match(versionRegex);
      if (match) {
        const version = parseInt(match[1], 10);
        if (version > maxVersion) {
          maxVersion = version;
          latestFile = file;
        }
      }
    }

    return latestFile ? path.join(this.OPPORTUNITIES_FOLDER, latestFile) : null;
  }

  async getEnabledInputFields(opportunityId?: string): Promise<{ success: boolean; message: string; error?: string; inputFields?: any[] }> {
    this.logger.log(`Getting enabled input fields${opportunityId ? ` for opportunity: ${opportunityId}` : ''}`);

    try {
      // Determine which file to use
      let excelFilePath: string;
      if (opportunityId) {
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using opportunity file: ${excelFilePath}`);
        } else {
          const error = `Opportunity file not found: ${opportunityFilePath}. Please ensure radio button selections have been applied first.`;
          this.logger.error(error);
          return { success: false, message: error };
        }
      } else {
        excelFilePath = this.TEMPLATE_FILE_PATH;
        this.logger.log(`Using template file: ${excelFilePath}`);
      }

      // Check if file exists
      if (!fs.existsSync(excelFilePath)) {
        const error = `Excel file not found at: ${excelFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Check if we're on Windows (required for Excel COM automation)
      if (process.platform !== 'win32') {
        const error = 'Excel automation requires Windows platform';
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Create PowerShell script to check enabled cells
      const psScript = this.createCheckEnabledCellsScript(excelFilePath);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-check-cells-${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created temporary PowerShell script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Failed to cleanup temporary script: ${cleanupError.message}`);
      }

      if (result.success) {
        // Parse the PowerShell output to extract enabled field information
        const enabledFields = this.parseEnabledFieldsFromOutput(result.output || '');
        
        this.logger.log(`Successfully retrieved ${enabledFields.length} enabled input fields`);
        return {
          success: true,
          message: `Successfully retrieved ${enabledFields.length} enabled input fields`,
          inputFields: enabledFields
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to get enabled input fields`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in getEnabledInputFields: ${error.message}`);
      return {
        success: false,
        message: `Error getting enabled input fields`,
        error: error.message
      };
    }
  }

  private createCheckEnabledCellsScript(excelFilePath: string): string {
    return `
# PowerShell script to check which input cells are enabled in the Excel sheet
param(
    [string]$ExcelFilePath
)

try {
    Write-Host "Opening Excel file: $ExcelFilePath" -ForegroundColor Green
    
    # Create Excel application object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open the workbook
    $workbook = $excel.Workbooks.Open($ExcelFilePath)
    $worksheet = $workbook.Worksheets.Item(1)
    
    Write-Host "Checking which input cells are enabled..." -ForegroundColor Green
    
    # Define all possible input fields with their cell references
    $allFields = @(
        @{ id = "single_day_rate"; label = "Single / Day Rate (pence per kWh)"; type = "number"; cellReference = "H20" },
        @{ id = "night_rate"; label = "Night Rate (pence per kWh)"; type = "number"; cellReference = "H21" },
        @{ id = "off_peak_hours"; label = "No. of Off-Peak Hours"; type = "number"; cellReference = "H22" },
        @{ id = "new_day_rate"; label = "Day Rate (pence per kWh)"; type = "number"; cellReference = "H24" },
        @{ id = "new_night_rate"; label = "Night Rate (pence per kWh)"; type = "number"; cellReference = "H25" },
        @{ id = "annual_usage"; label = "Estimated Annual Usage (kWh)"; type = "number"; cellReference = "H27" },
        @{ id = "standing_charge"; label = "Standing Charge (pence per day)"; type = "number"; cellReference = "H28" },
        @{ id = "annual_spend"; label = "Annual Spend (£)"; type = "number"; cellReference = "H29" },
        @{ id = "export_tariff_rate"; label = "Export Tariff Rate (pence per kWh)"; type = "number"; cellReference = "H31" },
        @{ id = "existing_sem"; label = "Existing SEM"; type = "text"; cellReference = "H36" },
        @{ id = "commissioning_date"; label = "Approximate Commissioning Date"; type = "text"; cellReference = "H37" },
        @{ id = "sem_percentage"; label = "Percentage of above SEM used to quote self-consumption savings"; type = "number"; cellReference = "H38" },
        @{ id = "panel_manufacturer"; label = "Panel Manufacturer"; type = "text"; cellReference = "H42" },
        @{ id = "panel_model"; label = "Panel Model"; type = "text"; cellReference = "H43" },
        @{ id = "no_of_arrays"; label = "No. of Arrays"; type = "number"; cellReference = "H44" }
    )
    
    $enabledFields = @()
    
    # Check each field to see if it's enabled
    foreach ($field in $allFields) {
        $cell = $worksheet.Range($field.cellReference)
        
        # Check if cell is unlocked and not protected
        $isLocked = $cell.Locked
        $isWorksheetProtected = $worksheet.ProtectContents
        
        # A cell is enabled if it's unlocked AND the worksheet is not protected
        # OR if the worksheet is protected but the cell is unlocked
        $isEnabled = (-not $isLocked) -or (-not $isWorksheetProtected)
        
        Write-Host "Cell $($field.cellReference) ($($field.label)): Locked=$isLocked, Protected=$isWorksheetProtected, Enabled=$isEnabled" -ForegroundColor Cyan
        
        if ($isEnabled) {
            $enabledFields += @{
                id = $field.id
                label = $field.label
                type = $field.type
                cellReference = $field.cellReference
                required = $false
                enabled = $true
                value = ""
            }
        }
    }
    
    Write-Host "Found $($enabledFields.Count) enabled fields" -ForegroundColor Green
    
    # Convert to JSON and output
    $result = @{
        success = $true
        enabledFields = $enabledFields
    }
    
    # Output JSON result
    $jsonResult = $result | ConvertTo-Json -Depth 3 -Compress
    Write-Host "RESULT: $jsonResult"
    
    # Close Excel
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    exit 0
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    
    # Try to close Excel if it's still open
    try {
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {
        Write-Host "Warning: Could not properly close Excel: $_" -ForegroundColor Yellow
    }
    
    exit 1
}
`;
  }

  private async runPowerShellScript(scriptPath: string): Promise<{ success: boolean; output?: string; error?: string }> {
    return new Promise((resolve) => {
      const powershell = spawn('powershell.exe', ['-ExecutionPolicy', 'Bypass', '-File', scriptPath]);

      let output = '';
      let errorOutput = '';

      powershell.stdout.on('data', (data) => {
        output += data.toString();
      });

      powershell.stderr.on('data', (data) => {
        errorOutput += data.toString();
      });

      powershell.on('close', (code) => {
        if (code === 0) {
          resolve({ success: true, output: output });
        } else {
          resolve({ success: false, error: `PowerShell script failed with code ${code}. Error: ${errorOutput}` });
        }
      });

      powershell.on('error', (error) => {
        this.logger.error(`Failed to start PowerShell: ${error.message}`);
        resolve({ success: false, error: error.message });
      });
    });
  }

  private parseEnabledFieldsFromOutput(output: string): any[] {
    try {
      // Extract JSON from PowerShell output
      const match = output.match(/RESULT:\s*(\{[\s\S]*?\})/);
      if (!match) {
        this.logger.warn('No RESULT found in PowerShell output');
        return [];
      }

      const jsonStr = match[1].trim();
      const result = JSON.parse(jsonStr);
      
      if (result.success && result.enabledFields) {
        return result.enabledFields;
      } else {
        this.logger.warn('Invalid result structure from PowerShell');
        return [];
      }
    } catch (error) {
      this.logger.warn(`Failed to parse PowerShell output: ${error.message}`);
      return [];
    }
  }
}
