import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import { spawn } from 'child_process';
import { SessionManagementService } from '../session-management/session-management.service';
import { ComProcessManagerService } from '../session-management/com-process-manager.service';

@Injectable()
export class PricingService {
  private readonly logger = new Logger(PricingService.name);

  constructor(
    private readonly sessionManagementService: SessionManagementService,
    private readonly comProcessManagerService: ComProcessManagerService
  ) {}

  async savePricingToExcel(
    opportunityId: string,
    pricingData: {
      batteryType: string;
      numberOfPanels: number;
      additionalItems: string[];
      totalSystemCost: number;
      targetCell: string;
      calculatorType: string;
    }
  ) {
    try {
      this.logger.log(`💰 Saving pricing for opportunity ${opportunityId} to Excel cell ${pricingData.targetCell}`);
      this.logger.log('Pricing data:', pricingData);

      // Find the Excel file for this opportunity
      const calculatorType = pricingData.calculatorType as 'flux' | 'off-peak';
      const excelFilePath = await this.findOpportunityExcelFile(opportunityId, calculatorType);
      if (!excelFilePath) {
        throw new Error(`No Excel file found for opportunity ${opportunityId} with calculator type ${calculatorType}`);
      }

      this.logger.log(`📁 Found Excel file: ${excelFilePath}`);

      // Determine the correct target cell based on calculator type
      let targetCell = pricingData.targetCell;
      if (pricingData.calculatorType === 'flux') {
        targetCell = 'H81'; // EPVS/Flux uses H81
        this.logger.log(`🔄 Flux calculator detected, using cell H81`);
      } else {
        targetCell = 'H80'; // Off Peak uses H80
        this.logger.log(`⚡ Off Peak calculator detected, using cell H80`);
      }
      
      // Use COM automation to save the pricing to Excel
      const result = await this.saveToExcelUsingCOM(
        excelFilePath,
        targetCell,
        pricingData.totalSystemCost
      );

      this.logger.log(`✅ Successfully saved £${pricingData.totalSystemCost} to cell ${targetCell}`);

      return {
        opportunityId,
        targetCell: targetCell,
        totalSystemCost: pricingData.totalSystemCost,
        calculatorType: pricingData.calculatorType,
        excelFilePath,
        savedAt: new Date().toISOString()
      };

    } catch (error) {
      this.logger.error(`❌ Error saving pricing to Excel: ${error.message}`);
      throw error;
    }
  }

  async getSystemCostsInputs(opportunityId: string) {
    try {
      this.logger.log(`📊 Getting system costs inputs for opportunity ${opportunityId}`);
      
      // Return the available input fields for system costs
      return {
        totalSystemCost: 'H80', // For Off Peak calculator
        deposit: 'H81',
        interestRate: 'H82',
        interestRateType: 'H83',
        paymentTerm: 'H84',
        // For EPVS/Flux calculator, the total system cost goes to H81
        totalSystemCostFlux: 'H81'
      };
    } catch (error) {
      this.logger.error(`❌ Error getting system costs inputs: ${error.message}`);
      throw error;
    }
  }

  private async findOpportunityExcelFile(opportunityId: string, calculatorType?: 'flux' | 'off-peak'): Promise<string | null> {
    try {
      // Determine which directory to search based on calculator type
      let searchDir: string;
      let baseFileName: string;
      
      if (calculatorType === 'flux') {
        // For EPVS/Flux calculator, search in epvs-opportunities directory
        searchDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');
        baseFileName = `EPVS Calculator Creativ - 06.02-${opportunityId}`;
        this.logger.log(`🔍 Searching for EPVS/Flux file with base name: ${baseFileName}`);
      } else {
        // For Off Peak calculator, search in opportunities directory
        searchDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
        baseFileName = `Off peak V2.1 Eon SEG-${opportunityId}`;
        this.logger.log(`🔍 Searching for Off Peak file with base name: ${baseFileName}`);
      }
      
      if (!fs.existsSync(searchDir)) {
        this.logger.warn(`${calculatorType === 'flux' ? 'EPVS' : 'Off Peak'} opportunities directory not found: ${searchDir}`);
        return null;
      }

      const files = fs.readdirSync(searchDir);
      this.logger.log(`Found ${files.length} files in ${calculatorType === 'flux' ? 'EPVS' : 'Off Peak'} opportunities directory`);
      
      // Look for versioned files (v1, v2, v3, etc.) and find the latest one
      const basePattern = baseFileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // Escape special regex characters
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

      if (latestFile) {
        const filePath = path.join(searchDir, latestFile);
        this.logger.log(`Found latest version (v${maxVersion}) ${calculatorType === 'flux' ? 'EPVS' : 'Off Peak'} file: ${filePath}`);
        return filePath;
      }

      // Fallback: look for files that contain the opportunity ID (old format)
      for (const file of files) {
        if (file.includes(opportunityId) && (file.endsWith('.xlsx') || file.endsWith('.xlsm'))) {
          const filePath = path.join(searchDir, file);
          this.logger.log(`Found matching ${calculatorType === 'flux' ? 'EPVS' : 'Off Peak'} file (old format): ${filePath}`);
          return filePath;
        }
      }

      // If no exact match, look for any Excel file (fallback)
      for (const file of files) {
        if (file.endsWith('.xlsx') || file.endsWith('.xlsm')) {
          this.logger.warn(`No exact match found for ${opportunityId}, using fallback file: ${file}`);
          return path.join(searchDir, file);
        }
      }

      this.logger.warn(`No Excel files found in ${searchDir}`);
      return null;
    } catch (error) {
      this.logger.error(`Error finding Excel file: ${error.message}`);
      return null;
    }
  }

  private async saveToExcelUsingCOM(
    excelFilePath: string,
    targetCell: string,
    totalSystemCost: number
  ): Promise<void> {
    try {
      this.logger.log(`🔧 Using PowerShell COM automation to write £${totalSystemCost} to cell ${targetCell} in ${excelFilePath}`);
      
      // Create a PowerShell script to handle COM automation (same approach as ExcelAutomationService)
      const scriptContent = `
# Pricing COM Automation Script
$ErrorActionPreference = "Stop"

# Configuration
$excelFilePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
$targetCell = "${targetCell}"
$totalSystemCost = ${totalSystemCost}
$password = "99"

Write-Host "Starting pricing COM automation..." -ForegroundColor Green
Write-Host "Excel file: $excelFilePath" -ForegroundColor Yellow
Write-Host "Target cell: $targetCell" -ForegroundColor Yellow
Write-Host "Total system cost: £$totalSystemCost" -ForegroundColor Yellow

try {
    # Kill any existing Excel processes to avoid file locks
    Write-Host "Checking for existing Excel processes..." -ForegroundColor Yellow
    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        Write-Host "Found existing Excel processes, terminating them..." -ForegroundColor Yellow
        $excelProcesses | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
    
    # Create Excel application
    Write-Host "Creating Excel application..." -ForegroundColor Green
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Check if file exists and is accessible
    Write-Host "Checking file accessibility..." -ForegroundColor Yellow
    if (-not (Test-Path $excelFilePath)) {
        throw "Excel file does not exist: $excelFilePath"
    }
    
    # Try to get file info to check if it's locked
    try {
        $fileInfo = Get-Item $excelFilePath
        Write-Host "Excel file size: $($fileInfo.Length) bytes" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not get file info: $_" -ForegroundColor Yellow
    }
    
    # Open the workbook
    Write-Host "Opening workbook: $excelFilePath" -ForegroundColor Green
    
    try {
        $workbook = $excel.Workbooks.Open($excelFilePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without password..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($excelFilePath)
            Write-Host "Workbook opened successfully without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook: $_" -ForegroundColor Red
            throw "Could not open workbook"
        }
    }
    
    # Get the Inputs worksheet (same as ExcelAutomationService)
    $worksheet = $workbook.Worksheets.Item("Inputs")
    Write-Host "Accessing worksheet: $($worksheet.Name)" -ForegroundColor Green
    
    # Check if worksheet is protected and unprotect if needed
    if ($worksheet.ProtectContents) {
        Write-Host "Worksheet is protected, attempting to unprotect..." -ForegroundColor Yellow
        try {
            $worksheet.Unprotect($password)
            Write-Host "Successfully unprotected worksheet" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not unprotect worksheet: $_" -ForegroundColor Yellow
        }
    }
    
    # Write the total system cost to the target cell
    Write-Host "Writing £$totalSystemCost to cell $targetCell..." -ForegroundColor Green
    $worksheet.Range($targetCell).Value = $totalSystemCost
    Write-Host "Successfully wrote £$totalSystemCost to cell $targetCell" -ForegroundColor Green
    
    # Save the workbook
    Write-Host "Saving workbook..." -ForegroundColor Green
    $workbook.Save()
    Write-Host "Workbook saved successfully" -ForegroundColor Green
    
    # Close the workbook
    Write-Host "Closing workbook..." -ForegroundColor Green
    $workbook.Close()
    
    # Quit Excel
    Write-Host "Quitting Excel..." -ForegroundColor Green
    $excel.Quit()
    
    # Release COM objects
    Write-Host "Releasing COM objects..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Successfully completed pricing COM automation!" -ForegroundColor Green
    Write-Host "Total system cost £$totalSystemCost written to cell $targetCell" -ForegroundColor Green
    
} catch {
    Write-Error "Critical error in pricing COM automation: $_"
    exit 1
}
      `;
      
      // Write script to temporary file
      const scriptPath = path.join(process.cwd(), `temp-pricing-com-${Date.now()}.ps1`);
      fs.writeFileSync(scriptPath, scriptContent);
      
      this.logger.log(`Created PowerShell script: ${scriptPath}`);
      
      // Execute PowerShell script using the same approach as ExcelAutomationService
      const result = await this.runPowerShellScript(scriptPath);
      
      // Clean up temporary script
      try {
        fs.unlinkSync(scriptPath);
        this.logger.log(`Cleaned up temporary script: ${scriptPath}`);
      } catch (cleanupError) {
        this.logger.warn(`Failed to cleanup temporary script: ${cleanupError.message}`);
      }
      
      if (result.success) {
        this.logger.log(`✅ Successfully wrote £${totalSystemCost} to cell ${targetCell} using COM automation`);
      } else {
        throw new Error(`PowerShell script failed: ${result.error}`);
      }
      
    } catch (error) {
      this.logger.error(`Error in COM automation: ${error.message}`);
      throw new Error(`Failed to write to Excel using COM: ${error.message}`);
    }
  }

  private async runPowerShellScript(scriptPath: string): Promise<{ success: boolean; error?: string; output?: string }> {
    return new Promise((resolve) => {
      this.logger.log(`Executing PowerShell script for pricing COM automation...`);
      
      const powershell = spawn('powershell.exe', ['-ExecutionPolicy', 'Bypass', '-File', scriptPath], {
        stdio: ['pipe', 'pipe', 'pipe'],
        shell: true
      });

      let stdout = '';
      let stderr = '';

      powershell.stdout.on('data', (data) => {
        stdout += data.toString();
        this.logger.log(`PowerShell output: ${data.toString().trim()}`);
      });

      powershell.stderr.on('data', (data) => {
        stderr += data.toString();
        this.logger.error(`PowerShell error: ${data.toString().trim()}`);
      });

      powershell.on('close', (code) => {
        this.logger.log(`PowerShell script completed with code: ${code}`);
        
        if (code === 0) {
          resolve({ success: true, output: stdout });
        } else {
          const error = stderr || `PowerShell script failed with code ${code}`;
          resolve({ success: false, error });
        }
      });

      powershell.on('error', (error) => {
        this.logger.error(`Failed to start PowerShell: ${error.message}`);
        resolve({ success: false, error: error.message });
      });
    });
  }

  /**
   * Save pricing to Excel with user session isolation
   */
  async savePricingToExcelWithSession(
    userId: string,
    opportunityId: string,
    pricingData: {
      batteryType: string;
      numberOfPanels: number;
      additionalItems: string[];
      totalSystemCost: number;
      targetCell: string;
      calculatorType: string;
    }
  ): Promise<{ success: boolean; message: string; error?: string }> {
    this.logger.log(`Saving pricing with session isolation for user: ${userId}, opportunity: ${opportunityId}`);

    try {
      // Queue the request through session management
      const result = await this.sessionManagementService.queueRequest(
        userId,
        'pricing_save',
        'database',
        {
          opportunityId,
          pricingData
        },
        1 // High priority
      );

      return result;
    } catch (error) {
      this.logger.error(`Session-based pricing save failed for user ${userId}:`, error);
      return { 
        success: false, 
        message: 'Pricing save failed', 
        error: error.message 
      };
    }
  }

  /**
   * Execute pricing save with user isolation (called by session management)
   */
  async executePricingSaveWithIsolation(
    userId: string,
    data: {
      opportunityId: string;
      pricingData: {
        batteryType: string;
        numberOfPanels: number;
        additionalItems: string[];
        totalSystemCost: number;
        targetCell: string;
        calculatorType: string;
      };
    }
  ): Promise<{ success: boolean; message: string; error?: string }> {
    const session = await this.sessionManagementService.createOrGetSession(userId);
    const workingDirectory = session.workingDirectory;

    this.logger.log(`Executing pricing save with isolation for user: ${userId}`);

    try {
      // Find the Excel file for this opportunity
      const calculatorType = data.pricingData.calculatorType as 'flux' | 'off-peak';
      const excelFilePath = await this.findOpportunityExcelFile(data.opportunityId, calculatorType);
      if (!excelFilePath) {
        throw new Error(`No Excel file found for opportunity ${data.opportunityId} with calculator type ${calculatorType}`);
      }

      this.logger.log(`📁 Found Excel file: ${excelFilePath}`);

      // Determine the correct target cell based on calculator type
      let targetCell = data.pricingData.targetCell;
      if (calculatorType === 'flux') {
        // For Flux calculator, use different cell
        targetCell = 'H50'; // Adjust as needed for Flux calculator
      }

      // Create user-specific PowerShell script
      const scriptContent = this.createIsolatedPricingScript(
        excelFilePath,
        targetCell,
        data.pricingData.totalSystemCost,
        workingDirectory
      );
      
      // Create temporary script file in user's directory
      const tempScriptPath = path.join(workingDirectory, 'temp', `pricing_save_${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, scriptContent);
      
      this.logger.log(`Created isolated pricing PowerShell script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Could not clean up temporary pricing script: ${cleanupError.message}`);
      }

      if (result.success) {
        return {
          success: true,
          message: 'Pricing saved successfully with user isolation'
        };
      } else {
        return {
          success: false,
          message: 'Pricing save failed',
          error: result.error || 'Unknown error'
        };
      }
    } catch (error) {
      this.logger.error(`Isolated pricing save failed for user ${userId}:`, error);
      return { 
        success: false, 
        message: 'Pricing save failed', 
        error: error.message 
      };
    }
  }

  /**
   * Create isolated pricing script for user session
   */
  private createIsolatedPricingScript(
    excelFilePath: string,
    targetCell: string,
    totalSystemCost: number,
    workingDirectory: string
  ): string {
    return `
# Isolated Pricing COM Automation Script - User Session
$ErrorActionPreference = "Stop"

# Configuration
$excelFilePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
$targetCell = "${targetCell}"
$totalSystemCost = ${totalSystemCost}
$workingDirectory = "${workingDirectory.replace(/\\/g, '\\\\')}"
$password = "99"

Write-Host "Starting isolated pricing COM automation..." -ForegroundColor Green
Write-Host "Working Directory: $workingDirectory" -ForegroundColor Yellow
Write-Host "Excel file: $excelFilePath" -ForegroundColor Yellow
Write-Host "Target cell: $targetCell" -ForegroundColor Yellow
Write-Host "Total system cost: £$totalSystemCost" -ForegroundColor Yellow

try {
    # Create user-specific Excel application
    Write-Host "Creating isolated Excel application for pricing..." -ForegroundColor Green
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Isolated Excel application created successfully" -ForegroundColor Green
    
    # Check if file exists and is accessible
    Write-Host "Checking file accessibility..." -ForegroundColor Yellow
    if (-not (Test-Path $excelFilePath)) {
        throw "Excel file does not exist: $excelFilePath"
    }
    
    # Open the Excel file
    Write-Host "Opening Excel file..." -ForegroundColor Green
    $workbook = $excel.Workbooks.Open($excelFilePath, 0, $false, 5, $password)
    Write-Host "Excel file opened successfully" -ForegroundColor Green
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    Write-Host "Accessing worksheet: $($worksheet.Name)" -ForegroundColor Green
    
    # Unprotect the worksheet if needed
    if ($worksheet.ProtectContents) {
        Write-Host "Unprotecting worksheet..." -ForegroundColor Yellow
        $worksheet.Unprotect($password)
        Write-Host "Worksheet unprotected successfully" -ForegroundColor Green
    }
    
    # Write the pricing to the target cell
    Write-Host "Writing £$totalSystemCost to cell $targetCell..." -ForegroundColor Green
    $worksheet.Range($targetCell).Value = $totalSystemCost
    Write-Host "Pricing written successfully" -ForegroundColor Green
    
    # Save the workbook
    Write-Host "Saving workbook..." -ForegroundColor Green
    $workbook.Save()
    Write-Host "Workbook saved successfully" -ForegroundColor Green
    
    # Close workbook
    $workbook.Close()
    Write-Host "Workbook closed successfully" -ForegroundColor Green
    
    # Quit Excel
    Write-Host "Quitting Excel..." -ForegroundColor Green
    $excel.Quit()
    
    # Release COM objects
    Write-Host "Releasing COM objects..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Successfully completed isolated pricing COM automation!" -ForegroundColor Green
    Write-Host "Total system cost £$totalSystemCost written to cell $targetCell" -ForegroundColor Green
    
} catch {
    Write-Error "Critical error in isolated pricing COM automation: $_"
    exit 1
}

Write-Host "Isolated pricing automation completed!" -ForegroundColor Green
exit 0
`;
  }
}
