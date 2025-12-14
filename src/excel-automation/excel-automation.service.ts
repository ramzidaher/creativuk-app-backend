import { Injectable, Logger } from '@nestjs/common';
import { spawn } from 'child_process';
import * as path from 'path';
import * as fs from 'fs';
import * as XLSX from 'xlsx';
import { PdfSignatureService } from '../pdf-signature/pdf-signature.service';
import { SessionManagementService } from '../session-management/session-management.service';
import { ComProcessManagerService } from '../session-management/com-process-manager.service';

@Injectable()
export class ExcelAutomationService {
  private readonly logger = new Logger(ExcelAutomationService.name)

  /**
   * Clean and escape a string for PowerShell usage
   */
  private escapeForPowerShell(value: string): string {
    if (!value) return '';
    
    return String(value)
      // Remove or replace problematic Unicode characters
      .replace(/[\u2018\u2019]/g, "'") // Replace smart quotes with regular apostrophes
      .replace(/[\u201C\u201D]/g, '"') // Replace smart quotes with regular quotes
      .replace(/[\u2022\u2023\u25E6\u2043\u2219]/g, '•') // Replace various bullet points
      .replace(/[\u2013\u2014]/g, '-') // Replace en/em dashes with regular dashes
      .replace(/[\u00A0]/g, ' ') // Replace non-breaking spaces with regular spaces
      .replace(/[\u00B0]/g, '°') // Replace degree symbol
      .replace(/[\u00A3]/g, '£') // Replace pound symbol
      .replace(/[\u20AC]/g, '€') // Replace euro symbol
      .replace(/[\u00A5]/g, '¥') // Replace yen symbol
      .replace(/[\u00A2]/g, '¢') // Replace cent symbol
      // Escape PowerShell special characters
      .replace(/"/g, '\\"') // Escape double quotes
      .replace(/\$/g, '`$') // Escape dollar signs
      .replace(/`/g, '``') // Escape backticks
      .replace(/\n/g, ' ') // Replace newlines with spaces
      .replace(/\r/g, ' ') // Replace carriage returns with spaces
      .replace(/\t/g, ' ') // Replace tabs with spaces
      .trim(); // Remove leading/trailing whitespace
  };

  constructor(
    private readonly pdfSignatureService: PdfSignatureService,
    private readonly sessionManagementService: SessionManagementService,
    private readonly comProcessManagerService: ComProcessManagerService
  ) {}

  private readonly TEMPLATES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'templates');
  private readonly DEFAULT_TEMPLATE_FILE = 'Off peak V2.1 Eon SEG - All Options.xlsm';
  private readonly OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
  private readonly EPVS_OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');
  private readonly PASSWORD = '99';

  /**
   * Get the correct Excel file path based on calculator type with versioning support
   * For template selection - always creates new version
   */
  private getNewOpportunityFilePath(opportunityId: string, calculatorType?: 'flux' | 'off-peak'): string {
    if (calculatorType === 'flux') {
      // For EPVS/Flux calculator, use the EPVS opportunities folder
      return this.getNewVersionedFilePath(this.EPVS_OPPORTUNITIES_FOLDER, `EPVS Calculator Creativ - 06.02-${opportunityId}`, 'xlsm');
    } else {
      // For Off Peak calculator, use the regular opportunities folder
      return this.getNewVersionedFilePath(this.OPPORTUNITIES_FOLDER, `Off peak V2.1 Eon SEG-${opportunityId}`, 'xlsm');
    }
  }

  /**
   * Get the correct Excel file path based on calculator type with versioning support
   * For regular operations - uses existing file if available
   */
  private getOpportunityFilePath(opportunityId: string, calculatorType?: 'flux' | 'off-peak'): string {
    if (calculatorType === 'flux') {
      // For EPVS/Flux calculator, use the EPVS opportunities folder
      return this.getVersionedFilePath(this.EPVS_OPPORTUNITIES_FOLDER, `EPVS Calculator Creativ - 06.02-${opportunityId}`, 'xlsm');
    } else {
      // For Off Peak calculator, use the regular opportunities folder
      return this.getVersionedFilePath(this.OPPORTUNITIES_FOLDER, `Off peak V2.1 Eon SEG-${opportunityId}`, 'xlsm');
    }
  }

  /**
   * Get versioned file path (v1, v2, v3, etc.) for opportunity files
   * Only creates new versions when no file exists for the opportunity
   */
  private getVersionedFilePath(directory: string, baseFileName: string, extension: string): string {
    // Ensure directory exists
    if (!fs.existsSync(directory)) {
      fs.mkdirSync(directory, { recursive: true });
    }

    // Check for existing files with the same base name and find max version
    const files = fs.readdirSync(directory);
    const basePattern = baseFileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // Escape special regex characters
    const versionRegex = new RegExp(`^${basePattern}-v(\\d+)\\.${extension}$`);
    
    let maxVersion = 0;
    for (const file of files) {
      // Check for versioned files: baseFileName-v1.xlsm, baseFileName-v2.xlsm, etc.
      const versionMatch = file.match(versionRegex);
      if (versionMatch) {
        const version = parseInt(versionMatch[1], 10);
        if (version > maxVersion) {
          maxVersion = version;
        }
      }
      
      // Check for non-versioned files: baseFileName.xlsm (treat as v0)
      if (file === `${baseFileName}.${extension}`) {
        maxVersion = Math.max(maxVersion, 0);
      }
    }

    // Always create the next version (v1 if no files exist, v2 if v1 exists, etc.)
    const nextVersion = maxVersion + 1;
    this.logger.log(`🎯 getVersionedFilePath: Creating new version v${nextVersion} (max existing: v${maxVersion})`);
    return path.join(directory, `${baseFileName}-v${nextVersion}.${extension}`);
  }

  /**
   * Get new versioned file path for template selection
   * This method ALWAYS creates a new version (for when user selects a new template)
   */
  private getNewVersionedFilePath(directory: string, baseFileName: string, extension: string): string {
    // Ensure directory exists
    if (!fs.existsSync(directory)) {
      fs.mkdirSync(directory, { recursive: true });
    }

    // Check for existing files with the same base name
    const files = fs.readdirSync(directory);
    const basePattern = baseFileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // Escape special regex characters
    const versionRegex = new RegExp(`^${basePattern}-v(\\d+)\\.${extension}$`);
    
    this.logger.log(`🔍 getNewVersionedFilePath: Looking for files matching base: ${baseFileName}`);
    this.logger.log(`🔍 Regex pattern: ^${basePattern}-v(\\d+)\\.${extension}$`);
    this.logger.log(`🔍 Files in directory (${files.length} total):`);
    files.slice(0, 10).forEach(file => {
      this.logger.log(`   - ${file}`);
    });
    
    let maxVersion = 0;
    for (const file of files) {
      const match = file.match(versionRegex);
      if (match) {
        const version = parseInt(match[1], 10);
        this.logger.log(`✅ Found version match: ${file} -> version ${version}`);
        if (version > maxVersion) {
          maxVersion = version;
        }
      }
    }

    // Return the next version
    const nextVersion = maxVersion + 1;
    this.logger.log(`🎯 Creating new version v${nextVersion} (max existing: v${maxVersion})`);
    const newFilePath = path.join(directory, `${baseFileName}-v${nextVersion}.${extension}`);
    this.logger.log(`📝 New file path: ${newFilePath}`);
    return newFilePath;
  }

  /**
   * Find the latest version of an existing opportunity file
   */
  private findLatestOpportunityFile(opportunityId: string, calculatorType?: 'flux' | 'off-peak', fileName?: string): string | null {
    const directory = calculatorType === 'flux' ? this.EPVS_OPPORTUNITIES_FOLDER : this.OPPORTUNITIES_FOLDER;
    const baseFileName = calculatorType === 'flux' ? `EPVS Calculator Creativ - 06.02-${opportunityId}` : `Off peak V2.1 Eon SEG-${opportunityId}`;
    
    if (!fs.existsSync(directory)) {
      return null;
    }

    const files = fs.readdirSync(directory);
    
    // If fileName is provided, try to find the exact file first
    if (fileName) {
      this.logger.log(`🔍 Looking for specific file: ${fileName}`);
      const exactFile = files.find(file => file === fileName);
      if (exactFile) {
        const fullPath = path.join(directory, exactFile);
        this.logger.log(`✅ Found exact file match: ${fullPath}`);
        return fullPath;
      } else {
        this.logger.log(`⚠️ Exact file ${fileName} not found, falling back to version search`);
      }
    }
    
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

    return latestFile ? path.join(directory, latestFile) : null;
  }

  private getTemplateFilePath(templateFileName?: string): string {
    const fileName = templateFileName || this.DEFAULT_TEMPLATE_FILE;
    return path.join(this.TEMPLATES_FOLDER, fileName);
  }

  async selectRadioButton(shapeName: string, opportunityId?: string): Promise<{ success: boolean; message: string; error?: string }> {
    this.logger.log(`Starting radio button automation for shape: ${shapeName}${opportunityId ? ` (Opportunity: ${opportunityId})` : ''}`);

    try {
      // Determine which file to use
      let excelFilePath: string;
      if (opportunityId) {
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId, 'off-peak');
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
        } else {
          this.logger.warn(`Opportunity file not found, using template: ${opportunityFilePath}`);
          excelFilePath = this.getTemplateFilePath();
        }
      } else {
        excelFilePath = this.getTemplateFilePath();
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

      // Create PowerShell script
      const psScript = this.createRadioButtonScript(shapeName, excelFilePath);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-radio-button-${Date.now()}.ps1`);
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
        this.logger.log(`Successfully selected radio button: ${shapeName}`);
        return {
          success: true,
          message: `Successfully selected radio button: ${shapeName}`
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to select radio button: ${shapeName}`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in selectRadioButton: ${error.message}`);
      return {
        success: false,
        message: `Error selecting radio button: ${shapeName}`,
        error: error.message
      };
    }
  }

  async createOpportunityFile(opportunityId: string, customerDetails: { customerName: string; address: string; postcode: string }, templateFileName?: string, isTemplateSelection: boolean = false): Promise<{ success: boolean; message: string; error?: string; filePath?: string }> {
    this.logger.log(`Creating opportunity file for: ${opportunityId} with template: ${templateFileName || 'default'}`);

    try {
      // Check if template file exists
      const templateFilePath = this.getTemplateFilePath(templateFileName);
      this.logger.log(`Using template file path: ${templateFilePath}`);
      if (!fs.existsSync(templateFilePath)) {
        const error = `Template file not found at: ${templateFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Check if we're on Windows (required for Excel COM automation)
      if (process.platform !== 'win32') {
        const error = 'Excel automation requires Windows platform';
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Create opportunities folder if it doesn't exist
      if (!fs.existsSync(this.OPPORTUNITIES_FOLDER)) {
        fs.mkdirSync(this.OPPORTUNITIES_FOLDER, { recursive: true });
        this.logger.log(`Created opportunities folder: ${this.OPPORTUNITIES_FOLDER}`);
      }

      // Create PowerShell script
      const psScript = this.createOpportunityFileScript(opportunityId, customerDetails, templateFilePath, 'off-peak', isTemplateSelection);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-create-opportunity-${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created temporary PowerShell script: ${tempScriptPath}`);
      this.logger.log(`Script content preview: ${psScript.substring(0, 500)}...`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Failed to cleanup temporary script: ${cleanupError.message}`);
      }

      if (result.success) {
        const filePath = this.getOpportunityFilePath(opportunityId, 'off-peak');
        this.logger.log(`Successfully created opportunity file: ${filePath}`);
        return {
          success: true,
          message: `Successfully created opportunity file for ${opportunityId}`,
          filePath
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to create opportunity file: ${opportunityId}`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in createOpportunityFile: ${error.message}`);
      return {
        success: false,
        message: `Error creating opportunity file: ${opportunityId}`,
        error: error.message
      };
    }
  }

  async checkOpportunityFileExists(opportunityId: string): Promise<{ success: boolean; exists: boolean; message: string; filePath?: string }> {
    try {
      const filePath = this.getOpportunityFilePath(opportunityId, 'off-peak');
      
      if (fs.existsSync(filePath)) {
        this.logger.log(`Opportunity file exists: ${filePath}`);
        return {
          success: true,
          exists: true,
          message: `Opportunity file exists for ${opportunityId}`,
          filePath: filePath
        };
      } else {
        this.logger.log(`Opportunity file does not exist: ${filePath}`);
        return {
          success: true,
          exists: false,
          message: `Opportunity file does not exist for ${opportunityId}`,
          filePath: filePath
        };
      }
    } catch (error) {
      this.logger.error(`Error checking opportunity file existence: ${error.message}`);
      return {
        success: false,
        exists: false,
        message: `Error checking opportunity file: ${error.message}`
      };
    }
  }

  private createOpportunityFileScript(opportunityId: string, customerDetails: { customerName: string; address: string; postcode: string }, templateFilePath?: string, calculatorType?: 'flux' | 'off-peak', isNewTemplate?: boolean): string {
    const templatePath = (templateFilePath || this.getTemplateFilePath()).replace(/\\/g, '\\\\');
    const opportunitiesFolder = this.OPPORTUNITIES_FOLDER.replace(/\\/g, '\\\\');
    const newFilePath = isNewTemplate ? 
      this.getNewOpportunityFilePath(opportunityId, calculatorType).replace(/\\/g, '\\\\') :
      this.getOpportunityFilePath(opportunityId, calculatorType).replace(/\\/g, '\\\\');
    const isFlux = calculatorType === 'flux';
    
    // Build consumption mappings based on calculator type
    const consumptionMappings = isFlux ? 
      `# Flux calculator uses H26-H32 for consumption fields
        "estimated_annual_usage" = "H26"
        "annual_usage" = "H26"
        "estimated_peak_annual_usage" = "H27"
        "estimated_off_peak_usage" = "H28"
        "standing_charges" = "H29"
        "total_annual_spend" = "H30"
        "peak_annual_spend" = "H31"
        "off_peak_annual_spend" = "H32"` : 
      `# Off-peak calculator uses H26-H28 for consumption
        "annual_usage" = "H26"
        "estimated_annual_usage" = "H26"
        "standing_charge" = "H27"
        "annual_spend" = "H28"
        
        # ENERGY USE - EXPORT TARIFF
        "export_tariff_rate" = "H30"`;
    
    return `
# Create Opportunity File with Customer Details
$ErrorActionPreference = "Stop"

# Configuration
$templatePath = "${templatePath}"
$newFilePath = "${newFilePath}"
$password = "${this.PASSWORD}"

Write-Host "Creating opportunity file for: ${opportunityId}" -ForegroundColor Green
Write-Host "Template path: $templatePath" -ForegroundColor Yellow
Write-Host "New file path: $newFilePath" -ForegroundColor Yellow

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
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Check if template file exists and is accessible
    Write-Host "Checking template file accessibility..." -ForegroundColor Yellow
    if (-not (Test-Path $templatePath)) {
        throw "Template file does not exist: $templatePath"
    }
    
    # Try to get file info to check if it's locked
    try {
        $fileInfo = Get-Item $templatePath
        Write-Host "Template file size: $($fileInfo.Length) bytes" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not get file info: $_" -ForegroundColor Yellow
    }
    
    # First, try to copy the template file directly
    Write-Host "Attempting to copy template file..." -ForegroundColor Yellow
    try {
        Copy-Item -Path $templatePath -Destination $newFilePath -Force
        Write-Host "Template file copied successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to copy template file: $_" -ForegroundColor Red
        throw "Could not copy template file"
    }
    
    # Now open the copied file
    Write-Host "Opening copied workbook: $newFilePath" -ForegroundColor Green
    
    try {
        $workbook = $excel.Workbooks.Open($newFilePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without password..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($newFilePath)
            Write-Host "Workbook opened successfully without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook: $_" -ForegroundColor Red
            throw "Could not open copied workbook"
        }
    }
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
    # Check if worksheet is protected and unprotect if needed
    if ($worksheet.ProtectContents) {
        Write-Host "Worksheet is protected, attempting to unprotect..." -ForegroundColor Yellow
        $worksheet.Unprotect($password)
        Write-Host "Successfully unprotected worksheet" -ForegroundColor Green
    }
    
    # Fill in customer details
    Write-Host "Filling in customer details..." -ForegroundColor Green
    
    # Customer Name (H12)
    $customerName = "${this.escapeForPowerShell(customerDetails.customerName)}"
    $worksheet.Range("H12").Value = $customerName
    Write-Host "Set Customer Name: $customerName" -ForegroundColor Green
    
    # Address (H13)
    $address = "${this.escapeForPowerShell(customerDetails.address)}"
    $worksheet.Range("H13").Value = $address
    Write-Host "Set Address: $address" -ForegroundColor Green
    
    # Postcode (H14) - only if provided
    $postcode = "${this.escapeForPowerShell(customerDetails.postcode)}"
    if ($postcode -ne "") {
        $worksheet.Range("H14").Value = $postcode
        Write-Host "Set Postcode: $postcode" -ForegroundColor Green
    }
    
    # Save the modified file
    Write-Host "Saving modified file..." -ForegroundColor Green
    try {
        $workbook.Save()
        Write-Host "File saved successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to save file: $_" -ForegroundColor Red
        throw "Could not save modified file"
    }
    
    # Close workbook and Excel
    $workbook.Close($false)
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Opportunity file creation completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    exit 1
}
`;
  }

  private createRadioButtonScript(shapeName: string, excelFilePath: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    return `
# Radio Button Automation using COM
$ErrorActionPreference = "Stop"

# Configuration
$shapeName = "${shapeName}"

Write-Host "Starting radio button automation for shape: $shapeName" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook with password
    $filePath = "${excelFilePathEscaped}"
    
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, "${this.PASSWORD}")
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        $workbook = $excel.Workbooks.Open($filePath)
        Write-Host "Workbook opened without password" -ForegroundColor Green
    }
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    if (!$worksheet) {
        throw "Inputs worksheet not found"
    }
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
    # Unprotect ALL worksheets in the workbook
    Write-Host "Unprotecting all worksheets..." -ForegroundColor Yellow
    foreach ($ws in $workbook.Worksheets) {
        try {
            if ($ws.ProtectContents) {
                Write-Host "Unprotecting worksheet: $($ws.Name)" -ForegroundColor Cyan
                try {
                    $ws.Unprotect("${this.PASSWORD}")
                    Write-Host "Successfully unprotected worksheet: $($ws.Name)" -ForegroundColor Green
                } catch {
                    try {
                        $ws.Unprotect()
                        Write-Host "Successfully unprotected worksheet without password: $($ws.Name)" -ForegroundColor Green
                    } catch {
                        Write-Host "Warning: Could not unprotect worksheet $($ws.Name): $_" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "Worksheet $($ws.Name) is not protected" -ForegroundColor Green
            }
        } catch {
            Write-Host "Error processing worksheet $($ws.Name): $_" -ForegroundColor Yellow
        }
    }
    
    Write-Host "Looking for radio button shape: $shapeName" -ForegroundColor Yellow
    $shapeFound = $false
    
    # Look for the specific shape
    try {
        $targetShape = $worksheet.Shapes($shapeName)
        Write-Host "Found target radio button shape: $shapeName" -ForegroundColor Green
        
        # Get shape details
        $shapeText = ""
        try {
            if ($targetShape.TextFrame -and $targetShape.TextFrame.Characters) {
                $shapeText = $targetShape.TextFrame.Characters().Text
            }
        } catch {
            # Shape might not have text
        }
        
        Write-Host "Radio button text: '$shapeText'" -ForegroundColor Cyan
        
        # Check if it has ControlFormat
        $hasControlFormat = $false
        try {
            if ($targetShape.ControlFormat) {
                $hasControlFormat = $true
                Write-Host "Shape has ControlFormat" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "Shape does not have ControlFormat" -ForegroundColor Yellow
        }
        
        # Check if it has OnAction
        $onActionMacro = ""
        try {
            $onActionMacro = $targetShape.OnAction
            if ($onActionMacro) {
                Write-Host "Shape has OnAction macro: $onActionMacro" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "Shape does not have OnAction macro" -ForegroundColor Yellow
        }
        
        # Try to interact with the radio button
        $interactionSuccess = $false
        
        # Method 1: Try ControlFormat if available
        if ($hasControlFormat) {
            try {
                Write-Host "Attempting to select radio button using ControlFormat..." -ForegroundColor Yellow
                $currentValue = $targetShape.ControlFormat.Value
                Write-Host "Current ControlFormat value: $currentValue" -ForegroundColor Cyan
                
                # Set the radio button to selected (value = 1)
                $targetShape.ControlFormat.Value = 1
                
                Write-Host "Successfully selected radio button: $shapeName" -ForegroundColor Green
                $interactionSuccess = $true
            } catch {
                Write-Host "ControlFormat interaction failed: $_" -ForegroundColor Yellow
            }
        }
        
        # Method 2: Try to select the shape
        if (-not $interactionSuccess) {
            try {
                Write-Host "Attempting to select radio button shape..." -ForegroundColor Yellow
                $targetShape.Select()
                Start-Sleep -Milliseconds 100
                Write-Host "Successfully selected radio button shape: $shapeName" -ForegroundColor Green
                $interactionSuccess = $true
            } catch {
                Write-Host "Shape selection failed: $_" -ForegroundColor Yellow
            }
        }
        
        # Method 3: Try to click the shape
        if (-not $interactionSuccess) {
            try {
                Write-Host "Attempting to click radio button..." -ForegroundColor Yellow
                $targetShape.Select()
                Start-Sleep -Milliseconds 100
                # Simulate a click by pressing Enter
                $excel.SendKeys("{ENTER}")
                Start-Sleep -Milliseconds 100
                Write-Host "Successfully clicked radio button: $shapeName" -ForegroundColor Green
                $interactionSuccess = $true
            } catch {
                Write-Host "Radio button clicking failed: $_" -ForegroundColor Yellow
            }
        }
        
        if ($interactionSuccess) {
            # Try to trigger the OnAction macro if it exists
            if ($onActionMacro -and $onActionMacro.Trim() -ne "") {
                try {
                    Write-Host "Executing OnAction macro: $onActionMacro" -ForegroundColor Cyan
                    $excel.Run($onActionMacro)
                    Write-Host "Successfully executed OnAction macro: $onActionMacro" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to execute OnAction macro: $_" -ForegroundColor Yellow
                }
            }
            
            # Save the workbook
            try {
                $workbook.Save()
                Write-Host "Workbook saved successfully" -ForegroundColor Green
            } catch {
                Write-Host "Warning: Could not save workbook: $_" -ForegroundColor Yellow
            }
            
            $shapeFound = $true
        } else {
            Write-Host "Failed to interact with radio button: $shapeName" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "Error finding or interacting with shape '$shapeName': $_" -ForegroundColor Red
    }
    
    # Close Excel
    try {
        $workbook.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    if ($shapeFound) {
        Write-Host "Radio button automation completed successfully!" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "Radio button automation failed!" -ForegroundColor Red
        exit 1
    }
    
} catch {
    Write-Host "Critical error in radio button automation: $_" -ForegroundColor Red
    exit 1
}
    `;
  }

  private async runPowerShellScript(scriptPath: string): Promise<{ success: boolean; error?: string; output?: string }> {
    return new Promise((resolve) => {
      this.logger.log(`Executing PowerShell script for radio button automation...`);
      
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

  async getDynamicInputs(opportunityId?: string, templateFileName?: string): Promise<{ success: boolean; message: string; error?: string; inputFields?: any[] }> {
    this.logger.log(`Getting dynamic inputs${opportunityId ? ` for opportunity: ${opportunityId}` : ''}${templateFileName ? ` with template: ${templateFileName}` : ''}`);

    try {
      // Determine which file to use - prioritize opportunity file if it exists
      let excelFilePath: string;
      if (opportunityId) {
        // First, try to use the opportunity file (created from template)
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId, 'off-peak');
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using opportunity file: ${opportunityFilePath}`);
        } else {
          // If opportunity file doesn't exist, fall back to template file
          if (templateFileName) {
            excelFilePath = this.getTemplateFilePath(templateFileName);
            this.logger.log(`Opportunity file not found, using template file: ${excelFilePath}`);
        } else {
          const error = `Opportunity file not found: ${opportunityFilePath}. Please ensure radio button selections have been applied first.`;
          this.logger.error(error);
          return { success: false, message: error };
        }
        }
      } else if (templateFileName) {
        // No opportunityId, use template file directly
        excelFilePath = this.getTemplateFilePath(templateFileName);
        this.logger.log(`Using template file: ${excelFilePath}`);
      } else {
        // Fallback to default template
        excelFilePath = this.getTemplateFilePath();
        this.logger.log(`Using default template file: ${excelFilePath}`);
      }

      // Check if file exists
      if (!fs.existsSync(excelFilePath)) {
        const error = `Excel file not found at: ${excelFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Use Node.js script for reliable Excel analysis
      this.logger.log('Using Node.js script for Excel analysis...');
      const result = await this.runNodeJsExcelAnalysis(excelFilePath);
      
      if (result.success) {
        this.logger.log(`Node.js: Successfully retrieved ${result.inputFields?.length || 0} input fields`);
        
        // Debug: Log the enabled/disabled status of each field
        result.inputFields?.forEach(field => {
          this.logger.log(`Field ${field.id}: enabled=${field.enabled}, reason=${field.reason || 'N/A'}`);
        });
        
        // Get ALL dropdown options for ALL dropdown fields (not just enabled ones)
        const allDropdownOptions = await this.getAllDropdownOptions(excelFilePath);
        
        // Merge dropdown options with input fields
        const enhancedInputFields = result.inputFields?.map(field => {
          // Check if this is a dropdown field
          const isDropdownField = ['panel_manufacturer', 'panel_model', 'battery_manufacturer', 'battery_model', 
                                  'solar_inverter_manufacturer', 'solar_inverter_model', 
                                  'battery_inverter_manufacturer', 'battery_inverter_model'].includes(field.id);
          
          if (isDropdownField) {
            const options = allDropdownOptions[field.id] || [];
            this.logger.log(`Adding dropdown options for ${field.id}: ${options.length} options (enabled: ${field.enabled})`);
            return {
              ...field,
              type: 'dropdown',
              dropdownOptions: options
            };
          }
          return field;
        });
        
        this.logger.log(`Enhanced ${enhancedInputFields?.length || 0} input fields with dropdown options`);
        
        // Debug: Log what we're returning
        enhancedInputFields?.forEach(field => {
          if (field.type === 'dropdown') {
            this.logger.log(`Returning dropdown field ${field.id} with ${field.dropdownOptions?.length || 0} options (enabled: ${field.enabled})`);
          }
        });
        
        return {
          ...result,
          inputFields: enhancedInputFields
        };
      }

      // If Node.js method fails, return error
      this.logger.error('Node.js method failed');
      return {
        success: false,
        message: 'Failed to get dynamic inputs - Node.js method failed',
        error: result.error
      };

    } catch (error) {
      this.logger.error(`Error in getDynamicInputs: ${error.message}`);
      return {
        success: false,
        message: `Error getting dynamic inputs`,
        error: error.message
      };
    }
  }

  private parseInputFieldsFromPowerShellOutput(output: string): any[] {
    const inputFields: any[] = [];
    
    try {
      // Look for JSON-like output in the PowerShell response
      const jsonMatch = output.match(/RESULT:\s*(\{[\s\S]*?\})/);
      if (jsonMatch) {
        // Clean up the JSON string by removing extra whitespace and newlines
        let jsonString = jsonMatch[1].replace(/\s+/g, ' ').trim();
        
        // Fix common JSON issues
        jsonString = jsonString.replace(/\\"/g, '"'); // Fix escaped quotes
        jsonString = jsonString.replace(/\\n/g, ''); // Remove newlines
        jsonString = jsonString.replace(/\\r/g, ''); // Remove carriage returns
        jsonString = jsonString.replace(/\\t/g, ''); // Remove tabs
        
        // Try to parse the cleaned JSON
        try {
          const parsed = JSON.parse(jsonString);
          return parsed.inputFields || [];
        } catch (parseError) {
          this.logger.warn(`Failed to parse cleaned JSON: ${parseError.message}`);
          this.logger.warn(`JSON string: ${jsonString.substring(0, 200)}...`);
        }
      }
      
             // Fallback: parse line-by-line output
       const lines = output.split('\n');
       let currentField: any = null;
       
       for (const line of lines) {
         const trimmed = line.trim();
         
         if (trimmed.startsWith('FIELD:')) {
           if (currentField) {
             inputFields.push(currentField);
           }
           currentField = {};
           const fieldData = trimmed.substring(6).split('|');
           if (fieldData.length >= 4) {
             currentField.id = fieldData[0].trim();
             currentField.label = fieldData[1].trim();
             currentField.type = fieldData[2].trim();
             currentField.enabled = fieldData[3].trim() === 'true';
             currentField.cellReference = fieldData[4]?.trim() || '';
             currentField.required = fieldData[5]?.trim() === 'true' || false;
             currentField.value = fieldData[6]?.trim() || '';
             
             // Clean up the value field (remove the "Variant Value" text)
             if (currentField.value && currentField.value.includes('Variant Value')) {
               currentField.value = '';
             }
           }
         } else if (trimmed.startsWith('OPTIONS:') && currentField) {
           const options = trimmed.substring(8).split(',').map(opt => opt.trim());
           currentField.dropdownOptions = options;
         }
       }
       
       if (currentField) {
         inputFields.push(currentField);
       }
      
    } catch (error) {
      this.logger.warn(`Failed to parse PowerShell output: ${error.message}`);
    }
    
         // Return sample data if parsing failed - but only enabled fields
     if (inputFields.length === 0) {
       this.logger.warn('No input fields found, returning sample enabled fields');
       return [
         {
           id: 'annual_usage',
           label: 'Estimated Annual Usage (kWh)',
           type: 'number',
           value: '',
           required: true,
           enabled: true,
           cellReference: 'H27'
         },
         {
           id: 'standing_charge',
           label: 'Standing Charge (pence per day)',
           type: 'number',
           value: '',
           required: true,
           enabled: true,
           cellReference: 'H28'
         }
       ];
     }
    
    return inputFields;
  }

  async saveDynamicInputs(opportunityId: string | undefined, inputs: Record<string, string>, templateFileName?: string, calculatorType?: 'flux' | 'off-peak'): Promise<{ success: boolean; message: string; error?: string }> {
    this.logger.log(`Saving dynamic inputs${opportunityId ? ` for opportunity: ${opportunityId}` : ''}${templateFileName ? ` with template: ${templateFileName}` : ''}${calculatorType ? ` with calculator type: ${calculatorType}` : ''}`);

    try {
      // Determine which file to use - prioritize opportunity file if it exists
      let excelFilePath: string;
      if (opportunityId) {
        // First, try to use the opportunity file (created from template)
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId, calculatorType);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using ${calculatorType === 'flux' ? 'EPVS' : 'Off Peak'} opportunity file: ${opportunityFilePath}`);
        } else {
          // If opportunity file doesn't exist, fall back to template file
          if (templateFileName) {
            excelFilePath = this.getTemplateFilePath(templateFileName);
            this.logger.log(`Opportunity file not found, using template file: ${excelFilePath}`);
          } else {
            this.logger.warn(`Opportunity file not found, using default template: ${opportunityFilePath}`);
            excelFilePath = this.getTemplateFilePath();
          }
        }
      } else if (templateFileName) {
        // No opportunityId, use template file directly
        excelFilePath = this.getTemplateFilePath(templateFileName);
        this.logger.log(`Using template file: ${excelFilePath}`);
      } else {
        // Fallback to default template
        excelFilePath = this.getTemplateFilePath();
        this.logger.log(`Using default template file: ${excelFilePath}`);
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

      // Create PowerShell script
      const psScript = this.createSaveDynamicInputsScript(excelFilePath, inputs, calculatorType);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-save-inputs-${Date.now()}.ps1`);
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
        this.logger.log(`Successfully saved dynamic inputs`);
        return {
          success: true,
          message: `Successfully saved dynamic inputs`
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to save dynamic inputs`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in saveDynamicInputs: ${error.message}`);
      return {
        success: false,
        message: `Error saving dynamic inputs`,
        error: error.message
      };
    }
  }

  private createGetDynamicInputsScript(excelFilePath: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    return `
# Get Dynamic Inputs from Excel
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePathEscaped}"
$password = "${this.PASSWORD}"

Write-Host "Getting dynamic inputs from: $filePath" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        $workbook = $excel.Workbooks.Open($filePath)
        Write-Host "Workbook opened without password" -ForegroundColor Green
    }
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    if (!$worksheet) {
        throw "Inputs worksheet not found"
    }
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
         # Define ALL the input fields based on the Excel template structure
     $inputFields = @(
         # ENERGY USE - CURRENT ELECTRICITY TARIFF
         @{
             id = "single_day_rate"
             label = "Single / Day Rate (pence per kWh)"
             cellReference = "H19"
             type = "number"
             required = $true
         },
         @{
             id = "night_rate"
             label = "Night Rate (pence per kWh)"
             cellReference = "H20"
             type = "number"
             required = $false
         },
         @{
             id = "off_peak_hours"
             label = "No. of Off-Peak Hours"
             cellReference = "H21"
             type = "number"
             required = $false
         },
         
         # ENERGY USE - NEW ELECTRICITY TARIFF
         @{
             id = "new_day_rate"
             label = "Day Rate (pence per kWh)"
             cellReference = "H24"
             type = "number"
             required = $false
         },
         @{
             id = "new_night_rate"
             label = "Night Rate (pence per kWh)"
             cellReference = "H25"
             type = "number"
             required = $false
         },
         
         # ENERGY USE - ELECTRICITY CONSUMPTION
         @{
             id = "annual_usage"
             label = "Estimated Annual Usage (kWh)"
             cellReference = "H27"
             type = "number"
             required = $false
         },
         @{
             id = "standing_charge"
             label = "Standing Charge (pence per day)"
             cellReference = "H28"
             type = "number"
             required = $false
         },
         @{
             id = "annual_spend"
             label = "Annual Spend (£)"
             cellReference = "H29"
             type = "number"
             required = $false
         },
         
         # ENERGY USE - EXPORT TARIFF
         @{
             id = "export_tariff_rate"
             label = "Export Tariff Rate (pence per kWh)"
             cellReference = "H31"
             type = "number"
             required = $false
         },
         
         # EXISTING SYSTEM
         @{
             id = "existing_sem"
             label = "Existing SEM"
             cellReference = "H36"
             type = "text"
             required = $false
         },
         @{
             id = "commissioning_date"
             label = "Approximate Commissioning Date"
             cellReference = "H37"
             type = "text"
             required = $false
         },
         @{
             id = "sem_percentage"
             label = "Percentage of above SEM used to quote self-consumption savings"
             cellReference = "H38"
             type = "number"
             required = $false
         },
         
         # NEW SYSTEM - SOLAR
         @{
             id = "panel_manufacturer"
             label = "Panel Manufacturer"
             cellReference = "H42"
             type = "text"
             required = $false
         },
         @{
             id = "panel_model"
             label = "Panel Model"
             cellReference = "H43"
             type = "text"
             required = $false
         },
         @{
             id = "no_of_arrays"
             label = "No. of Arrays"
             cellReference = "H44"
             type = "number"
             required = $false
         },
         
         # NEW SYSTEM - BATTERY
         @{
             id = "battery_manufacturer"
             label = "Battery Manufacturer"
             cellReference = "H48"
             type = "text"
             required = $false
         },
         @{
             id = "battery_model"
             label = "Battery Model"
             cellReference = "H49"
             type = "text"
             required = $false
         },
         @{
             id = "battery_warranty_years"
             label = "Battery Warranty Period (years)"
             cellReference = "H50"
             type = "number"
             required = $false
         },
         @{
             id = "battery_extended_warranty_years"
             label = "Battery Extended Warranty Period (years)"
             cellReference = "H52"
             type = "number"
             required = $false
         },
         @{
             id = "battery_replacement_cost"
             label = "Battery Replacement Cost (£)"
             cellReference = "H53"
             type = "number"
             required = $false
         },
         
         # NEW SYSTEM - SOLAR/HYBRID INVERTER
         @{
             id = "solar_inverter_manufacturer"
             label = "Solar/Hybrid Inverter Manufacturer"
             cellReference = "H57"
             type = "text"
             required = $false
         },
         @{
             id = "solar_inverter_model"
             label = "Solar/Hybrid Inverter Model"
             cellReference = "H58"
             type = "text"
             required = $false
         },
         @{
             id = "solar_inverter_warranty_years"
             label = "Solar Inverter Warranty Period (years)"
             cellReference = "H59"
             type = "number"
             required = $false
         },
         @{
             id = "solar_inverter_extended_warranty_years"
             label = "Solar Inverter Extended Warranty Period (years)"
             cellReference = "H61"
             type = "number"
             required = $false
         },
         @{
             id = "solar_inverter_replacement_cost"
             label = "Solar Inverter Replacement Cost (£)"
             cellReference = "H62"
             type = "number"
             required = $false
         },
         
         # NEW SYSTEM - BATTERY INVERTER
         @{
             id = "battery_inverter_manufacturer"
             label = "Battery Inverter Manufacturer"
             cellReference = "H66"
             type = "text"
             required = $false
         },
         @{
             id = "battery_inverter_model"
             label = "Battery Inverter Model"
             cellReference = "H67"
             type = "text"
             required = $false
         },
         @{
             id = "battery_inverter_warranty_years"
             label = "Battery Inverter Warranty Period (years)"
             cellReference = "H68"
             type = "number"
             required = $false
         },
         @{
             id = "battery_inverter_extended_warranty_years"
             label = "Battery Inverter Extended Warranty Period (years)"
             cellReference = "H70"
             type = "number"
             required = $false
         },
         @{
             id = "battery_inverter_replacement_cost"
             label = "Battery Inverter Replacement Cost (£)"
             cellReference = "H71"
             type = "number"
             required = $false
         }
     )
    
    Write-Host "Analyzing input fields..." -ForegroundColor Green
    
    $enabledFields = @()
    
         foreach ($field in $inputFields) {
         try {
             $cell = $worksheet.Range($field.cellReference)
             
                     # Check if cell is locked (disabled) - only show enabled fields
        $isLocked = $cell.Locked
        $isWorksheetProtected = $worksheet.ProtectContents
        
        # A cell is enabled if:
        # - Worksheet is not protected (all cells enabled), OR
        # - Worksheet is protected but cell is not locked
        $isEnabled = (-not $isWorksheetProtected) -or (-not $isLocked)
             
             # Get current value
             $currentValue = $cell.Value
             if ($null -eq $currentValue) {
                 $currentValue = ""
             }
             
             # Create field object - only show ENABLED fields
             $fieldObj = @{
                 id = $field.id
                 label = $field.label
                 type = $field.type
                 enabled = $isEnabled  # Use actual enabled status
                 cellReference = $field.cellReference
                 required = $field.required
                 value = $currentValue.ToString()
             }
             
             # Add dropdown options for specific fields
             if ($field.id -eq "panel_manufacturer") {
                 $fieldObj.dropdownOptions = @("LG", "Panasonic", "SunPower", "Canadian Solar", "Jinko Solar", "Trina Solar", "Other")
             } elseif ($field.id -eq "battery_manufacturer") {
                 $fieldObj.dropdownOptions = @("Tesla", "LG Chem", "Sonnen", "Enphase", "Pylontech", "Other")
             } elseif ($field.id -eq "solar_inverter_manufacturer") {
                 $fieldObj.dropdownOptions = @("SMA", "Fronius", "SolarEdge", "Enphase", "Growatt", "Other")
             } elseif ($field.id -eq "battery_inverter_manufacturer") {
                 $fieldObj.dropdownOptions = @("SMA", "Fronius", "SolarEdge", "Enphase", "Growatt", "Other")
             }
             
             # Only add ENABLED fields to the list
             if ($isEnabled) {
                 $enabledFields += $fieldObj
                 Write-Host "✅ ADDED: $($field.label) (Locked: $isLocked, Protected: $isWorksheetProtected)" -ForegroundColor Green
             } else {
                 Write-Host "❌ SKIPPED: $($field.label) (Locked: $isLocked, Protected: $isWorksheetProtected)" -ForegroundColor Red
             }
             
             Write-Host "FIELD:$($field.id)|$($field.label)|$($field.type)|$isEnabled|$($field.cellReference)|$($field.required)|$currentValue" -ForegroundColor Cyan
             
         } catch {
             Write-Host "Error analyzing field $($field.id): $_" -ForegroundColor Yellow
         }
     }
    
         # Output JSON result - ONLY enabled fields
     $result = @{
         inputFields = $enabledFields
         totalFields = $enabledFields.Count
         enabledFields = $enabledFields.Count
     }
     
     Write-Host "DEBUG: Found $($enabledFields.Count) enabled fields out of $($inputFields.Count) total fields" -ForegroundColor Magenta
     
     # Convert to JSON with proper encoding
     $jsonResult = $result | ConvertTo-Json -Depth 3 -Compress
     Write-Host "RESULT: $jsonResult" -ForegroundColor Green
    
    # Close Excel
    try {
        $workbook.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Dynamic inputs retrieval completed successfully!" -ForegroundColor Green
    exit 0
    
} catch {
    Write-Host "Critical error in dynamic inputs retrieval: $_" -ForegroundColor Red
    exit 1
}
`;
  }

  private createSaveDynamicInputsScript(excelFilePath: string, inputs: Record<string, string>, calculatorType?: 'flux' | 'off-peak'): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    // Convert inputs to PowerShell format - fix the hashtable syntax
    const inputsString = Object.entries(inputs)
      .map(([key, value]) => {
        const cleanedValue = this.escapeForPowerShell(value);
        return `    "${key}" = "${cleanedValue}"`;
      })
      .join('\n');
    
         // Define cell mappings for ALL input fields - MUST MATCH ALL_INPUT_FIELDS
     const cellMappings = {
       // Customer Details (always enabled)
       customer_name: 'H12',
       address: 'H13',
       postcode: 'H14',
       
       // ENERGY USE - CURRENT ELECTRICITY TARIFF
       single_day_rate: 'H19',
       night_rate: 'H20',
       off_peak_hours: 'H21',
       
       // ENERGY USE - NEW ELECTRICITY TARIFF
       new_day_rate: 'H23',
       new_night_rate: 'H24',
       
       // ENERGY USE - ELECTRICITY CONSUMPTION
       // Flux calculator uses H26-H32 for consumption fields
       annual_usage: calculatorType === 'flux' ? undefined : 'H26',
       estimated_annual_usage: calculatorType === 'flux' ? 'H26' : undefined,
       estimated_peak_annual_usage: calculatorType === 'flux' ? 'H27' : undefined,
       estimated_off_peak_usage: calculatorType === 'flux' ? 'H28' : undefined,
       standing_charge: calculatorType === 'flux' ? undefined : 'H27',
       standing_charges: calculatorType === 'flux' ? 'H29' : undefined,
       annual_spend: calculatorType === 'flux' ? undefined : 'H28',
       total_annual_spend: calculatorType === 'flux' ? 'H30' : undefined,
       peak_annual_spend: calculatorType === 'flux' ? 'H31' : undefined,
       off_peak_annual_spend: calculatorType === 'flux' ? 'H32' : undefined,
       
       // ENERGY USE - EXPORT TARIFF
       export_tariff_rate: calculatorType === 'flux' ? undefined : 'H30',
       
       // EXISTING SYSTEM
       existing_sem: 'H34',
       commissioning_date: 'H35',
       sem_percentage: 'H36',
       
       // NEW SYSTEM - SOLAR
       panel_manufacturer: 'H41',
       panel_model: 'H42',
       no_of_arrays: 'H43',
       
       // NEW SYSTEM - BATTERY
       battery_manufacturer: 'H45',
       battery_model: 'H46',
       battery_extended_warranty_period: 'H49',
       battery_replacement_cost: 'H50',
       
       // NEW SYSTEM - SOLAR/HYBRID INVERTER
       solar_inverter_manufacturer: 'H52',
       solar_inverter_model: 'H53',
       solar_inverter_extended_warranty_period: 'H56',
       solar_inverter_replacement_cost: 'H57',
       
       // NEW SYSTEM - BATTERY INVERTER
       battery_inverter_manufacturer: 'H59',
       battery_inverter_model: 'H60',
       battery_inverter_extended_warranty_period: 'H63',
       battery_inverter_replacement_cost: 'H64',
       
       // NEW PRODUCTS - SOLAR (Arrays 1-8)
       // Array 1
       array_1_num_panels: 'C69',
       array_1_panel_size_wp: 'D69',
       array_1_array_size_kwp: 'E69',
       array_1_orientation_deg_from_south: 'F69',
       array_1_pitch_deg_from_flat: 'G69',
       array_1_irradiance_kk: 'H69',
       array_1_shading_factor: 'I69',
       
       // Array 2
       array_2_num_panels: 'C70',
       array_2_panel_size_wp: 'D70',
       array_2_array_size_kwp: 'E70',
       array_2_orientation_deg_from_south: 'F70',
       array_2_pitch_deg_from_flat: 'G70',
       array_2_irradiance_kk: 'H70',
       array_2_shading_factor: 'I70',
       
       // Array 3
       array_3_num_panels: 'C71',
       array_3_panel_size_wp: 'D71',
       array_3_array_size_kwp: 'E71',
       array_3_orientation_deg_from_south: 'F71',
       array_3_pitch_deg_from_flat: 'G71',
       array_3_irradiance_kk: 'H71',
       array_3_shading_factor: 'I71',
       
       // Array 4
       array_4_num_panels: 'C72',
       array_4_panel_size_wp: 'D72',
       array_4_array_size_kwp: 'E72',
       array_4_orientation_deg_from_south: 'F72',
       array_4_pitch_deg_from_flat: 'G72',
       array_4_irradiance_kk: 'H72',
       array_4_shading_factor: 'I72',
       
       // Array 5
       array_5_num_panels: 'C73',
       array_5_panel_size_wp: 'D73',
       array_5_array_size_kwp: 'E73',
       array_5_orientation_deg_from_south: 'F73',
       array_5_pitch_deg_from_flat: 'G73',
       array_5_irradiance_kk: 'H73',
       array_5_shading_factor: 'I73',
       
       // Array 6
       array_6_num_panels: 'C74',
       array_6_panel_size_wp: 'D74',
       array_6_array_size_kwp: 'E74',
       array_6_orientation_deg_from_south: 'F74',
       array_6_pitch_deg_from_flat: 'G74',
       array_6_irradiance_kk: 'H74',
       array_6_shading_factor: 'I74',
       
       // Array 7
       array_7_num_panels: 'C75',
       array_7_panel_size_wp: 'D75',
       array_7_array_size_kwp: 'E75',
       array_7_orientation_deg_from_south: 'F75',
       array_7_pitch_deg_from_flat: 'G75',
       array_7_irradiance_kk: 'H75',
       array_7_shading_factor: 'I75',
       
      // Array 8
      array_8_num_panels: 'C76',
      array_8_panel_size_wp: 'D76',
      array_8_array_size_kwp: 'E76',
      array_8_orientation_deg_from_south: 'F76',
      array_8_pitch_deg_from_flat: 'G76',
      array_8_irradiance_kk: 'H76',
      array_8_shading_factor: 'I76',
      
      // PRICING FIELDS - Dynamic based on calculator type
      total_system_cost: calculatorType === 'flux' ? 'H81' : 'H80',    // EPVS: H81, Off Peak: H80
      deposit: calculatorType === 'flux' ? 'H82' : 'H81',              // EPVS: H82, Off Peak: H81
      interest_rate: calculatorType === 'flux' ? 'H83' : 'H82',        // EPVS: H83, Off Peak: H82
      interest_rate_type: calculatorType === 'flux' ? 'H84' : 'H83',   // EPVS: H84, Off Peak: H83
      payment_term: calculatorType === 'flux' ? 'H85' : 'H84'          // EPVS: H85, Off Peak: H84
    };
    
    // Convert cell mappings to PowerShell format - filter out undefined values
    const cellMappingsString = Object.entries(cellMappings)
      .filter(([key, value]) => value !== undefined)
      .map(([key, value]) => `    "${key}" = "${value}"`)
      .join('\n');
    
    return `
# Save Dynamic Inputs to Excel
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePathEscaped}"
$password = "${this.PASSWORD}"

# Input values
$inputs = @{
${inputsString}
}

# Cell mappings
$cellMappings = @{
${cellMappingsString}
}

Write-Host "Saving dynamic inputs to: $filePath" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($filePath)
            Write-Host "Workbook opened without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook: $_" -ForegroundColor Red
            throw "Could not open workbook: $filePath"
        }
    }
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    if (!$worksheet) {
        throw "Inputs worksheet not found"
    }
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
    # Check if worksheet is protected and unprotect if needed
    if ($worksheet.ProtectContents) {
        Write-Host "Worksheet is protected, attempting to unprotect..." -ForegroundColor Yellow
        try {
            $worksheet.Unprotect($password)
            Write-Host "Successfully unprotected worksheet" -ForegroundColor Green
        } catch {
            Write-Host "Failed to unprotect with password, trying without password..." -ForegroundColor Yellow
            try {
                $worksheet.Unprotect()
                Write-Host "Successfully unprotected worksheet without password" -ForegroundColor Green
            } catch {
                Write-Host "Warning: Could not unprotect worksheet: $_" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "Worksheet is not protected" -ForegroundColor Green
    }
    
    # Save input values to Excel cells with array-by-array processing
    Write-Host "Saving input values to Excel cells..." -ForegroundColor Green
    Write-Host ""
    $savedCount = 0
    
    # Check if this is an array-focused save (has array data)
    $hasArrayData = $false
    Write-Host "DEBUG: Checking for array data in inputs..." -ForegroundColor Yellow
    foreach ($key in $inputs.Keys) {
        Write-Host "DEBUG: Input key: $key" -ForegroundColor Yellow
        if ($key -match "^array_[0-9]+_") {
            Write-Host "DEBUG: Found array key: $key" -ForegroundColor Green
            $hasArrayData = $true
            break
        }
    }
    Write-Host "DEBUG: hasArrayData = $hasArrayData" -ForegroundColor Yellow
    
    # Step 1: Process no_of_arrays first (if present) - also check for number_of_arrays
    $noOfArraysValue = $null
    if ($inputs.ContainsKey("no_of_arrays") -and -not [string]::IsNullOrWhiteSpace($inputs["no_of_arrays"])) {
        $noOfArraysValue = $inputs["no_of_arrays"]
    } elseif ($inputs.ContainsKey("number_of_arrays") -and -not [string]::IsNullOrWhiteSpace($inputs["number_of_arrays"])) {
        $noOfArraysValue = $inputs["number_of_arrays"]
    }
    
    if ($noOfArraysValue -ne $null) {
        $cellReference = $cellMappings["no_of_arrays"]
        
        Write-Host "=== STEP 1: Processing no_of_arrays first ===" -ForegroundColor Magenta
        
        try {
            $cell = $worksheet.Range($cellReference)
            
            # Always trigger VBA when setting no_of_arrays to unlock array cells
            # This is needed even if there's no array data yet - it unlocks the cells for future input
            Write-Host "Processing no_of_arrays dropdown with special VBA triggering..." -ForegroundColor Cyan
            
            # VBA triggering path - this unlocks the array cells
            try {
                    Write-Host "Triggering VBA to unlock array cells for no_of_arrays..." -ForegroundColor Cyan
                    
                    # Enable events specifically for this dropdown to trigger VBA
                    $excel.EnableEvents = $true
                    Write-Host "Enabled Excel events for no_of_arrays VBA triggering" -ForegroundColor Cyan
                    
                    # Clear the cell first
                    $cell.Value2 = $null
                    Start-Sleep -Milliseconds 100
                    
                    # Set the value as string (dropdown values are strings)
                    $cell.Value2 = [string]$noOfArraysValue
                    Write-Host "Set no_of_arrays = '$noOfArraysValue' (as string) to cell $cellReference" -ForegroundColor Green
                    
                    # Select the cell to simulate user interaction
                    $cell.Select()
                    Start-Sleep -Milliseconds 100
                    
                    # Force calculations to trigger VBA
                    $excel.Calculate()
                    $excel.CalculateFullRebuild()
                    Start-Sleep -Milliseconds 500
                    
                    # Trigger worksheet change event by selecting and re-setting
                    $worksheet.Range($cellReference).Select()
                    $worksheet.Range($cellReference).Value2 = [string]$noOfArraysValue
                    
                    # Wait for VBA to complete and unlock cells
                    Start-Sleep -Milliseconds 1000
                    
                    # Disable events again for other operations
                    $excel.EnableEvents = $false
                    Write-Host "Disabled Excel events after VBA processing" -ForegroundColor Cyan
                    
                    Write-Host "VBA triggering completed for no_of_arrays - array cells should now be unlocked" -ForegroundColor Green
                    $savedCount++
                    
                } catch {
                    $errorMessage = $_.Exception.Message
                    Write-Host "Warning: Error during no_of_arrays VBA triggering : $errorMessage" -ForegroundColor Yellow
                    # Ensure events are disabled even on error
                    try { $excel.EnableEvents = $false } catch {}
                    # Fallback to simple assignment
                    $cell.Value2 = [string]$noOfArraysValue
                    Write-Host "Fallback: Saved no_of_arrays = $noOfArraysValue (as string) to cell $cellReference" -ForegroundColor Green
                    $savedCount++
                }
            
        } catch {
            Write-Host "Error saving no_of_arrays to cell $cellReference : $_" -ForegroundColor Red
        }
    }
    
    # Step 2: Process arrays one by one (array_1, then array_2, etc.)
    $maxArrays = 8
    for ($arrayNum = 1; $arrayNum -le $maxArrays; $arrayNum++) {
        $arrayFields = @()
        $hasArrayData = $false
        
        # Collect all fields for this array
        foreach ($inputKey in $inputs.Keys) {
            if ($inputKey -match "^array_$arrayNum" + "_") {
                $arrayFields += @{Key = $inputKey; Value = $inputs[$inputKey]}
                $hasArrayData = $true
            }
        }
        
        if ($hasArrayData) {
            Write-Host ""
            Write-Host ("=== STEP 2." + $arrayNum + ": Processing Array " + $arrayNum + " ===") -ForegroundColor Magenta
            
            # Sort array fields in logical order: num_panels, orientation, pitch, shading
            $sortedFields = $arrayFields | Sort-Object {
                switch -Regex ($_.Key) {
                    "_num_panels$" { return 1 }
                    "_orientation_" { return 2 }
                    "_pitch_" { return 3 }
                    "_shading_" { return 4 }
                    default { return 5 }
                }
            }
            
            foreach ($fieldData in $sortedFields) {
                $inputKey = $fieldData.Key
                $inputValue = $fieldData.Value
                $cellReference = $cellMappings[$inputKey]
                
                if ($cellReference) {
                    try {
                        $cell = $worksheet.Range($cellReference)
                        
                        # Check if cell is locked
                        if ($cell.Locked) {
                            Write-Host "Warning: Cell $cellReference is locked, skipping input $inputKey" -ForegroundColor Yellow
                            continue
                        }
                        
                        # Set as string for array fields
                        $cell.Value2 = [string]$inputValue
                        Write-Host "Saved $inputKey = $inputValue (as string) to cell $cellReference" -ForegroundColor Green
                        $savedCount++
                        
                        # Small delay between array field entries
                        Start-Sleep -Milliseconds 50
                        
                    } catch {
                        Write-Host "Error saving $inputKey to cell $cellReference : $_" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Warning: No cell mapping found for input: $inputKey" -ForegroundColor Yellow
                }
            }
        }
    }
    
    # Step 3: Process all non-array fields
    Write-Host ""
    Write-Host "=== STEP 3: Processing Non-Array Fields ===" -ForegroundColor Magenta
    
    foreach ($inputKey in $inputs.Keys) {
        $inputValue = $inputs[$inputKey]
        
        # Skip no_of_arrays and array fields (already processed)
        if ($inputKey -eq "no_of_arrays" -or $inputKey -match "^array_\d+_") {
            continue
        }
        
        $cellReference = $cellMappings[$inputKey]
        
        if ($cellReference) {
            try {
                $cell = $worksheet.Range($cellReference)
                
                # Check if cell is locked
                if ($cell.Locked) {
                    Write-Host "Warning: Cell $cellReference is locked, skipping input $inputKey" -ForegroundColor Yellow
                    continue
                }
                
                # Regular field processing
                # Convert value based on expected type
                $convertedValue = $inputValue
                
                # Identify dropdown fields (manufacturer, model fields)
                $dropdownFields = @("panel_manufacturer", "panel_model", "battery_manufacturer", "battery_model", 
                                   "solar_inverter_manufacturer", "solar_inverter_model", 
                                   "battery_inverter_manufacturer", "battery_inverter_model")
                $isDropdownField = $false
                foreach ($dropdownField in $dropdownFields) {
                    if ($inputKey -eq $dropdownField) {
                        $isDropdownField = $true
                        break
                    }
                }
                
                # Determine if this should be a number or string based on field type
                $numericFields = @("rate", "hours", "usage", "charge", "spend", "sem", "percentage", "annual", "cost", "period")
                $isNumericField = $false
                
                foreach ($numericPattern in $numericFields) {
                    if ($inputKey -match $numericPattern) {
                        $isNumericField = $true
                        break
                    }
                }
                
                if ($isDropdownField) {
                    # Special handling for dropdown fields - enable events, clear cell, set value, select cell
                    try {
                        Write-Host "Processing dropdown field $inputKey with special handling..." -ForegroundColor Cyan
                        
                        # Enable events to trigger VBA if needed
                        $excel.EnableEvents = $true
                        
                        # Clear the cell first
                        $cell.Value2 = $null
                        Start-Sleep -Milliseconds 100
                        
                        # Set the value as string (dropdown values must match options exactly)
                        $cell.Value2 = [string]$inputValue
                        Write-Host "Set dropdown $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                        
                        # Select the cell to simulate user interaction and trigger VBA if needed
                        $cell.Select()
                        Start-Sleep -Milliseconds 100
                        
                        # Force calculations to trigger VBA
                        $excel.Calculate()
                        Start-Sleep -Milliseconds 200
                        
                        # Disable events again for other operations
                        $excel.EnableEvents = $false
                        Write-Host "Dropdown field $inputKey saved successfully" -ForegroundColor Green
                        $savedCount++
                    } catch {
                        $errorMessage = $_.Exception.Message
                        Write-Host "Warning: Error during dropdown processing for $inputKey : $errorMessage" -ForegroundColor Yellow
                        # Ensure events are disabled even on error
                        try { $excel.EnableEvents = $false } catch {}
                        # Fallback to simple assignment
                        $cell.Value2 = [string]$inputValue
                        Write-Host "Fallback: Saved dropdown $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                        $savedCount++
                    }
                } elseif ($isNumericField -and $inputValue -ne "" -and $inputValue -ne $null) {
                    # Try to convert to number for numeric fields
                    try {
                        $convertedValue = [double]$inputValue
                        $cell.Value2 = $convertedValue
                        Write-Host "Saved $inputKey = $convertedValue (as number) to cell $cellReference" -ForegroundColor Green
                    } catch {
                        Write-Host "Warning: Could not convert '$inputValue' to number for $inputKey, setting as string" -ForegroundColor Yellow
                        $cell.Value2 = [string]$inputValue
                        Write-Host "Saved $inputKey = $inputValue (as string) to cell $cellReference" -ForegroundColor Green
                    }
                    $savedCount++
                } else {
                    # Set as string for text fields (address, etc.)
                    $cell.Value2 = [string]$inputValue
                    Write-Host "Saved $inputKey = $inputValue (as string) to cell $cellReference" -ForegroundColor Green
                    $savedCount++
                }
                
            } catch {
                Write-Host "Error saving $inputKey to cell $cellReference : $_" -ForegroundColor Red
            }
        } else {
            Write-Host "Warning: No cell mapping found for input $inputKey" -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
    Write-Host "Successfully saved $savedCount out of $($inputs.Count) input values" -ForegroundColor Green
    
    
    # Protect the worksheet with password
    try {
        $worksheet.Protect($password, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true)
        Write-Host "Worksheet protected with password successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not protect worksheet: $_" -ForegroundColor Yellow
    }
    
    # Save the workbook
    try {
        $workbook.Save()
        Write-Host "Workbook saved successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not save workbook: $_" -ForegroundColor Yellow
    }
    
    # Close Excel
    try {
        $workbook.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Dynamic inputs save completed successfully!" -ForegroundColor Green
    exit 0
    
} catch {
    Write-Host "Critical error in dynamic inputs save: $_" -ForegroundColor Red
    exit 1
}
`;
  }

  async performCompleteCalculation(
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string,
    existingFileName?: string
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string }> {
    this.logger.log(`Performing complete calculation for opportunity: ${opportunityId}${templateFileName ? ` with template: ${templateFileName}` : ''}`);

    try {
      // Check if template file exists
      const templateFilePath = this.getTemplateFilePath(templateFileName);
      if (!fs.existsSync(templateFilePath)) {
        const error = `Template file not found at: ${templateFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Check if we're on Windows (required for Excel COM automation)
      if (process.platform !== 'win32') {
        const error = 'Excel automation requires Windows platform';
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Create opportunities folder if it doesn't exist
      if (!fs.existsSync(this.OPPORTUNITIES_FOLDER)) {
        fs.mkdirSync(this.OPPORTUNITIES_FOLDER, { recursive: true });
        this.logger.log(`Created opportunities folder: ${this.OPPORTUNITIES_FOLDER}`);
      }

      // Create PowerShell script for complete operation
      const psScript = this.createCompleteCalculationScript(opportunityId, customerDetails, radioButtonSelections, dynamicInputs, templateFileName, existingFileName);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-complete-calculation-${Date.now()}.ps1`);
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
        const filePath = this.getOpportunityFilePath(opportunityId, 'off-peak');
        this.logger.log(`Successfully completed calculation: ${filePath}`);
        return {
          success: true,
          message: `Successfully completed calculation for ${opportunityId}`,
          filePath
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to complete calculation: ${opportunityId}`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in performCompleteCalculation: ${error.message}`);
      return {
        success: false,
        message: `Error completing calculation: ${opportunityId}`,
        error: error.message
      };
    }
  }

  private createCompleteCalculationScript(
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string,
    existingFileName?: string
  ): string {
    // Determine source file and target file path
    let sourceFilePath: string;
    let targetFilePath: string;
    
    if (existingFileName) {
      // Editing existing file: use the existing file directly (don't create a new version)
      const existingFilePath = path.join(this.OPPORTUNITIES_FOLDER, existingFileName);
      this.logger.log(`🔍 Editing mode: existingFilePath=${existingFilePath}`);
      targetFilePath = existingFilePath; // Edit the existing file directly
      sourceFilePath = existingFilePath; // Also use as source (no copy needed when editing)
      this.logger.log(`📝 Editing existing file: ${existingFileName} -> will update file directly`);
    } else {
      // Creating new file: copy from template
      sourceFilePath = this.getTemplateFilePath(templateFileName);
      const baseFileName = `Off peak V2.1 Eon SEG-${opportunityId}`;
      targetFilePath = this.getVersionedFilePath(this.OPPORTUNITIES_FOLDER, baseFileName, 'xlsm');
    }
    
    const templatePath = sourceFilePath.replace(/\\/g, '\\\\');
    const newFilePath = targetFilePath.replace(/\\/g, '\\\\');
    
    return `
# Complete Calculation - Single Excel Session
$ErrorActionPreference = "Stop"

# Configuration
$templatePath = "${templatePath}"
$newFilePath = "${newFilePath}"
$password = "${this.PASSWORD}"
$opportunityId = "${opportunityId}"

# Customer details
$customerName = "${this.escapeForPowerShell(customerDetails.customerName)}"
$address = "${this.escapeForPowerShell(customerDetails.address)}"
$postcode = "${this.escapeForPowerShell(customerDetails.postcode)}"

# Radio button selections
$radioButtonSelections = @(${radioButtonSelections.map(shape => `"${this.escapeForPowerShell(shape)}"`).join(', ')})

# Dynamic inputs (if any)
$dynamicInputs = @{
${dynamicInputs && Object.keys(dynamicInputs).length > 0 ? Object.entries(dynamicInputs).map(([key, value]) => `    "${this.escapeForPowerShell(key)}" = "${this.escapeForPowerShell(String(value || ''))}"`).join('\n') : '    # No dynamic inputs'}
}

Write-Host "Starting complete calculation for: $opportunityId" -ForegroundColor Green
Write-Host "Source file path: $templatePath" -ForegroundColor Yellow
Write-Host "New file path: $newFilePath" -ForegroundColor Yellow

try {
    # STEP 1: Create the file by copying source file (template) OR prepare existing file (editing)
    ${existingFileName ? `Write-Host "Step 1: Opening existing file for editing..." -ForegroundColor Green` : `Write-Host "Step 1: Creating file by copying template..." -ForegroundColor Green`}
    
    # Kill any existing Excel processes that might have the file locked
    Write-Host "Checking for existing Excel processes that might lock the file..." -ForegroundColor Yellow
    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        Write-Host "Found existing Excel processes, terminating them..." -ForegroundColor Yellow
        $excelProcesses | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "Excel processes terminated" -ForegroundColor Green
    }
    
    # Check if target file exists and is accessible
    Write-Host "Checking target file accessibility..." -ForegroundColor Yellow
    if (-not (Test-Path $newFilePath)) {
        # File doesn't exist, need to copy from template
        if (-not (Test-Path $templatePath)) {
            throw "Source file does not exist: $templatePath"
        }
        
        # Try to get source file info
        try {
            $fileInfo = Get-Item $templatePath
            Write-Host "Source file size: $($fileInfo.Length) bytes" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not get source file info: $_" -ForegroundColor Yellow
        }
        
        # Copy the source file to create the new opportunity file
        Write-Host "Attempting to copy source file..." -ForegroundColor Yellow
        try {
            Copy-Item -Path $templatePath -Destination $newFilePath -Force
            Write-Host "Source file copied successfully" -ForegroundColor Green
            
            # Update file modification date to current time
            Write-Host "Updating file modification date to current time..." -ForegroundColor Yellow
            try {
                $file = Get-Item $newFilePath
                $currentDate = Get-Date
                $file.LastWriteTime = $currentDate
                $file.LastAccessTime = $currentDate
                Write-Host "File modification date updated to: $currentDate" -ForegroundColor Green
            } catch {
                Write-Host "Warning: Could not update file modification date: $_" -ForegroundColor Yellow
            }
            
            Write-Host "New opportunity file created at: $newFilePath" -ForegroundColor Green
        } catch {
            Write-Host "Failed to copy source file: $_" -ForegroundColor Red
            throw "Could not copy source file"
        }
    } else {
        # File already exists (editing mode) - just verify it's accessible
        Write-Host "File already exists: $newFilePath (editing mode)" -ForegroundColor Green
        try {
            $fileInfo = Get-Item $newFilePath
            Write-Host "Target file size: $($fileInfo.Length) bytes" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not get target file info: $_" -ForegroundColor Yellow
        }
    }
    
    # Step 1 completed - file created successfully
    Write-Host "Step 1 completed: File created successfully at $newFilePath" -ForegroundColor Green
    
    # STEP 2: Open Excel, unprotect, and add customer details
    Write-Host "Step 2: Opening Excel and adding customer details..." -ForegroundColor Green
    
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open the copied file
    Write-Host "Opening workbook: $newFilePath" -ForegroundColor Green
    try {
        $workbook = $excel.Workbooks.Open($newFilePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without password..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($newFilePath)
            Write-Host "Workbook opened successfully without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook: $_" -ForegroundColor Red
            throw "Could not open workbook"
        }
    }
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
    # Check if worksheet is protected and unprotect if needed
    if ($worksheet.ProtectContents) {
        Write-Host "Worksheet is protected, attempting to unprotect..." -ForegroundColor Yellow
        $worksheet.Unprotect($password)
        Write-Host "Successfully unprotected worksheet" -ForegroundColor Green
    }
    
    # Fill in customer details
    Write-Host "Filling in customer details..." -ForegroundColor Green
    
    # Customer Name (H12)
    $worksheet.Range("H12").Value = $customerName
    Write-Host "Set Customer Name: $customerName" -ForegroundColor Green
    
    # Address (H13)
    $worksheet.Range("H13").Value = $address
    Write-Host "Set Address: $address" -ForegroundColor Green
    
    # Postcode (H14) - only if provided
    if ($postcode -ne "") {
        $worksheet.Range("H14").Value = $postcode
        Write-Host "Set Postcode: $postcode" -ForegroundColor Green
    }
    
    Write-Host "Step 2 completed: Customer details added (workbook will be saved at the end)" -ForegroundColor Green
    
    # STEP 3: Select all radio buttons
    Write-Host "Step 3: Selecting radio buttons..." -ForegroundColor Green
    
    # Ensure all worksheets are unprotected for radio button selection
    Write-Host "Unprotecting all worksheets for radio button selection..." -ForegroundColor Yellow
    foreach ($ws in $workbook.Worksheets) {
        try {
            if ($ws.ProtectContents) {
                Write-Host "Unprotecting worksheet: $($ws.Name)" -ForegroundColor Cyan
                try {
                    $ws.Unprotect($password)
                    Write-Host "Successfully unprotected worksheet: $($ws.Name)" -ForegroundColor Green
                } catch {
                    try {
                        $ws.Unprotect()
                        Write-Host "Successfully unprotected worksheet without password: $($ws.Name)" -ForegroundColor Green
                    } catch {
                        Write-Host "Warning: Could not unprotect worksheet $($ws.Name): $_" -ForegroundColor Yellow
                    }
                }
            }
        } catch {
            Write-Host "Error processing worksheet $($ws.Name): $_" -ForegroundColor Yellow
        }
    }
    
    $successfulSelections = 0
    
    # Process each radio button
    foreach ($shapeName in $radioButtonSelections) {
        Write-Host "Processing radio button: $shapeName" -ForegroundColor Yellow
        
        try {
            $targetShape = $worksheet.Shapes($shapeName)
            Write-Host "Found shape: $shapeName" -ForegroundColor Green
            
            # Get shape details
            $onActionMacro = ""
            try {
                $onActionMacro = $targetShape.OnAction
                if ($onActionMacro) {
                    Write-Host "Shape has OnAction macro: $onActionMacro" -ForegroundColor Cyan
                }
            } catch {
                Write-Host "Shape does not have OnAction macro" -ForegroundColor Yellow
            }
            
            # Try to interact with the radio button
            $interactionSuccess = $false
            
            # Method 1: Try ControlFormat if available
            try {
                if ($targetShape.ControlFormat) {
                    Write-Host "Attempting to select radio button using ControlFormat..." -ForegroundColor Yellow
                    $targetShape.ControlFormat.Value = 1
                    Write-Host "Successfully selected radio button using ControlFormat: $shapeName" -ForegroundColor Green
                    $interactionSuccess = $true
                }
            } catch {
                Write-Host "ControlFormat interaction failed: $_" -ForegroundColor Yellow
            }
            
            # Method 2: Try to select the shape
            if (-not $interactionSuccess) {
                try {
                    Write-Host "Attempting to select radio button shape..." -ForegroundColor Yellow
                    $targetShape.Select()
                    Start-Sleep -Milliseconds 100
                    Write-Host "Successfully selected radio button shape: $shapeName" -ForegroundColor Green
                    $interactionSuccess = $true
                } catch {
                    Write-Host "Shape selection failed: $_" -ForegroundColor Yellow
                }
            }
            
            # Method 3: Try to click the shape
            if (-not $interactionSuccess) {
                try {
                    Write-Host "Attempting to click radio button..." -ForegroundColor Yellow
                    $targetShape.Select()
                    Start-Sleep -Milliseconds 100
                    $excel.SendKeys("{ENTER}")
                    Start-Sleep -Milliseconds 100
                    Write-Host "Successfully clicked radio button: $shapeName" -ForegroundColor Green
                    $interactionSuccess = $true
                } catch {
                    Write-Host "Radio button clicking failed: $_" -ForegroundColor Yellow
                }
            }
            
            if ($interactionSuccess) {
                $successfulSelections++
                
                # Try to trigger the OnAction macro if it exists
                    if ($onActionMacro -and $onActionMacro.Trim() -ne "") {
                    try {
                        Write-Host "Executing OnAction macro: $onActionMacro" -ForegroundColor Cyan
                        $excel.Run($onActionMacro)
                        Write-Host "Successfully executed OnAction macro: $onActionMacro" -ForegroundColor Green
                } catch {
                        Write-Host "Failed to execute OnAction macro: $_" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "Failed to interact with radio button: $shapeName" -ForegroundColor Red
            }
            
        } catch {
            Write-Host "Error processing radio button '$shapeName': $_" -ForegroundColor Red
        }
    }
    
    Write-Host "Successfully selected $successfulSelections out of $($radioButtonSelections.Count) radio buttons" -ForegroundColor Green
    
    # Provide additional information about the results
    if ($successfulSelections -eq $radioButtonSelections.Count) {
        Write-Host "All radio buttons were successfully selected!" -ForegroundColor Green
    } elseif ($successfulSelections -gt 0) {
        Write-Host "Some radio buttons were selected successfully. Check logs above for any issues." -ForegroundColor Yellow
    } else {
        Write-Host "No radio buttons were successfully selected. Check logs above for errors." -ForegroundColor Red
    }
    
    Write-Host "Step 3 completed: Radio buttons selected (workbook will be saved at the end)" -ForegroundColor Green
    
    # STEP 4: Input dynamic data (non-solar fields only)
    Write-Host "Step 4: Inputting dynamic data (non-solar fields)..." -ForegroundColor Green
    
    # Define cell mappings for non-solar dynamic inputs
    $cellMappings = @{
        # ENERGY USE - CURRENT ELECTRICITY TARIFF
        "current_single_day_rate" = "H19"
        "current_night_rate" = "H20"
        "current_off_peak_hours" = "H21"
        "single_day_rate" = "H19"
        "night_rate" = "H20"
        "off_peak_hours" = "H21"
        
        # ENERGY USE - NEW ELECTRICITY TARIFF
        "new_day_rate" = "H23"
        "new_night_rate" = "H24"
        
        # ENERGY USE - ELECTRICITY CONSUMPTION
        # Off-peak calculator uses H26-H28 for consumption
        "annual_usage" = "H26"
        "estimated_annual_usage" = "H26"
        "standing_charge" = "H27"
        "annual_spend" = "H28"
        
        # ENERGY USE - EXPORT TARIFF
        "export_tariff_rate" = "H30"
        
        # EXISTING SYSTEM (not solar - these should be input in Step 4)
        "existing_sem" = "H34"
        "commissioning_date" = "H35"
        "approximate_commissioning_date" = "H35"
        "sem_percentage" = "H36"
        "percentage_above_sem" = "H36"
    }
    
    # Define solar-related field prefixes to exclude
    $solarFieldPrefixes = @(
        "panel_", "battery_", "solar_inverter_", "battery_inverter_",
        "array_", "no_of_arrays"
    )
    
    $savedCount = 0
    
    # Process each dynamic input (skip solar-related fields)
    foreach ($inputKey in $dynamicInputs.Keys) {
        # Skip solar-related fields
        $isSolarField = $false
        foreach ($prefix in $solarFieldPrefixes) {
            if ($inputKey -like "$prefix*") {
                $isSolarField = $true
                break
            }
        }
        
        if ($isSolarField) {
            Write-Host "Skipping solar-related field: $inputKey" -ForegroundColor Yellow
            continue
        }
        
        # Skip customer details (already done in Step 2)
        if ($inputKey -eq "customer_name" -or $inputKey -eq "address" -or $inputKey -eq "postcode") {
            continue
        }
        
            $inputValue = $dynamicInputs[$inputKey]
        
        # Skip empty values first (before priority checks)
        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            continue
        }
        
        # Skip variant fields if their preferred equivalent exists AND has a non-empty value (prioritize specific variants)
        if ($inputKey -eq "single_day_rate" -and $dynamicInputs.ContainsKey("current_single_day_rate")) {
            $preferredValue = $dynamicInputs["current_single_day_rate"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because current_single_day_rate has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "night_rate" -and $dynamicInputs.ContainsKey("current_night_rate")) {
            $preferredValue = $dynamicInputs["current_night_rate"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because current_night_rate has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "off_peak_hours" -and $dynamicInputs.ContainsKey("current_off_peak_hours")) {
            $preferredValue = $dynamicInputs["current_off_peak_hours"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because current_off_peak_hours has value" -ForegroundColor Yellow
                continue
            }
        }
        
        # Skip annual_usage if estimated_annual_usage exists AND has a non-empty value (prioritize estimated_annual_usage)
        if ($inputKey -eq "annual_usage" -and $dynamicInputs.ContainsKey("estimated_annual_usage")) {
            $preferredValue = $dynamicInputs["estimated_annual_usage"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because estimated_annual_usage has value" -ForegroundColor Yellow
                continue
            }
        }
        
        # Skip commissioning_date if approximate_commissioning_date exists AND has a non-empty value (prioritize approximate_commissioning_date)
        if ($inputKey -eq "commissioning_date" -and $dynamicInputs.ContainsKey("approximate_commissioning_date")) {
            $preferredValue = $dynamicInputs["approximate_commissioning_date"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because approximate_commissioning_date has value" -ForegroundColor Yellow
                continue
            }
        }
        
        # Skip sem_percentage if percentage_above_sem exists AND has a non-empty value (prioritize percentage_above_sem)
        if ($inputKey -eq "sem_percentage" -and $dynamicInputs.ContainsKey("percentage_above_sem")) {
            $preferredValue = $dynamicInputs["percentage_above_sem"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because percentage_above_sem has value" -ForegroundColor Yellow
                continue
            }
        }
        
            $cellReference = $cellMappings[$inputKey]
            
            if ($cellReference) {
                try {
                    $cell = $worksheet.Range($cellReference)
                    
                # Check if cell is locked - skip it (cells are locked when disabled based on radio button selections)
                    if ($cell.Locked) {
                    Write-Host "Warning: Cell $cellReference is locked (disabled), skipping input $inputKey" -ForegroundColor Yellow
                        continue
                    }
                    
                    # Convert value based on expected type
                $numericFields = @("rate", "hours", "usage", "charge", "charges", "spend", "off_peak", "peak")
                $isNumericField = $false
                
                foreach ($numericPattern in $numericFields) {
                    if ($inputKey -match $numericPattern) {
                        $isNumericField = $true
                        break
                    }
                }
                
                # Check if field is a date field
                $isDateField = $false
                if ($inputKey -match "date|Date|DATE") {
                    $isDateField = $true
                }
                
                # Handle date fields - try to parse and convert to Excel date
                if ($isDateField -and $inputValue -ne "" -and $inputValue -ne $null) {
                    try {
                        # PowerShell DateTime.TryParse requires different syntax
                        $parsedDate = $null
                        if ([DateTime]::TryParse($inputValue, [ref]$parsedDate)) {
                            $dateValue = $parsedDate
                            # Excel date format: number of days since 1899-12-30 (Excel epoch accounting for leap year bug)
                            $excelEpoch = [DateTime]::Parse("1899-12-30")
                            $excelDate = ($dateValue - $excelEpoch).TotalDays
                            $cell.Value2 = $excelDate
                            $cell.NumberFormat = "dd/mm/yyyy"
                            Write-Host "Saved $inputKey = $inputValue (as Excel date: $excelDate) to cell $cellReference" -ForegroundColor Green
                            # Clear any data validation to prevent Excel popups
                            try { $cell.Validation.Delete() } catch { }
                            $savedCount++
                            continue
                        } else {
                            # If parsing fails, save as string
                            Write-Host "Warning: Could not parse date '$inputValue' for $inputKey, saving as string" -ForegroundColor Yellow
                            $cell.Value2 = [string]$inputValue
                            Write-Host "Saved $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                            # Clear any data validation to prevent Excel popups
                            try { $cell.Validation.Delete() } catch { }
                            $savedCount++
                            continue
                        }
                    } catch {
                        Write-Host "Warning: Error parsing date '$inputValue' for $inputKey, saving as string. Error: $_" -ForegroundColor Yellow
                        $cell.Value2 = [string]$inputValue
                        Write-Host "Saved $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                        # Clear any data validation to prevent Excel popups
                        try { $cell.Validation.Delete() } catch { }
                        $savedCount++
                        continue
                    }
                }
                
                if ($isNumericField -and $inputValue -ne "" -and $inputValue -ne $null) {
                    # Try to convert to number for numeric fields
                    try {
                        $cleanValue = $inputValue.ToString().Trim()
                        
                        # Try different conversion methods
                        $convertedValue = 0
                        if ([double]::TryParse($cleanValue, [ref]$convertedValue)) {
                            $cell.Value2 = $convertedValue
                            Write-Host "Saved $inputKey = $convertedValue (as number) to cell $cellReference" -ForegroundColor Green
                        } else {
                            # Try parsing as integer first
                            $intValue = 0
                            if ([int]::TryParse($cleanValue, [ref]$intValue)) {
                                $cell.Value2 = $intValue
                                Write-Host "Saved $inputKey = $intValue (as integer) to cell $cellReference" -ForegroundColor Green
                            } else {
                                # Fallback to string if conversion fails
                                $cell.Value2 = [string]$inputValue
                                Write-Host "Warning: Could not convert '$inputValue' to number for $inputKey, saved as string" -ForegroundColor Yellow
                            }
                        }
                        # Clear any data validation to prevent Excel popups
                        try { $cell.Validation.Delete() } catch { }
                        } catch {
                        Write-Host "Warning: Could not convert '$inputValue' to number for $inputKey, setting as string. Error: $_" -ForegroundColor Yellow
                        $cell.Value2 = [string]$inputValue
                        # Clear any data validation to prevent Excel popups
                        try { $cell.Validation.Delete() } catch { }
                        }
                    } else {
                    # Set as string for text fields
                    $cell.Value2 = [string]$inputValue
                    Write-Host "Saved $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                    # Clear any data validation to prevent Excel popups
                    try { $cell.Validation.Delete() } catch { }
                }
                    $savedCount++
                    
                } catch {
                    Write-Host "Error saving $inputKey to cell $cellReference : $_" -ForegroundColor Red
                }
            } else {
            Write-Host "Warning: No cell mapping found for input $inputKey (skipping)" -ForegroundColor Yellow
        }
    }
    
    Write-Host "Successfully saved $savedCount dynamic input values (non-solar)" -ForegroundColor Green
    
    Write-Host "Step 4 completed: Dynamic inputs (non-solar) added" -ForegroundColor Green
    
    # STEP 5: Input solar system data (panels, batteries, inverters)
    Write-Host "Step 5: Inputting solar system data..." -ForegroundColor Green
    
    # Define cell mappings for solar system fields (Off-Peak)
    $solarCellMappings = @{
        # SOLAR PANEL FIELDS
        "panel_manufacturer" = "H41"
        "panel_model" = "H42"
        "number_of_arrays" = "H43"
        "no_of_arrays" = "H43"
        
        # BATTERY FIELDS
        "battery_manufacturer" = "H45"
        "battery_model" = "H46"
        "battery_extended_warranty_period" = "H49"
        "battery_extended_warranty_years" = "H49"
        "battery_replacement_cost" = "H50"
        
        # SOLAR/HYBRID INVERTER FIELDS
        "solar_inverter_manufacturer" = "H52"
        "solar_inverter_model" = "H53"
        "solar_inverter_extended_warranty_period" = "H56"
        "solar_inverter_extended_warranty_years" = "H56"
        "solar_inverter_replacement_cost" = "H57"
        
        # BATTERY INVERTER FIELDS
        "battery_inverter_manufacturer" = "H59"
        "battery_inverter_model" = "H60"
        "battery_inverter_extended_warranty_period" = "H63"
        "battery_inverter_extended_warranty_years" = "H63"
        "battery_inverter_replacement_cost" = "H64"
    }
    
    # Define solar-related field prefixes to include
    $solarFieldPrefixes = @(
        "panel_", "battery_", "solar_inverter_", "battery_inverter_"
    )
    
    $solarSavedCount = 0
    
    # Process each dynamic input (only solar-related fields)
    foreach ($inputKey in $dynamicInputs.Keys) {
        # Only process solar-related fields
        $isSolarField = $false
        foreach ($prefix in $solarFieldPrefixes) {
            if ($inputKey -like "$prefix*") {
                $isSolarField = $true
                break
            }
        }
        
        # Also check for number_of_arrays/no_of_arrays
        if ($inputKey -eq "number_of_arrays" -or $inputKey -eq "no_of_arrays") {
            $isSolarField = $true
        }
        
        if (-not $isSolarField) {
            continue
        }
        
        $inputValue = $dynamicInputs[$inputKey]
        
        # Skip empty values
        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            continue
        }
        
        # Handle field name aliases (e.g., battery_extended_warranty_years vs battery_extended_warranty_period)
        # Priority: Check for more specific field names first
        if ($inputKey -eq "battery_extended_warranty_period" -and $dynamicInputs.ContainsKey("battery_extended_warranty_years")) {
            $preferredValue = $dynamicInputs["battery_extended_warranty_years"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because battery_extended_warranty_years has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "solar_inverter_extended_warranty_period" -and $dynamicInputs.ContainsKey("solar_inverter_extended_warranty_years")) {
            $preferredValue = $dynamicInputs["solar_inverter_extended_warranty_years"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because solar_inverter_extended_warranty_years has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "battery_inverter_extended_warranty_period" -and $dynamicInputs.ContainsKey("battery_inverter_extended_warranty_years")) {
            $preferredValue = $dynamicInputs["battery_inverter_extended_warranty_years"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because battery_inverter_extended_warranty_years has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "no_of_arrays" -and $dynamicInputs.ContainsKey("number_of_arrays")) {
            $preferredValue = $dynamicInputs["number_of_arrays"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because number_of_arrays has value" -ForegroundColor Yellow
                continue
            }
        }
        
        $cellReference = $solarCellMappings[$inputKey]
        
        if ($cellReference) {
            try {
                $cell = $worksheet.Range($cellReference)
                
                # Check if cell is locked - skip it (cells are locked when disabled based on template/radio selections)
                if ($cell.Locked) {
                    Write-Host "Warning: Cell $cellReference is locked (disabled), skipping input $inputKey" -ForegroundColor Yellow
                    continue
                }
                
                    # Identify dropdown fields (manufacturer, model fields)
                    $dropdownFields = @("panel_manufacturer", "panel_model", "battery_manufacturer", "battery_model", 
                                       "solar_inverter_manufacturer", "solar_inverter_model", 
                                       "battery_inverter_manufacturer", "battery_inverter_model")
                    $isDropdownField = $false
                    foreach ($dropdownField in $dropdownFields) {
                        if ($inputKey -eq $dropdownField) {
                            $isDropdownField = $true
                            break
                        }
                    }
                    
                    # Convert value based on expected type
                $numericFields = @("replacement_cost", "extended_warranty", "warranty_period", "warranty_years", "number_of_arrays", "no_of_arrays")
                $isNumericField = $false
                
                foreach ($numericPattern in $numericFields) {
                    if ($inputKey -match $numericPattern) {
                        $isNumericField = $true
                        break
                    }
                }
                
                if ($isDropdownField) {
                    # Special handling for dropdown fields - enable events, clear cell, set value, select cell
                    try {
                        Write-Host "Processing dropdown field $inputKey with special handling..." -ForegroundColor Cyan
                        
                        # Enable events to trigger VBA if needed
                        $excel.EnableEvents = $true
                        
                        # Clear the cell first
                        $cell.Value2 = $null
                        Start-Sleep -Milliseconds 100
                        
                        # Set the value as string (dropdown values must match options exactly)
                        $cell.Value2 = [string]$inputValue
                        Write-Host "Set dropdown $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                        
                        # Select the cell to simulate user interaction and trigger VBA if needed
                        $cell.Select()
                        Start-Sleep -Milliseconds 100
                        
                        # Force calculations to trigger VBA
                        $excel.Calculate()
                        Start-Sleep -Milliseconds 200
                        
                        # Disable events again for other operations
                        $excel.EnableEvents = $false
                        Write-Host "Dropdown field $inputKey saved successfully" -ForegroundColor Green
                        $solarSavedCount++
                    } catch {
                        $errorMessage = $_.Exception.Message
                        Write-Host "Warning: Error during dropdown processing for $inputKey : $errorMessage" -ForegroundColor Yellow
                        # Ensure events are disabled even on error
                        try { $excel.EnableEvents = $false } catch {}
                        # Fallback to simple assignment
                        $cell.Value2 = [string]$inputValue
                        Write-Host "Fallback: Saved dropdown $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                        $solarSavedCount++
                    }
                } elseif ($isNumericField -and $inputValue -ne "" -and $inputValue -ne $null) {
                    # Try to convert to number for numeric fields
                    try {
                        $cleanValue = $inputValue.ToString().Trim()
                        
                        # Try different conversion methods
                        $convertedValue = 0
                        if ([double]::TryParse($cleanValue, [ref]$convertedValue)) {
                            $cell.Value2 = $convertedValue
                            Write-Host "Saved $inputKey = $convertedValue (as number) to cell $cellReference" -ForegroundColor Green
                        } else {
                            # Try parsing as integer
                            $intValue = 0
                            if ([int]::TryParse($cleanValue, [ref]$intValue)) {
                                $cell.Value2 = $intValue
                                Write-Host "Saved $inputKey = $intValue (as integer) to cell $cellReference" -ForegroundColor Green
                            } else {
                                # Fallback to string if conversion fails
                                $cell.Value2 = [string]$inputValue
                                Write-Host "Warning: Could not convert '$inputValue' to number for $inputKey, saved as string" -ForegroundColor Yellow
                            }
                        }
                        # Clear any data validation to prevent Excel popups
                        try { $cell.Validation.Delete() } catch { }
        } catch {
                        Write-Host "Warning: Could not convert '$inputValue' to number for $inputKey, setting as string. Error: $_" -ForegroundColor Yellow
                        $cell.Value2 = [string]$inputValue
                        # Clear any data validation to prevent Excel popups
                        try { $cell.Validation.Delete() } catch { }
                    }
                    $solarSavedCount++
                } else {
                    # Set as string for text fields (address, etc.)
                        $cell.Value2 = [string]$inputValue
                    Write-Host "Saved $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                    # Clear any data validation to prevent Excel popups
                    try { $cell.Validation.Delete() } catch { }
                    $solarSavedCount++
                }
                
            } catch {
                Write-Host "Error saving $inputKey to cell $cellReference : $_" -ForegroundColor Red
            }
        } else {
            Write-Host "Warning: No cell mapping found for solar field $inputKey (skipping)" -ForegroundColor Yellow
        }
    }
    
    Write-Host "Successfully saved $solarSavedCount solar system field values" -ForegroundColor Green
    
    Write-Host "Step 5 completed: Solar system data added" -ForegroundColor Green
    
    # STEP 6: Input solar array data
    Write-Host "Step 6: Inputting solar array data..." -ForegroundColor Green
    
    # Define cell mappings for array fields (Off-Peak format: array_1_num_panels, array_1_orientation_deg_from_south, etc.)
    $arrayCellMappings = @{
        # ARRAY 1
        "array_1_num_panels" = "C69"
        "array_1_panel_size_wp" = "D69"
        "array_1_array_size_kwp" = "E69"
        "array_1_orientation_deg_from_south" = "F69"
        "array_1_pitch_deg_from_flat" = "G69"
        "array_1_irradiance_kk" = "H69"
        "array_1_shading_factor" = "I69"
        
        # ARRAY 2
        "array_2_num_panels" = "C70"
        "array_2_panel_size_wp" = "D70"
        "array_2_array_size_kwp" = "E70"
        "array_2_orientation_deg_from_south" = "F70"
        "array_2_pitch_deg_from_flat" = "G70"
        "array_2_irradiance_kk" = "H70"
        "array_2_shading_factor" = "I70"
        
        # ARRAY 3
        "array_3_num_panels" = "C71"
        "array_3_panel_size_wp" = "D71"
        "array_3_array_size_kwp" = "E71"
        "array_3_orientation_deg_from_south" = "F71"
        "array_3_pitch_deg_from_flat" = "G71"
        "array_3_irradiance_kk" = "H71"
        "array_3_shading_factor" = "I71"
        
        # ARRAY 4
        "array_4_num_panels" = "C72"
        "array_4_panel_size_wp" = "D72"
        "array_4_array_size_kwp" = "E72"
        "array_4_orientation_deg_from_south" = "F72"
        "array_4_pitch_deg_from_flat" = "G72"
        "array_4_irradiance_kk" = "H72"
        "array_4_shading_factor" = "I72"
        
        # ARRAY 5
        "array_5_num_panels" = "C73"
        "array_5_panel_size_wp" = "D73"
        "array_5_array_size_kwp" = "E73"
        "array_5_orientation_deg_from_south" = "F73"
        "array_5_pitch_deg_from_flat" = "G73"
        "array_5_irradiance_kk" = "H73"
        "array_5_shading_factor" = "I73"
        
        # ARRAY 6
        "array_6_num_panels" = "C74"
        "array_6_panel_size_wp" = "D74"
        "array_6_array_size_kwp" = "E74"
        "array_6_orientation_deg_from_south" = "F74"
        "array_6_pitch_deg_from_flat" = "G74"
        "array_6_irradiance_kk" = "H74"
        "array_6_shading_factor" = "I74"
        
        # ARRAY 7
        "array_7_num_panels" = "C75"
        "array_7_panel_size_wp" = "D75"
        "array_7_array_size_kwp" = "E75"
        "array_7_orientation_deg_from_south" = "F75"
        "array_7_pitch_deg_from_flat" = "G75"
        "array_7_irradiance_kk" = "H75"
        "array_7_shading_factor" = "I75"
        
        # ARRAY 8
        "array_8_num_panels" = "C76"
        "array_8_panel_size_wp" = "D76"
        "array_8_array_size_kwp" = "E76"
        "array_8_orientation_deg_from_south" = "F76"
        "array_8_pitch_deg_from_flat" = "G76"
        "array_8_irradiance_kk" = "H76"
        "array_8_shading_factor" = "I76"
    }
    
    # Check if we have array data that requires VBA triggering for no_of_arrays
    $hasArrayData = $false
    Write-Host "Checking for array data in dynamic inputs..." -ForegroundColor Yellow
    foreach ($inputKey in $dynamicInputs.Keys) {
        if ($inputKey -match "^array_[0-9]+_") {
            Write-Host "Found array key: $inputKey" -ForegroundColor Green
            $hasArrayData = $true
            break
        }
    }
    Write-Host "hasArrayData = $hasArrayData" -ForegroundColor Yellow
    
    $arraySavedCount = 0
    
    # STEP 6.1: Process no_of_arrays FIRST with VBA triggering if array data is present (this unlocks array cells)
    # Count enabled arrays directly from dynamic inputs (more reliable than relying on number_of_arrays/no_of_arrays)
    $enabledArraysCount = 0
    $foundArrayIndices = @{}
    foreach ($inputKey in $dynamicInputs.Keys) {
        if ($inputKey -match "^array_([0-9]+)_") {
            $arrayIndex = [int]$matches[1]
            if (-not $foundArrayIndices.ContainsKey($arrayIndex)) {
                # Check if this array has any non-empty values
                $hasData = $false
                foreach ($checkKey in $dynamicInputs.Keys) {
                    if ($checkKey -match "^array_$arrayIndex" + "_" -and -not [string]::IsNullOrWhiteSpace($dynamicInputs[$checkKey])) {
                        $hasData = $true
                        break
                    }
                }
                if ($hasData) {
                    $foundArrayIndices[$arrayIndex] = $true
                    $enabledArraysCount = [Math]::Max($enabledArraysCount, $arrayIndex)
                }
            }
        }
    }
    
    # Use the count of enabled arrays, or fallback to no_of_arrays/number_of_arrays
    $noOfArraysValue = $null
    if ($enabledArraysCount -gt 0) {
        $noOfArraysValue = [string]$enabledArraysCount
        Write-Host "Counted $enabledArraysCount enabled arrays from dynamic inputs" -ForegroundColor Green
    } elseif ($dynamicInputs.ContainsKey("no_of_arrays") -and -not [string]::IsNullOrWhiteSpace($dynamicInputs["no_of_arrays"])) {
        $noOfArraysValue = $dynamicInputs["no_of_arrays"]
        Write-Host "Using no_of_arrays value: $noOfArraysValue" -ForegroundColor Yellow
    } elseif ($dynamicInputs.ContainsKey("number_of_arrays") -and -not [string]::IsNullOrWhiteSpace($dynamicInputs["number_of_arrays"])) {
        $noOfArraysValue = $dynamicInputs["number_of_arrays"]
        Write-Host "Using number_of_arrays value: $noOfArraysValue" -ForegroundColor Yellow
    }
    
    if ($noOfArraysValue -ne $null) {
        $cellReference = "H43"  # Off-peak: no_of_arrays is at H43
        
        Write-Host "=== STEP 6.1: Processing no_of_arrays first ===" -ForegroundColor Magenta
        Write-Host "Setting no_of_arrays to $noOfArraysValue to unlock arrays 1 through $noOfArraysValue" -ForegroundColor Cyan
        
        try {
            $cell = $worksheet.Range($cellReference)
            
            # Always trigger VBA when setting no_of_arrays to unlock array cells
            # This is needed even if there's no array data yet - it unlocks the cells for future input
            Write-Host "Processing no_of_arrays dropdown with special VBA triggering..." -ForegroundColor Cyan
            
            # VBA triggering path - this unlocks the array cells
            try {
                    Write-Host "Triggering VBA to unlock array cells for no_of_arrays..." -ForegroundColor Cyan
                    
                    # Ensure worksheet is active
                    $worksheet.Activate()
                    Start-Sleep -Milliseconds 100
                    
                    # Get fresh cell reference
                    $cell = $worksheet.Range($cellReference)
                    if ($null -eq $cell) {
                        throw "Could not get cell reference $cellReference"
                    }
                    
                    Write-Host "Current cell value before change: $($cell.Value2)" -ForegroundColor Yellow
                    
                    # Enable events specifically for this dropdown to trigger VBA
                    $excel.EnableEvents = $true
                    Write-Host "Enabled Excel events for no_of_arrays VBA triggering" -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    
                    # Get the cell again after enabling events
                    $cell = $worksheet.Range($cellReference)
                    
                    # Clear the cell first (only if needed)
                    if ($null -ne $cell) {
                        try {
                            $cell.Value2 = ""
                            Start-Sleep -Milliseconds 100
                            Write-Host "Cleared cell $cellReference" -ForegroundColor Yellow
    } catch {
                            Write-Host "Warning: Could not clear cell (may already be empty): $_" -ForegroundColor Yellow
                        }
                    }
                    
                    # Set the value as string (dropdown values are strings)
                    $cell.Value2 = [string]$noOfArraysValue
                    Write-Host "Set no_of_arrays = '$noOfArraysValue' (as string) to cell $cellReference" -ForegroundColor Green
                    Start-Sleep -Milliseconds 100
                    
                    # Select the cell to simulate user interaction
                    $worksheet.Activate()
                    $cell.Select()
                    Start-Sleep -Milliseconds 200
                    
                    # Force calculations to trigger VBA
                    $excel.Calculate()
                    $excel.CalculateFullRebuild()
                    Start-Sleep -Milliseconds 500
                    
                    # Trigger worksheet change event by selecting and re-setting
                    $worksheet.Range($cellReference).Select()
                    $worksheet.Range($cellReference).Value2 = [string]$noOfArraysValue
                    Start-Sleep -Milliseconds 100
                    
                    # Wait for VBA to complete and unlock cells
                    Start-Sleep -Milliseconds 1500
                    
                    # Disable events again for other operations
                    $excel.EnableEvents = $false
                    Write-Host "Disabled Excel events after VBA processing" -ForegroundColor Cyan
                    
                    Write-Host "VBA triggering completed for no_of_arrays - array cells should now be unlocked" -ForegroundColor Green
                    $arraySavedCount++
                    
                } catch {
                    $errorMessage = $_.Exception.Message
                    Write-Host "Warning: Error during no_of_arrays VBA triggering : $errorMessage" -ForegroundColor Yellow
                    # Ensure events are disabled even on error
                    try { $excel.EnableEvents = $false } catch {}
                    # Fallback to simple assignment
                    $cell.Value2 = [string]$noOfArraysValue
                    Write-Host "Fallback: Saved no_of_arrays = $noOfArraysValue (as string) to cell $cellReference" -ForegroundColor Green
                    $arraySavedCount++
                }
            
        } catch {
            Write-Host "Error saving no_of_arrays to cell $cellReference : $_" -ForegroundColor Red
        }
    }
    
    # STEP 6.2: Process arrays one by one (array_1, then array_2, etc.)
    $maxArrays = 8
    for ($arrayNum = 1; $arrayNum -le $maxArrays; $arrayNum++) {
        $arrayFields = @()
        $hasArrayData = $false
        
        # Collect all fields for this array
        foreach ($inputKey in $dynamicInputs.Keys) {
            if ($inputKey -match "^array_$arrayNum" + "_") {
                $arrayFields += @{Key = $inputKey; Value = $dynamicInputs[$inputKey]}
                $hasArrayData = $true
            }
        }
        
        if ($hasArrayData) {
            Write-Host ""
            Write-Host ("=== STEP 6.2.$arrayNum : Processing Array $arrayNum ===") -ForegroundColor Magenta
            
            # Sort array fields in logical order: num_panels, orientation, pitch, shading
            $sortedFields = $arrayFields | Sort-Object {
                switch -Regex ($_.Key) {
                    "_num_panels$" { return 1 }
                    "_orientation_" { return 2 }
                    "_pitch_" { return 3 }
                    "_shading_" { return 4 }
                    default { return 5 }
                }
            }
            
            foreach ($fieldData in $sortedFields) {
                $inputKey = $fieldData.Key
                $inputValue = $fieldData.Value
                
                # Skip empty values
                if ([string]::IsNullOrWhiteSpace($inputValue)) {
                    continue
                }
                
                $cellReference = $arrayCellMappings[$inputKey]
                
                if ($cellReference) {
                    try {
                        $cell = $worksheet.Range($cellReference)
                        
                        # Check if cell is locked
                        if ($cell.Locked) {
                            Write-Host "Warning: Cell $cellReference is locked, skipping input $inputKey" -ForegroundColor Yellow
                            continue
                        }
                        
                        # All array fields are numeric
                        try {
                            $cleanValue = $inputValue.ToString().Trim()
                            
                            # Try different conversion methods
                            $convertedValue = 0
                            if ([double]::TryParse($cleanValue, [ref]$convertedValue)) {
                                $cell.Value2 = $convertedValue
                                Write-Host "Saved $inputKey = $convertedValue (as number) to cell $cellReference" -ForegroundColor Green
                            } else {
                                # Try parsing as integer
                                $intValue = 0
                                if ([int]::TryParse($cleanValue, [ref]$intValue)) {
                                    $cell.Value2 = $intValue
                                    Write-Host "Saved $inputKey = $intValue (as integer) to cell $cellReference" -ForegroundColor Green
                                } else {
                                    # Fallback to string if conversion fails
                                    $cell.Value2 = [string]$inputValue
                                    Write-Host "Warning: Could not convert '$inputValue' to number for $inputKey, saved as string" -ForegroundColor Yellow
                                }
                            }
                            # Clear any data validation to prevent Excel popups
                            try { $cell.Validation.Delete() } catch { }
                        } catch {
                            Write-Host "Warning: Could not convert '$inputValue' to number for $inputKey, setting as string. Error: $_" -ForegroundColor Yellow
                            $cell.Value2 = [string]$inputValue
                            # Clear any data validation to prevent Excel popups
                            try { $cell.Validation.Delete() } catch { }
                        }
                        $arraySavedCount++
                        
                        # Small delay between array field entries
                        Start-Sleep -Milliseconds 50
                        
                    } catch {
                        Write-Host "Error saving $inputKey to cell $cellReference : $_" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Warning: No cell mapping found for input: $inputKey" -ForegroundColor Yellow
                }
            }
        }
    }
    
    Write-Host "Successfully saved $arraySavedCount array field values" -ForegroundColor Green
    
    Write-Host "Step 6 completed: Solar array data added" -ForegroundColor Green
    
    # STEP 7: Payment/Pricing Data
    Write-Host ""
    Write-Host "Step 7: Inputting payment/pricing data..." -ForegroundColor Magenta
    Write-Host ""
    
    $paymentSavedCount = 0
    
    # Process payment method selection (ALWAYS from pricingData, overrides Step 3)
    $paymentMethod = $null
    if ($dynamicInputs.ContainsKey("payment_method")) {
        $paymentMethod = $dynamicInputs["payment_method"]
        Write-Host "Found payment_method in dynamicInputs: $paymentMethod" -ForegroundColor Cyan
    }
    
    if ($paymentMethod) {
        Write-Host "Setting payment method from pricingData: $paymentMethod" -ForegroundColor Cyan
        try {
            # Map payment method from pricingData to radio button shape name
            $shapeName = switch ($paymentMethod.ToLower()) {
                "hometree" { "Hometree" }
                "cash" { "Cash" }
                "finance" { "Finance" }
                "newfinance" { "NewFinance" }
                default { 
                    Write-Host "Unknown payment method '$paymentMethod', defaulting to Hometree" -ForegroundColor Yellow
                    "Hometree" 
                }
            }
            
            Write-Host "Looking for payment radio button shape: $shapeName" -ForegroundColor Yellow
            
            # Activate the worksheet and ensure Excel is ready
            $worksheet.Activate()
            Start-Sleep -Milliseconds 300
            
            # Enable events to allow macro to execute properly
            $excel.EnableEvents = $true
            Start-Sleep -Milliseconds 100
            
            # Try to find and select the radio button shape (like old routes do)
            $shapeFound = $false
            try {
                $targetShape = $worksheet.Shapes($shapeName)
                Write-Host "Found payment radio button shape: $shapeName" -ForegroundColor Green
                
                # Get the OnAction macro from the shape
                $onActionMacro = ""
                try {
                    $onActionMacro = $targetShape.OnAction
                    if ($onActionMacro) {
                        Write-Host "Shape has OnAction macro: $onActionMacro" -ForegroundColor Cyan
                    }
                } catch {
                    Write-Host "Shape does not have OnAction macro" -ForegroundColor Yellow
                }
                
                # Try to select the radio button using ControlFormat
                try {
                    if ($targetShape.ControlFormat) {
                        Write-Host "Selecting payment radio button using ControlFormat..." -ForegroundColor Yellow
                        $targetShape.ControlFormat.Value = 1
                        $shapeFound = $true
                        Write-Host "Successfully selected payment radio button: $shapeName" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "ControlFormat selection failed, trying shape selection..." -ForegroundColor Yellow
                }
                
                # If ControlFormat didn't work, try selecting the shape
                if (-not $shapeFound) {
                    try {
                        Write-Host "Selecting payment radio button shape..." -ForegroundColor Yellow
                        $targetShape.Select()
                        Start-Sleep -Milliseconds 200
                        $excel.SendKeys("{ENTER}")
                        Start-Sleep -Milliseconds 200
                        $shapeFound = $true
                        Write-Host "Successfully selected payment radio button shape: $shapeName" -ForegroundColor Green
                    } catch {
                        Write-Host "Shape selection failed: $_" -ForegroundColor Yellow
                    }
                }
                
                # Try to trigger the OnAction macro if it exists
                if ($onActionMacro -and $onActionMacro.Trim() -ne "") {
                    try {
                        Write-Host "Executing OnAction macro: $onActionMacro" -ForegroundColor Cyan
                        $excel.Run($onActionMacro)
                        Write-Host "Successfully executed OnAction macro: $onActionMacro" -ForegroundColor Green
                    } catch {
                        Write-Host "Failed to execute OnAction macro: $_" -ForegroundColor Yellow
                    }
                }
            } catch {
                Write-Host "Payment radio button shape '$shapeName' not found, trying macro approach..." -ForegroundColor Yellow
                
                # Fallback: Try to run the macro directly
                # Macro names: Hometree -> SetOptionHomeTree, Cash -> SetOptionCash, Finance/NewFinance -> SetOptionNewFinance
                $macroName = switch ($paymentMethod.ToLower()) {
                    "hometree" { "SetOptionHomeTree" }
                    "cash" { "SetOptionCash" }
                    "finance" { "SetOptionNewFinance" }
                    "newfinance" { "SetOptionNewFinance" }
                    default { 
                        Write-Host "Unknown payment method '$paymentMethod', defaulting to SetOptionHomeTree" -ForegroundColor Yellow
                        "SetOptionHomeTree" 
                    }
                }
                
                Write-Host "Executing payment method macro directly: $macroName" -ForegroundColor Yellow
                $excel.Run($macroName)
                Write-Host "Successfully executed payment method macro: $macroName" -ForegroundColor Green
            }
            
            # Wait for macro/shape selection to complete
            Start-Sleep -Milliseconds 1000
            
            # Force calculation after payment method change
            $excel.Calculate()
            $excel.CalculateFullRebuild()
            Start-Sleep -Milliseconds 500
            
            # Disable events again
            $excel.EnableEvents = $false
            
            Write-Host "Successfully set payment method: $paymentMethod (shape: $shapeName)" -ForegroundColor Green
            $paymentSavedCount++
        } catch {
            Write-Host "Error setting payment method: $_" -ForegroundColor Red
            Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        }
    } else {
        Write-Host "Warning: No payment_method found in dynamicInputs, skipping payment method selection" -ForegroundColor Yellow
    }
    
    # Payment data cell mappings for off-peak calculator
    # Based on createSaveDynamicInputsScript: Off Peak uses H80-H84
    $paymentMappings = @{
        "total_system_cost" = "H80"    # Off Peak: H80
        "deposit" = "H81"              # Off Peak: H81
        "interest_rate" = "H82"        # Off Peak: H82
        "interest_rate_type" = "H83"   # Off Peak: H83
        "payment_term" = "H84"          # Off Peak: H84
    }
    
    # Wait a bit for payment method selection to unlock payment fields
    Start-Sleep -Milliseconds 500
    
    # Process payment fields (including total_system_cost)
    foreach ($paymentKey in $paymentMappings.Keys) {
        if ($dynamicInputs.ContainsKey($paymentKey)) {
            $paymentValue = $dynamicInputs[$paymentKey]
            
            # Check if value is not empty
            if (-not [string]::IsNullOrWhiteSpace($paymentValue)) {
                $cellReference = $paymentMappings[$paymentKey]
                Write-Host "Processing payment field: $paymentKey = '$paymentValue' -> cell $cellReference" -ForegroundColor Cyan
                
                try {
                    $cell = $worksheet.Range($cellReference)
                    
                    # If cell is locked, try to unprotect worksheet first
                    if ($cell.Locked) {
                        Write-Host "Cell $cellReference is locked, attempting to unprotect worksheet..." -ForegroundColor Yellow
                        try {
                            if ($worksheet.ProtectContents) {
                                $worksheet.Unprotect($password)
                                Write-Host "Worksheet unprotected successfully" -ForegroundColor Green
                            }
                        } catch {
                            Write-Host "Could not unprotect worksheet: $_" -ForegroundColor Yellow
                        }
                        
                        # Check again if cell is still locked
                        if ($cell.Locked) {
                            Write-Host "Warning: Cell $cellReference is still locked after unprotecting, trying to unlock cell directly..." -ForegroundColor Yellow
                            try {
                                $cell.Locked = $false
                                Write-Host "Cell unlocked successfully" -ForegroundColor Green
                            } catch {
                                Write-Host "Could not unlock cell: $_" -ForegroundColor Yellow
                            }
                        }
                    }
                    
                    # Try to write to the cell (even if it's locked, the payment method might have unlocked it)
                    try {
                        # Try to parse as number
                        $numericValue = $null
                        if ([double]::TryParse($paymentValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numericValue)) {
                            $cell.Value2 = $numericValue
                            Write-Host "Saved $paymentKey = $numericValue (as number) to cell $cellReference" -ForegroundColor Green
                        } else {
                            $cell.Value2 = [string]$paymentValue
                            Write-Host "Saved $paymentKey = '$paymentValue' (as string) to cell $cellReference" -ForegroundColor Green
                        }
                        
                        # Clear data validation if present
                        try {
                            $cell.Validation.Delete()
                        } catch {
                            # Ignore - cell may not have validation
                        }
                        
                        $paymentSavedCount++
                    } catch {
                        Write-Host "Warning: Could not write to cell $cellReference (may be locked): $_" -ForegroundColor Yellow
                        Write-Host "Attempting to force write by unlocking cell..." -ForegroundColor Yellow
                        
                        # Last resort: try to unlock and write again
                        try {
                            $cell.Locked = $false
                            if ([double]::TryParse($paymentValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numericValue)) {
                                $cell.Value2 = $numericValue
                                Write-Host "Successfully saved $paymentKey = $numericValue (as number) to cell $cellReference after unlocking" -ForegroundColor Green
                            } else {
                                $cell.Value2 = [string]$paymentValue
                                Write-Host "Successfully saved $paymentKey = '$paymentValue' (as string) to cell $cellReference after unlocking" -ForegroundColor Green
                            }
                            $paymentSavedCount++
                        } catch {
                            Write-Host "Error: Could not save $paymentKey to $cellReference even after unlocking: $_" -ForegroundColor Red
                        }
                    }
                } catch {
                    Write-Host "Error accessing cell $cellReference for $paymentKey : $_" -ForegroundColor Red
                }
            } else {
                Write-Host "Skipping payment field $paymentKey (value is empty or whitespace)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Payment field $paymentKey not found in dynamicInputs" -ForegroundColor Yellow
        }
    }
    
    Write-Host "Successfully saved $paymentSavedCount payment field values" -ForegroundColor Green
    Write-Host "Step 7 completed: Payment/pricing data added" -ForegroundColor Green
    
    # Save and close workbook (Step 7 is the final step)
    Write-Host ""
    Write-Host "Saving and closing workbook..." -ForegroundColor Green
    try {
        # Calculate formulas before saving
        Write-Host "Calculating formulas..." -ForegroundColor Yellow
        $excel.Calculate()
        $excel.CalculateFullRebuild()
        Start-Sleep -Milliseconds 500
        
        # Save the workbook
        Write-Host "Saving workbook with all changes..." -ForegroundColor Green
        $workbook.Save()
        Write-Host "Workbook saved successfully" -ForegroundColor Green
        
        # Close workbook
        Write-Host "Closing workbook..." -ForegroundColor Green
        $workbook.Close($true)
        Write-Host "Workbook closed successfully" -ForegroundColor Green
        
        # Quit Excel
        Write-Host "Quitting Excel application..." -ForegroundColor Green
        $excel.Quit()
        
        # Cleanup COM objects
        Write-Host "Releasing COM objects..." -ForegroundColor Yellow
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "COM objects released" -ForegroundColor Green
        
        Write-Host "All steps completed successfully: Workbook saved and Excel closed." -ForegroundColor Green
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error during save/close : $errorMessage" -ForegroundColor Red
        throw "Could not save or close workbook"
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    
    # Cleanup on error
    if ($worksheet) {
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null } catch {}
    }
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    
    exit 1
}
`;
  }

  private createCheckEnabledCellsScript(excelFilePath: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
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
        @{ id = "single_day_rate"; label = "Single / Day Rate (pence per kWh)"; type = "number"; cellReference = "H19" },
        @{ id = "night_rate"; label = "Night Rate (pence per kWh)"; type = "number"; cellReference = "H20" },
        @{ id = "off_peak_hours"; label = "No. of Off-Peak Hours"; type = "number"; cellReference = "H21" },
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

  private parseEnabledCellsFromPowerShellOutput(output: string): any[] {
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

  private async getAllDropdownOptions(excelFilePath: string): Promise<Record<string, string[]>> {
    try {
      this.logger.log('🚀 STARTING: Getting ALL dropdown options from Excel using direct range reading...');
      this.logger.log(`🚀 DEBUG: Excel file path: ${excelFilePath}`);
      this.logger.log(`🚀 DEBUG: File exists: ${fs.existsSync(excelFilePath)}`);
      
      // Read the Excel file
      const workbook = XLSX.readFile(excelFilePath, { 
        password: this.PASSWORD,
        cellStyles: true,
        cellDates: true,
        cellFormula: true
      });
      
      this.logger.log(`🚀 DEBUG: Excel file read successfully. Available sheets: ${Object.keys(workbook.Sheets).join(', ')}`);

      const dropdownData: Record<string, string[]> = {};

      // Define the dropdown field mappings to their Excel ranges where the options are stored
      const dropdownRanges = {
        panel_manufacturer: { sheet: 'Panels', range: 'P4:AP4' }, // Panel manufacturers in row 4, columns P to AP
        panel_model: { sheet: 'Panels', range: 'Q5:AP50' }, // Panel models in columns Q to AP, rows 5-50
        battery_manufacturer: { sheet: 'Batteries', range: 'P4:AP4' }, // Battery manufacturers in row 4, columns P to AP
        battery_model: { sheet: 'Batteries', range: 'Q5:AP50' }, // Battery models in columns Q to AP, rows 5-50
        solar_inverter_manufacturer: { sheet: 'Inverters', range: 'L4:AA4' }, // Solar inverter manufacturers in row 4, columns L to AA
        solar_inverter_model: { sheet: 'Inverters', range: 'L5:AA50' }, // Solar inverter models in columns L to AA, rows 5-50
        battery_inverter_manufacturer: { sheet: 'Inverters', range: 'L65:N65' }, // Battery inverter manufacturers in row 65, columns L to N
        battery_inverter_model: { sheet: 'Inverters', range: 'L66:N100' }, // Battery inverter models in columns L to N, rows 66-100
        interest_rate_type: { sheet: 'Inputs', range: 'H84:H86' }, // Interest rate type options in Inputs sheet
        no_of_arrays: { static: ['1', '2', '3', '4', '5', '6', '7', '8'] } // Static dropdown options for number of arrays
      };

      // Read dropdown options for each field from their respective ranges
      for (const [fieldId, rangeInfo] of Object.entries(dropdownRanges)) {
        try {
          // Handle static dropdown options
          if ('static' in rangeInfo) {
            this.logger.log(`Using static dropdown options for ${fieldId}`);
            dropdownData[fieldId] = rangeInfo.static;
            this.logger.log(`Found ${rangeInfo.static.length} static options for ${fieldId}: ${rangeInfo.static.join(', ')}`);
            continue;
          }
          
          this.logger.log(`Reading dropdown options for ${fieldId} from ${rangeInfo.sheet}:${rangeInfo.range}`);
          
          const worksheet = workbook.Sheets[rangeInfo.sheet];
          if (!worksheet) {
            this.logger.warn(`Sheet ${rangeInfo.sheet} not found for ${fieldId}`);
            dropdownData[fieldId] = [];
            continue;
          }
          
          const options = this.readOptionsFromRange(workbook, `${rangeInfo.sheet}!${rangeInfo.range}`);
          
          if (options && options.length > 0) {
            dropdownData[fieldId] = options;
            this.logger.log(`Found ${options.length} options for ${fieldId}: ${options.slice(0, 3).join(', ')}${options.length > 3 ? '...' : ''}`);
          } else {
            this.logger.warn(`No dropdown options found for ${fieldId} in range ${rangeInfo.sheet}:${rangeInfo.range}`);
            dropdownData[fieldId] = [];
          }
        } catch (error) {
          this.logger.error(`Error reading dropdown options for ${fieldId}:`, error);
          dropdownData[fieldId] = [];
        }
      }

      // Add fallback for interest rate type if not found in Excel or if it contains invalid options
      if (!dropdownData.interest_rate_type || dropdownData.interest_rate_type.length === 0 || 
          !dropdownData.interest_rate_type.includes('Fixed')) {
        this.logger.log('Adding fallback interest rate type options (Excel data was invalid or missing)');
        this.logger.log(`Excel returned: ${JSON.stringify(dropdownData.interest_rate_type)}`);
        dropdownData.interest_rate_type = ['Fixed', 'APR', 'Variable'];
      }

      this.logger.log('✅ SUCCESS: Successfully read all dropdown options');
      this.logger.log(`✅ DEBUG: Final dropdown data: ${JSON.stringify(dropdownData, null, 2)}`);
      return dropdownData;
      
    } catch (error) {
      this.logger.error('Error in getAllDropdownOptions:', error);
      return {};
    }
  }

  private async getDropdownOptionsForField(fieldId: string, cellReference: string, excelFilePath: string): Promise<string[]> {
    try {
      this.logger.log(`🔍 DEBUG: Reading dropdown options for ${fieldId} from cell ${cellReference}`);
      this.logger.log(`🔍 DEBUG: Reading from file: ${excelFilePath}`);
      this.logger.log(`🔍 DEBUG: File exists: ${fs.existsSync(excelFilePath)}`);
      
      this.logger.log(`🔍 DEBUG: About to read Excel file with XLSX...`);
      const workbook = XLSX.readFile(excelFilePath, { 
        password: this.PASSWORD,
        cellStyles: true,
        cellDates: true,
        cellFormula: true
      });
      this.logger.log(`🔍 DEBUG: Excel file read successfully!`);

      this.logger.log(`🔍 DEBUG: Excel file read successfully. Available sheets: ${Object.keys(workbook.Sheets).join(', ')}`);

      const worksheet = workbook.Sheets['Inputs'];
      if (!worksheet) {
        this.logger.warn(`🔍 DEBUG: Inputs sheet not found. Available sheets: ${Object.keys(workbook.Sheets).join(', ')}`);
        return [];
      }

      this.logger.log(`🔍 DEBUG: Inputs sheet found. Checking for data validation...`);

      // Check if the cell has data validation
      if (worksheet['!dataValidation']) {
        this.logger.log(`🔍 DEBUG: Data validation found in worksheet. Keys: ${Object.keys(worksheet['!dataValidation']).join(', ')}`);
        
        const validation = worksheet['!dataValidation'][cellReference];
        if (validation) {
          this.logger.log(`🔍 DEBUG: Found validation for ${cellReference}: ${JSON.stringify(validation)}`);
          
          if (validation.type === 'list') {
            this.logger.log(`🔍 DEBUG: Found list validation for ${cellReference}: ${validation.formula1}`);
            
            // Parse the validation formula to get options
            const options = this.parseValidationFormula(validation.formula1, workbook);
            this.logger.log(`🔍 DEBUG: Parsed ${options.length} options from validation formula`);
            return options;
          } else {
            this.logger.log(`🔍 DEBUG: Validation type is not 'list': ${validation.type}`);
          }
        } else {
          this.logger.log(`🔍 DEBUG: No validation found for cell ${cellReference}`);
        }
      } else {
        this.logger.log(`🔍 DEBUG: No data validation found in worksheet`);
      }

      // If no data validation found, try to read from a lookup table
      this.logger.log(`🔍 DEBUG: Trying to read from lookup table for ${fieldId}`);
      const lookupOptions = this.readOptionsFromLookupTable(workbook, fieldId);
      this.logger.log(`🔍 DEBUG: Found ${lookupOptions.length} options from lookup table`);
      this.logger.log(`🔍 DEBUG: Returning ${lookupOptions.length} options for ${fieldId}: ${lookupOptions.slice(0, 3).join(', ')}${lookupOptions.length > 3 ? '...' : ''}`);
      return lookupOptions;
      
    } catch (error) {
      this.logger.error(`🔍 DEBUG: Error reading dropdown options for ${fieldId}:`, error);
      return [];
    }
  }

  private parseValidationFormula(formula: string, workbook: any): string[] {
    try {
      if (!formula) return [];

      // Remove the '=' if present
      let cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;

      // Handle INDIRECT formulas
      if (cleanFormula.toUpperCase().includes('INDIRECT')) {
        // Extract the cell reference from INDIRECT
        const match = cleanFormula.match(/INDIRECT\(([^)]+)\)/i);
        if (match) {
          const cellRef = match[1].replace(/['"]/g, '').replace(/\$/g, '');
          return this.readOptionsFromCellReference(workbook, cellRef);
        }
      }

      // Handle direct range references
      if (cleanFormula.includes(':')) {
        return this.readOptionsFromRange(workbook, cleanFormula);
      }

      // Handle comma-separated lists
      if (cleanFormula.includes(',')) {
        return cleanFormula.split(',').map(opt => opt.trim().replace(/['"]/g, ''));
      }

      // Handle single cell references
      return this.readOptionsFromCellReference(workbook, cleanFormula);

    } catch (error) {
      this.logger.error('Error parsing validation formula:', error);
      return [];
    }
  }

  private readOptionsFromRange(workbook: any, range: string): string[] {
    try {
      const [sheetName, cellRange] = range.includes('!') ? range.split('!') : ['Inputs', range];
      const worksheet = workbook.Sheets[sheetName];
      
      if (!worksheet) return [];

      const rangeObj = XLSX.utils.decode_range(cellRange);
      const options: string[] = [];

      for (let row = rangeObj.s.r; row <= rangeObj.e.r; row++) {
        for (let col = rangeObj.s.c; col <= rangeObj.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
            options.push(String(cell.v).trim());
          }
        }
      }

      // Remove duplicates while preserving order
      const uniqueOptions = [...new Set(options)];
      
      if (uniqueOptions.length !== options.length) {
        this.logger.log(`Removed ${options.length - uniqueOptions.length} duplicate options from ${sheetName}!${cellRange}`);
      }

      return uniqueOptions;
    } catch (error) {
      this.logger.error('Error reading options from range:', error);
      return [];
    }
  }

  private readOptionsFromCellReference(workbook: any, cellRef: string): string[] {
    try {
      // This would need to be implemented based on how the cell reference is used
      // For now, return empty array
      return [];
    } catch (error) {
      this.logger.error('Error reading options from cell reference:', error);
      return [];
    }
  }

  private async getManufacturerSpecificModels(fieldId: string, manufacturer: string, opportunityId: string | undefined): Promise<string[]> {
    try {
      this.logger.log(`Getting manufacturer-specific models for ${fieldId} with manufacturer: ${manufacturer}`);
      
      // Determine which Excel file to use
      let excelFilePath: string;
      if (opportunityId) {
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId, 'off-peak');
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
        } else {
          this.logger.warn(`Opportunity file not found, using template: ${opportunityFilePath}`);
          excelFilePath = this.getTemplateFilePath();
        }
      } else {
        excelFilePath = this.getTemplateFilePath();
      }

      if (!fs.existsSync(excelFilePath)) {
        this.logger.error(`Excel file not found at: ${excelFilePath}`);
        return [];
      }

      // Read the Excel file
      const workbook = XLSX.readFile(excelFilePath, { 
        password: this.PASSWORD,
        cellStyles: true,
        cellDates: true,
        cellFormula: true
      });

      // Define manufacturer to column mapping based on the Excel structure
      const manufacturerMappings = {
        panel_model: {
          sheet: 'Panels',
          manufacturers: {
            'Tier 1': 'P5:P50',
            'Aiko': 'Q5:Q8',
            'AmeriSolar': 'R5:R8',
            'Canadian Solar': 'S5:S46',
            'DAS Solar': 'T5:T9',
            'DMEGC Solar': 'U5:U20',
            'Energizer': 'V5:V7',
            'Eurener': 'W5:W34',
            'Evolution': 'X5:X6',
            'Exiom Solutions': 'Y5:Y9',
            'Hanwha Q Cells': 'Z5:Z50',
            'Hyundai': 'AA5:AA8',
            'JA Solar': 'AB5:AB33',
            'Jinko Solar': 'AC5:AC14',
            'Longi': 'AD5:AD28',
            'Meyer Burger': 'AE5:AE19',
            'Perlight Solar': 'AF5:AF16',
            'Sharp': 'AG5:AG6',
            'Solarwatt': 'AH5:AH9',
            'Sunket': 'AI5:AI10',
            'Sunrise Energy': 'AJ5:AJ13',
            'Suntech': 'AK5:AK14',
            'Tenka Solar': 'AL5:AL9',
            'Tongwei': 'AM5:AM9',
            'Trina Solar': 'AN5:AN13',
            'UKSOL': 'AO5:AO9',
            'Ulica': 'AP5:AP10'
          }
        },
        battery_model: {
          sheet: 'Batteries',
          manufacturers: {
            'Generic Batteries': 'P5:P9',
            'Aoboet': 'Q5:Q20',
            'Alpha ESS': 'R5:R17',
            'Duracell': 'S5:S5',
            'Dyness': 'T5:T16',
            'Enphase Energy': 'U5:U7',
            'FoxESS': 'V5:V42',
            'GivEnergy': 'W5:W10',
            'GoodWe': 'X5:X10',
            'Greenlinx': 'Y5:Y12',
            'Growatt': 'Z5:Z31',
            'Hanchu ESS': 'AA5:AA6',
            'Huawei': 'AB5:AB10',
            'Lux Power': 'AC5:AC6',
            'myenergi': 'AD5:AD8',
            'Puredrive': 'AE5:AE9',
            'Pylontech': 'AF5:AF52',
            'Pytes ESS': 'AG5:AG11',
            'SAJ': 'AH5:AH12',
            'Sigenergy': 'AI5:AI10',
            'Sofar Solar': 'AJ5:AJ8',
            'SolarEdge': 'AK5:AK5',
            'Solax': 'AL5:AL25',
            'Sunsynk': 'AM5:AM27',
            'Soluna': 'AN5:AN10',
            'Wonderlux': 'AP5:AP6',
            'EcoFlow': 'AP5:AP7'
          }
        },
        solar_inverter_model: {
          sheet: 'Inverters',
          manufacturers: {
            'S Enphase Energy': 'L5:L10',
            'S Duracell': 'M5:M15',
            'S Fox ESS': 'N5:N20',
            'S Fronius': 'O5:O25',
            'S GivEnergy': 'P5:P15',
            'S GoodWe': 'Q5:Q20',
            'S Growatt': 'R5:R30',
            'S Huawei': 'S5:S25',
            'S Hypontec': 'T5:T15',
            'S Lux Power': 'U5:U20',
            'S Sigenergy': 'V5:V25',
            'S SolarEdge': 'W5:W15',
            'S SolaX': 'X5:X20',
            'S Solis': 'Y5:Y10',
            'S Sunsynk': 'Z5:Z25',
            'S EcoFlow': 'AA5:AA15'
          }
        },
        battery_inverter_model: {
          sheet: 'Inverters',
          manufacturers: {
            'B Growatt': 'L66:L100',
            'B Lux Power': 'M66:M100',
            'B Sunsynk': 'N66:N100'
          }
        }
      };

      const mapping = manufacturerMappings[fieldId];
      if (!mapping) {
        this.logger.warn(`No mapping found for field: ${fieldId}`);
        return [];
      }

      const worksheet = workbook.Sheets[mapping.sheet];
      if (!worksheet) {
        this.logger.warn(`Sheet ${mapping.sheet} not found`);
        return [];
      }

      const rangeAddress = mapping.manufacturers[manufacturer];
      if (!rangeAddress) {
        this.logger.warn(`No range found for manufacturer: ${manufacturer} in field: ${fieldId}`);
        return [];
      }

      this.logger.log(`Reading models for ${manufacturer} from ${mapping.sheet}:${rangeAddress}`);
      
      const options = this.readOptionsFromRange(workbook, `${mapping.sheet}!${rangeAddress}`);
      
      this.logger.log(`Found ${options.length} models for ${manufacturer}: ${options.slice(0, 3).join(', ')}${options.length > 3 ? '...' : ''}`);
      
      return options;
      
    } catch (error) {
      this.logger.error(`Error getting manufacturer-specific models for ${fieldId}:`, error);
      return [];
    }
  }

  private readOptionsFromLookupTable(workbook: any, fieldId: string): string[] {
    try {
      // Try to find a lookup table in common sheets
      const lookupSheets = ['Lookups', 'Data', 'PanelData', 'BatteryData', 'InverterData'];
      
      for (const lookupSheet of lookupSheets) {
        const worksheet = workbook.Sheets[lookupSheet];
        if (worksheet) {
          // Read all data from the lookup sheet
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
          const options: string[] = [];

          for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 0 }); // Column A
            const cell = worksheet[cellAddress];
            
            if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
              options.push(String(cell.v).trim());
            }
          }

          if (options.length > 0) {
            this.logger.log(`Found ${options.length} options in lookup sheet ${lookupSheet}`);
            return options;
          }
        }
      }

      return [];
    } catch (error) {
      this.logger.error('Error reading options from lookup table:', error);
      return [];
    }
  }

  private async runNodeJsExcelAnalysis(excelFilePath: string): Promise<{ success: boolean; message: string; error?: string; inputFields?: any[] }> {
    try {
      this.logger.log('Running integrated Excel analysis...');
      
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
        { id: 'battery_extended_warranty_period', cell: 'H49', label: 'Battery Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
        { id: 'battery_replacement_cost', cell: 'H50', label: 'Battery Replacement Cost (£)', type: 'number', required: false, section: 'New System' },
        
        // NEW SYSTEM - SOLAR/HYBRID INVERTER
        { id: 'solar_inverter_manufacturer', cell: 'H52', label: 'Solar/Hybrid Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
        { id: 'solar_inverter_model', cell: 'H53', label: 'Solar/Hybrid Inverter Model', type: 'text', required: false, section: 'New System' },
        { id: 'solar_inverter_extended_warranty_period', cell: 'H56', label: 'Solar Inverter Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
        { id: 'solar_inverter_replacement_cost', cell: 'H57', label: 'Solar Inverter Replacement Cost (£)', type: 'number', required: false, section: 'New System' },
        
        // NEW SYSTEM - BATTERY INVERTER
        { id: 'battery_inverter_manufacturer', cell: 'H59', label: 'Battery Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
        { id: 'battery_inverter_model', cell: 'H60', label: 'Battery Inverter Model', type: 'text', required: false, section: 'New System' },
        { id: 'battery_inverter_extended_warranty_period', cell: 'H63', label: 'Battery Inverter Extended Warranty Period (years)', type: 'number', required: false, section: 'New System' },
        { id: 'battery_inverter_replacement_cost', cell: 'H64', label: 'Battery Inverter Replacement Cost (£)', type: 'number', required: false, section: 'New System' }
      ];

      // Check if cell is correctly enabled based on background color
      const isCellCorrectlyEnabled = (cell: any) => {
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

        // Check if cell is locked (only matters if worksheet is protected)
        if (cell.l && cell.l.locked) {
          return { enabled: false, reason: 'Cell is locked' };
        }

        // Check if cell is hidden
        if (cell.l && cell.l.hidden) {
          return { enabled: false, reason: 'Cell is hidden' };
        }

        // If we get here, the cell appears to be enabled
        return { enabled: true, reason: 'Cell appears enabled' };
      };

      // Read Excel file and analyze
      const workbook = XLSX.readFile(excelFilePath, { 
        password: this.PASSWORD,
        cellStyles: true,
        cellDates: true,
        cellFormula: true
      });
      
      // Get the Inputs worksheet
      const inputsSheet = workbook.Sheets['Inputs'];
      if (!inputsSheet) {
        throw new Error('Inputs sheet not found');
      }
      
      // Analyze all input fields based on correct background color reading
      const results: any[] = [];
      let enabledCount = 0;
      let disabledCount = 0;
      
      for (const field of ALL_INPUT_FIELDS) {
        const cell = inputsSheet[field.cell];
        const enabledStatus = isCellCorrectlyEnabled(cell);
        
        this.logger.log(`Analyzing field ${field.id} (${field.cell}): enabled=${enabledStatus.enabled}, reason=${enabledStatus.reason}`);
        
        if (enabledStatus.enabled) {
          enabledCount++;
          results.push({
            id: field.id,
            label: field.label,
            type: field.type,
            required: field.required,
            value: cell ? (cell.v === 'undefined' ? '' : cell.v) : '',
            cellReference: field.cell,
            section: field.section,
            enabled: true
          });
        } else {
          disabledCount++;
          results.push({
            id: field.id,
            label: field.label,
            type: field.type,
            required: field.required,
            value: cell ? (cell.v === 'undefined' ? '' : cell.v) : '',
            cellReference: field.cell,
            section: field.section,
            enabled: false,
            reason: enabledStatus.reason
          });
        }
      }
      
      this.logger.log(`Excel analysis successful: Found ${enabledCount} enabled and ${disabledCount} disabled input fields out of ${ALL_INPUT_FIELDS.length} total`);
      
      return {
        success: true,
        message: `Found ${enabledCount} enabled and ${disabledCount} disabled input fields`,
        inputFields: results
      };
      
    } catch (error) {
      this.logger.error(`Integrated Excel analysis error: ${error.message}`);
      return {
        success: false,
        message: 'Integrated Excel analysis error',
        error: error.message
      };
    }
  }

  async getCascadingDropdownOptions(opportunityId: string | undefined, fieldId: string, dependsOnValue: string | undefined): Promise<{ success: boolean; message: string; error?: string; options?: string[] }> {
    try {
      this.logger.log(`Getting cascading dropdown options for ${fieldId} with dependency value: ${dependsOnValue}`);
      
      // Check if dependsOnValue is provided
      if (!dependsOnValue) {
        this.logger.warn(`No dependency value provided for field: ${fieldId}`);
        return { success: false, message: `No dependency value provided for field: ${fieldId}` };
      }
      
      // Define cascading dropdown field relationships
      const cascadingRelationships = {
        panel_model: { dependsOn: 'panel_manufacturer' },
        battery_model: { dependsOn: 'battery_manufacturer' },
        solar_inverter_model: { dependsOn: 'solar_inverter_manufacturer' },
        battery_inverter_model: { dependsOn: 'battery_inverter_manufacturer' }
      };

      // Get the cascading relationship for this field
      const relationship = cascadingRelationships[fieldId];
      if (!relationship) {
        this.logger.warn(`No cascading relationship found for field: ${fieldId}`);
        return { success: false, message: `No cascading relationship found for field: ${fieldId}` };
      }

      // Get manufacturer-specific models from Excel based on the selected manufacturer
      const manufacturerModels = await this.getManufacturerSpecificModels(fieldId, dependsOnValue, opportunityId);
      
      if (manufacturerModels.length === 0) {
        this.logger.warn(`No models found for ${fieldId} with manufacturer: ${dependsOnValue}`);
        return { success: false, message: `No models found for ${fieldId} with manufacturer: ${dependsOnValue}` };
      }

      const filteredOptions = manufacturerModels;

      this.logger.log(`Found ${filteredOptions.length} cascading options for ${fieldId}: ${filteredOptions.join(', ')}`);
      
      return {
        success: true,
        message: `Found ${filteredOptions.length} cascading options for ${fieldId}`,
        options: filteredOptions
      };

    } catch (error) {
      this.logger.error(`Error in getCascadingDropdownOptions: ${error.message}`);
      return {
        success: false,
        message: `Error getting cascading dropdown options`,
        error: error.message
      };
    }
  }

  async createOpportunityWithTemplate(opportunityId: string, customerDetails: { customerName: string; address: string; postcode: string }, templateFileName: string): Promise<{ success: boolean; message: string; error?: string; filePath?: string }> {
    this.logger.log(`Creating opportunity file for: ${opportunityId} with template: ${templateFileName}`);

    try {
      // Check if template file exists
      const templateFilePath = this.getTemplateFilePath(templateFileName);
      if (!fs.existsSync(templateFilePath)) {
        const error = `Template file not found at: ${templateFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Check if we're on Windows (required for Excel COM automation)
      if (process.platform !== 'win32') {
        const error = 'Excel automation requires Windows platform';
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Create opportunities folder if it doesn't exist
      if (!fs.existsSync(this.OPPORTUNITIES_FOLDER)) {
        fs.mkdirSync(this.OPPORTUNITIES_FOLDER, { recursive: true });
        this.logger.log(`Created opportunities folder: ${this.OPPORTUNITIES_FOLDER}`);
      }

      // Create PowerShell script
      const psScript = this.createOpportunityFileScript(opportunityId, customerDetails, templateFilePath, 'off-peak', true);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-create-opportunity-${Date.now()}.ps1`);
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
        // For template selection, we need to get the latest versioned file path
        const filePath = this.getNewOpportunityFilePath(opportunityId, 'off-peak');
        this.logger.log(`Successfully created opportunity file: ${filePath}`);
        return {
          success: true,
          message: `Successfully created opportunity file for: ${opportunityId}`,
          filePath: filePath
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to create opportunity file for: ${opportunityId}`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in createOpportunityWithTemplate: ${error.message}`);
      return {
        success: false,
        message: `Error creating opportunity file for: ${opportunityId}`,
        error: error.message
      };
    }
  }

  async getAllDropdownOptionsForFrontend(opportunityId: string | undefined, templateFileName?: string): Promise<{ success: boolean; message: string; error?: string; dropdownOptions?: Record<string, string[]> }> {
    try {
      this.logger.log(`Getting all dropdown options for frontend, opportunityId: ${opportunityId}, templateFileName: ${templateFileName}`);
      
      // Determine which Excel file to use - prioritize opportunity file if it exists
      let excelFilePath: string;
      if (opportunityId) {
        // First, try to use the opportunity file (created from template)
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId, 'off-peak');
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using opportunity file: ${opportunityFilePath}`);
        } else {
          // If opportunity file doesn't exist, fall back to template file
          if (templateFileName) {
            excelFilePath = this.getTemplateFilePath(templateFileName);
            this.logger.log(`Opportunity file not found, using template file: ${excelFilePath}`);
          } else {
            this.logger.warn(`Opportunity file not found, using default template: ${opportunityFilePath}`);
            excelFilePath = this.getTemplateFilePath();
          }
        }
      } else if (templateFileName) {
        // No opportunityId, use template file directly
        excelFilePath = this.getTemplateFilePath(templateFileName);
        this.logger.log(`Using template file: ${excelFilePath}`);
      } else {
        // Fallback to default template
        excelFilePath = this.getTemplateFilePath();
        this.logger.log(`Using default template file: ${excelFilePath}`);
      }

      if (!fs.existsSync(excelFilePath)) {
        const error = `Excel file not found at: ${excelFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Get all dropdown options from Excel dynamically
      const allDropdownOptions = await this.getAllDropdownOptions(excelFilePath);
      
      this.logger.log(`Successfully retrieved dropdown options for ${Object.keys(allDropdownOptions).length} fields`);
      this.logger.log(`✅ DEBUG: Final dropdown data:`, allDropdownOptions);
      
      return {
        success: true,
        message: `Successfully retrieved dropdown options for ${Object.keys(allDropdownOptions).length} fields`,
        dropdownOptions: allDropdownOptions
      };

    } catch (error) {
      this.logger.error(`Error in getAllDropdownOptionsForFrontend: ${error.message}`);
      return {
        success: false,
        message: `Error getting all dropdown options`,
        error: error.message
      };
    }
  }

  /**
   * Retrieve saved pricing data for an opportunity
   */
  async getSavedPricingData(opportunityId: string): Promise<{ success: boolean; data?: Record<string, string>; error?: string }> {
    try {
      this.logger.log(`🔍 Retrieving saved pricing data for opportunity: ${opportunityId}`);
      
      // Use the same file finding logic as generatePDF method
      this.logger.log(`🔍 Searching for files containing opportunity ID: ${opportunityId}`);
      const matchingFile = this.findLatestOpportunityFile(opportunityId);
      
      if (!matchingFile) {
        this.logger.error(`❌ No matching file found for opportunity: ${opportunityId}`);
        return {
          success: false,
          error: 'Excel file not found for opportunity'
        };
      }
      
      const excelFilePath = matchingFile;
      this.logger.log(`✅ Found matching file: ${excelFilePath}`);
      
      if (!fs.existsSync(excelFilePath)) {
        this.logger.error(`❌ Excel file not found: ${excelFilePath}`);
        return {
          success: false,
          error: 'Excel file not found for opportunity'
        };
      }
      
      this.logger.log(`✅ Excel file exists: ${excelFilePath}`);
      
      // Create PowerShell script to retrieve pricing data
      const psScript = this.createGetPricingDataScript(excelFilePath);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-get-pricing-data-${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created temporary pricing data retrieval script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Failed to cleanup temporary script: ${cleanupError.message}`);
      }

      if (result.success && result.output) {
        try {
          // Parse the JSON output from PowerShell
          const pricingData = JSON.parse(result.output);
          this.logger.log(`✅ Successfully retrieved pricing data:`, pricingData);
          
          return {
            success: true,
            data: pricingData
          };
        } catch (parseError) {
          this.logger.error(`Failed to parse pricing data JSON: ${parseError.message}`);
          return {
            success: false,
            error: 'Failed to parse pricing data'
          };
        }
      } else {
        this.logger.error(`Failed to retrieve pricing data: ${result.error}`);
        return {
          success: false,
          error: result.error || 'Failed to retrieve pricing data'
        };
      }

    } catch (error) {
      this.logger.error(`Error retrieving pricing data: ${error.message}`);
      return {
        success: false,
        error: error.message
      };
    }
  }

  async generatePDF(opportunityId: string, excelFilePath?: string, signatureData?: string, fileName?: string): Promise<{ success: boolean; message: string; error?: string; pdfPath?: string }> {
    try {
      this.logger.log(`🔍 Starting PDF generation for opportunity: ${opportunityId}`);
      
      // If no Excel file path provided, try to find the file
      if (!excelFilePath) {
        // Try different possible file patterns
        const possiblePaths = [
          // Regular calculator pattern
          this.findLatestOpportunityFile(opportunityId, 'off-peak', fileName),
          // EPVS pattern
          this.findLatestOpportunityFile(opportunityId, 'flux', fileName),
        ].filter(Boolean) as string[];
        
        // Find the first existing file
        for (const possiblePath of possiblePaths) {
          if (fs.existsSync(possiblePath)) {
            excelFilePath = possiblePath;
            this.logger.log(`✅ Found Excel file: ${excelFilePath}`);
            break;
          }
        }
        
        // If still no file found, try to find any file with the opportunity ID
        if (!excelFilePath) {
          this.logger.log(`🔍 Searching for files containing opportunity ID: ${opportunityId}`);
          
          // Search in both folders
          const searchFolders = [this.OPPORTUNITIES_FOLDER, this.EPVS_OPPORTUNITIES_FOLDER];
          
          for (const folder of searchFolders) {
            if (fs.existsSync(folder)) {
              const files = fs.readdirSync(folder);
              const matchingFile = files.find(file => file.includes(opportunityId) && file.endsWith('.xlsm'));
              
              if (matchingFile) {
                excelFilePath = path.join(folder, matchingFile);
                this.logger.log(`✅ Found matching file: ${excelFilePath}`);
                break;
              }
            }
          }
        }
      }
      
      this.logger.log(`🔍 Final Excel file path: ${excelFilePath}`);
      
      // Check if Excel file exists
      if (!excelFilePath || !fs.existsSync(excelFilePath)) {
        this.logger.error(`❌ Excel file not found for opportunity: ${opportunityId}`);
        this.logger.error(`❌ Searched paths: ${JSON.stringify([
          this.findLatestOpportunityFile(opportunityId, 'off-peak'),
          this.findLatestOpportunityFile(opportunityId, 'flux'),
        ].filter(Boolean))}`);
        return { success: false, message: 'Excel file not found', error: 'File does not exist' };
      }
      
      this.logger.log(`✅ Excel file exists: ${excelFilePath}`);
      
      // Determine the correct PDF folder based on the Excel file location
      let pdfFolder: string;
      let pdfFileName: string;
      
      if (excelFilePath.includes('epvs-opportunities')) {
        pdfFolder = path.join(this.EPVS_OPPORTUNITIES_FOLDER, 'pdfs');
        pdfFileName = `EPVS Calculator - ${opportunityId}.pdf`;
      } else {
        pdfFolder = path.join(this.OPPORTUNITIES_FOLDER, 'pdfs');
        pdfFileName = `Off Peak Calculator - ${opportunityId}.pdf`;
      }
      
      // Create PDF folder if it doesn't exist
      if (!fs.existsSync(pdfFolder)) {
        fs.mkdirSync(pdfFolder, { recursive: true });
      }
      
      const pdfPath = path.join(pdfFolder, pdfFileName);
      
      // Determine calculator type based on file path
      const isEPVS = excelFilePath.includes('epvs-opportunities');
      
      // Create PowerShell script for PDF generation (no pricing data re-population needed)
      const psScript = this.createPDFGenerationScript(excelFilePath, pdfPath, signatureData, isEPVS ? 'epvs' : 'regular');
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-pdf-generation-${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created temporary PDF generation script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Failed to cleanup temporary script: ${cleanupError.message}`);
      }

      if (result.success) {
        this.logger.log(`Successfully generated PDF: ${pdfPath}`);
        
        // Add signature to PDF if provided
        if (signatureData) {
          this.logger.log(`Adding signature to PDF: ${pdfPath}`);
          const signatureResult = await this.pdfSignatureService.addSignatureToPDF(pdfPath, signatureData, [19, 21]);
          
          if (!signatureResult.success) {
            this.logger.warn(`Failed to add signature to PDF: ${signatureResult.error}`);
            // Continue anyway, the PDF was generated successfully
          } else {
            this.logger.log(`Signature added successfully to PDF: ${pdfPath}`);
          }
        }
        
        return {
          success: true,
          message: 'Successfully generated PDF',
          pdfPath
        };
      } else {
        this.logger.error(`PDF generation failed: ${result.error}`);
        return {
          success: false,
          message: 'Failed to generate PDF',
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error generating PDF: ${error.message}`);
      return {
        success: false,
        message: 'Error generating PDF',
        error: error.message
      };
    }
  }

  private createPDFGenerationScript(excelFilePath: string, pdfPath: string, signatureData?: string, calculatorType: 'regular' | 'epvs' = 'regular'): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    const pdfPathEscaped = pdfPath.replace(/\\/g, '\\\\');
    
    return `
# Generate PDF from Excel file
$ErrorActionPreference = "Stop"

# Configuration
$excelFilePath = "${excelFilePathEscaped}"
$pdfPath = "${pdfPathEscaped}"
$password = "${this.PASSWORD}"

Write-Host "Generating PDF from Excel file..." -ForegroundColor Green
Write-Host "Excel file: $excelFilePath" -ForegroundColor Yellow
Write-Host "PDF output: $pdfPath" -ForegroundColor Yellow

# Validate paths
if (!(Test-Path $excelFilePath)) {
    throw "Excel file not found: $excelFilePath"
}

# Clean and validate PDF path
$pdfPath = [System.IO.Path]::GetFullPath($pdfPath)
$pdfDir = Split-Path $pdfPath -Parent

# Ensure the PDF directory exists and is writable
if (!(Test-Path $pdfDir)) {
    try {
        New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
        Write-Host "Created PDF directory: $pdfDir" -ForegroundColor Green
    } catch {
        throw "Failed to create PDF directory: $pdfDir - $_"
    }
}

# Test write permissions
try {
    $testFile = Join-Path $pdfDir "test-write.tmp"
    "test" | Out-File -FilePath $testFile -Encoding ASCII
    Remove-Item $testFile -Force
    Write-Host "Write permissions verified for: $pdfDir" -ForegroundColor Green
} catch {
    throw "No write permissions for PDF directory: $pdfDir - $_"
}

$excel = $null
$workbook = $null

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros for payment method selection (needed for regular calculator)
    $excel.AutomationSecurity = ${calculatorType === 'regular' ? 1 : 3}  # Enable macros for regular, disable for EPVS
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook
    Write-Host "Opening Excel workbook..." -ForegroundColor Yellow
    try {
        $workbook = $excel.Workbooks.Open($excelFilePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
        # Enable calculations for regular calculator
        if ("${calculatorType}" -eq "regular") {
            $excel.Calculation = -4105  # xlCalculationAutomatic
        }
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($excelFilePath)
            Write-Host "Workbook opened without password" -ForegroundColor Green
            if ("${calculatorType}" -eq "regular") {
                $excel.Calculation = -4105  # xlCalculationAutomatic
            }
        } catch {
            Write-Host "Failed to open workbook: $_" -ForegroundColor Red
            throw "Could not open workbook: $excelFilePath"
        }
    }
    
    # For regular calculator: Check and restore payment method before PDF generation
    if ("${calculatorType}" -eq "regular") {
        Write-Host "Checking and restoring payment method..." -ForegroundColor Cyan
        
        $currentPaymentMethod = $null
        $paymentMethodValue = $null
        $newFinanceValue = $null
        $paymentTypeText = $null
        $solarProjectionsSheet = $null
        
        # Try to read from Lookups sheet
        try {
            $lookupsSheet = $workbook.Worksheets.Item("Lookups")
            if ($lookupsSheet) {
                try {
                    $paymentMethodValue = $lookupsSheet.Range("PaymentMethodOptionValue").Value2
                    Write-Host "PaymentMethodOptionValue from Lookups: $paymentMethodValue" -ForegroundColor Green
                } catch {
                    Write-Host "Could not read PaymentMethodOptionValue: $_" -ForegroundColor Yellow
                }
                
                try {
                    $newFinanceValue = $lookupsSheet.Range("NewFinance").Value2
                    Write-Host "NewFinance from Lookups: $newFinanceValue" -ForegroundColor Green
                } catch {
                    Write-Host "Could not read NewFinance: $_" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "Could not access Lookups sheet: $_" -ForegroundColor Yellow
        }
        
        # Also try to read from Solar Projections sheet
        try {
            $solarProjectionsSheet = $workbook.Worksheets.Item("Solar Projections")
            if ($solarProjectionsSheet) {
                try {
                    $paymentTypeText = $solarProjectionsSheet.Range("B7").Value2
                    Write-Host "Payment type from Solar Projections B7: $paymentTypeText" -ForegroundColor Green
                } catch {
                    Write-Host "Could not read payment type from B7: $_" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "Could not access Solar Projections sheet: $_" -ForegroundColor Yellow
        }
        
        # Determine payment method name from value
        if ($paymentMethodValue -ne $null -and [int]$paymentMethodValue -gt 0) {
            switch ([int]$paymentMethodValue) {
                1 { 
                    $currentPaymentMethod = "cash"
                    Write-Host "Detected payment method: CASH (value: 1)" -ForegroundColor Cyan
                }
                2 { 
                    $currentPaymentMethod = "hometree"
                    Write-Host "Detected payment method: HOMETREE (value: 2)" -ForegroundColor Cyan
                }
                3 { 
                    $currentPaymentMethod = "finance"
                    Write-Host "Detected payment method: FINANCE (value: 3)" -ForegroundColor Cyan
                }
                default {
                    Write-Host "Unknown payment method value ($paymentMethodValue), checking Solar Projections..." -ForegroundColor Yellow
                }
            }
        }
        
        # If still not determined, try to infer from Solar Projections text
        if (-not $currentPaymentMethod -or ($currentPaymentMethod -eq "hometree" -and [int]$paymentMethodValue -eq 0)) {
            if ($solarProjectionsSheet -and $paymentTypeText) {
                try {
                    $paymentTypeLower = $paymentTypeText.ToString().ToLower()
                    if ($paymentTypeLower -like "*cash*") {
                        $currentPaymentMethod = "cash"
                        Write-Host "Detected payment method from text: CASH" -ForegroundColor Cyan
                    } elseif ($paymentTypeLower -like "*finance*" -or $paymentTypeLower -like "*loan*") {
                        $currentPaymentMethod = "finance"
                        Write-Host "Detected payment method from text: FINANCE" -ForegroundColor Cyan
                    } elseif ($paymentTypeLower -like "*hometree*" -or $paymentTypeLower -like "*home tree*") {
                        $currentPaymentMethod = "hometree"
                        Write-Host "Detected payment method from text: HOMETREE" -ForegroundColor Cyan
                    }
                } catch {
                    Write-Host "Could not parse payment type from text: $_" -ForegroundColor Yellow
                }
            }
        }
        
        # Final fallback
        if (-not $currentPaymentMethod) {
            $currentPaymentMethod = "hometree"
            Write-Host "Could not determine payment method, defaulting to HOMETREE" -ForegroundColor Yellow
        }
        
        Write-Host "Current payment method to restore: $currentPaymentMethod" -ForegroundColor Green
        
        # Find the Inputs sheet and select payment method back
        $inputsSheet = $null
        try {
            $inputsSheet = $workbook.Worksheets.Item("Inputs")
            Write-Host "Found Inputs sheet" -ForegroundColor Green
        } catch {
            Write-Host "Inputs sheet not found, trying to find it..." -ForegroundColor Yellow
            $sheetCount = $workbook.Worksheets.Count
            for ($i = 1; $i -le $sheetCount; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                if ($sheetName -eq "Inputs") {
                    $inputsSheet = $workbook.Worksheets.Item($i)
                    Write-Host "Found Inputs sheet: $sheetName" -ForegroundColor Green
                    break
                }
            }
        }
        
        if ($inputsSheet) {
            # Try macro approach first
            try {
                Write-Host "Using macro approach to set payment method: $currentPaymentMethod" -ForegroundColor Green
                $inputsSheet.Select()
                Start-Sleep -Milliseconds 500
                
                # Run the appropriate macro
                switch ($currentPaymentMethod.ToLower()) {
                    "hometree" { 
                        $excel.Run("SetOptionHomeTree")
                        Write-Host "✅ SetOptionHomeTree macro executed" -ForegroundColor Green
                    }
                    "cash" { 
                        $excel.Run("SetOptionCash")
                        Write-Host "✅ SetOptionCash macro executed" -ForegroundColor Green
                    }
                    "finance" { 
                        $excel.Run("SetOptionNewFinance")
                        Write-Host "✅ SetOptionNewFinance macro executed" -ForegroundColor Green
                    }
                    default { 
                        $excel.Run("SetOptionHomeTree")
                        Write-Host "✅ SetOptionHomeTree macro executed (default)" -ForegroundColor Green
                    }
                }
                
                Start-Sleep -Seconds 2
                Write-Host "Payment method selection completed" -ForegroundColor Green
            } catch {
                Write-Host "⚠️ Macro approach failed: $_" -ForegroundColor Yellow
                Write-Host "Trying fallback method..." -ForegroundColor Yellow
                
                # Fallback: Set values directly in Lookups sheet
                try {
                    $lookupsSheet = $workbook.Worksheets.Item("Lookups")
                    if ($lookupsSheet) {
                        # Map payment method to value
                        $paymentMethodValueToSet = switch ($currentPaymentMethod.ToLower()) {
                            "hometree" { 2 }
                            "cash" { 1 }
                            "finance" { 3 }
                            default { 2 }
                        }
                        
                        $lookupsSheet.Range("PaymentMethodOptionValue").Value2 = $paymentMethodValueToSet
                        Write-Host "Set Lookups PaymentMethodOptionValue to $paymentMethodValueToSet" -ForegroundColor Green
                        
                        $newFinanceValueToSet = if ($currentPaymentMethod.ToLower() -eq "finance") { "True" } else { "False" }
                        $lookupsSheet.Range("NewFinance").Value2 = $newFinanceValueToSet
                        Write-Host "Set Lookups NewFinance to $newFinanceValueToSet" -ForegroundColor Green
                        
                        $excel.Calculate()
                        Start-Sleep -Seconds 2
                        Write-Host "Payment method set via fallback method" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "Fallback method also failed: $_" -ForegroundColor Red
                    Write-Host "Continuing with PDF generation anyway..." -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "⚠️ Inputs sheet not found, skipping payment method selection" -ForegroundColor Yellow
        }
    }
    
    # Export as PDF - Select the correct illustration worksheet
    Write-Host "Exporting to PDF..." -ForegroundColor Green
    
    # Determine target worksheet based on calculator type
    $targetWorksheetName = "${calculatorType === 'epvs' ? 'FLX-Illustrations' : '25 Year Illustrations'}"
    
    # Try to access worksheet directly by name (much faster than iterating)
    $targetWorksheet = $null
    try {
        $targetWorksheet = $workbook.Worksheets.Item($targetWorksheetName)
        Write-Host "Found ${calculatorType === 'epvs' ? 'EPVS' : 'regular'} calculator worksheet: $($targetWorksheet.Name)" -ForegroundColor Green
    } catch {
        # If direct access fails, fall back to iteration (shouldn't happen normally)
        Write-Host "Direct access failed, searching for worksheet: $targetWorksheetName" -ForegroundColor Yellow
        for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
            $ws = $workbook.Worksheets.Item($i)
            if ($ws.Name -eq $targetWorksheetName) {
                $targetWorksheet = $ws
                break
            }
        }
    }
    
    if ($targetWorksheet) {
        
        # Use Excel's Print to PDF functionality
        try {
            Write-Host "Using Excel's Print to PDF..." -ForegroundColor Cyan
            
            # Delete existing PDF if it exists
            if (Test-Path $pdfPath) {
                Remove-Item $pdfPath -Force
            }
            
            # Set calculation mode to manual for faster processing
            $originalCalculation = $excel.Calculation
            $excel.Calculation = -4135  # xlCalculationManual
            
            # Get the current printer to restore later
            $originalPrinter = $excel.ActivePrinter
            
            try {
                # Set the active printer to Microsoft Print to PDF
                # Try the most common one first (usually "Microsoft Print to PDF on Ne01:")
                $selectedPrinter = $null
                if ($originalPrinter -like "*Microsoft Print to PDF*") {
                    # Already set to PDF printer, use it
                    $selectedPrinter = $originalPrinter
                } else {
                    # Try the most common printer name first
                    try {
                        $excel.ActivePrinter = "Microsoft Print to PDF on Ne01:"
                        $selectedPrinter = "Microsoft Print to PDF on Ne01:"
                    } catch {
                        # Fallback to other variations if needed
                        $pdfPrinterNames = @("Microsoft Print to PDF", "Microsoft Print To PDF")
                        foreach ($printerName in $pdfPrinterNames) {
                            try {
                                $excel.ActivePrinter = $printerName
                                $selectedPrinter = $printerName
                                break
                            } catch {}
                        }
                        if (-not $selectedPrinter) {
                            $selectedPrinter = $excel.ActivePrinter
                        }
                    }
                }
                
                # Activate the worksheet to ensure it's ready for printing
                $targetWorksheet.Activate()
                
                # Print to PDF using worksheet's PrintOut method
                # Parameters: From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas
                # Using [Type]::Missing for From/To to use default page range
                $targetWorksheet.PrintOut([Type]::Missing, [Type]::Missing, 1, $false, $selectedPrinter, $true, $false, $pdfPath, $false)
                
                # Wait for PDF file to be created and have content (not just exist)
                $maxWait = 30
                $waited = 0
                $fileReady = $false
                while (-not $fileReady -and $waited -lt $maxWait) {
                    Start-Sleep -Milliseconds 300
                    if (Test-Path $pdfPath) {
                        $fileInfo = Get-Item $pdfPath -ErrorAction SilentlyContinue
                        if ($fileInfo -and $fileInfo.Length -gt 0) {
                            $fileReady = $true
                        }
                    }
                    $waited++
                }
                
                if (-not $fileReady) {
                    throw "PDF file was not created or is empty after waiting"
                }
                
                Write-Host "PDF generated successfully using Print to PDF" -ForegroundColor Green
            } finally {
                # Restore original calculation mode
                try {
                    $excel.Calculation = $originalCalculation
                } catch {
                    Write-Host "Warning: Could not restore calculation mode" -ForegroundColor Yellow
                }
                # Restore original printer
                try {
                    $excel.ActivePrinter = $originalPrinter
                } catch {
                    Write-Host "Warning: Could not restore original printer" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "Print to PDF failed: $_" -ForegroundColor Red
            throw "Failed to generate PDF using Print to PDF: $_"
        }
    } else {
        Write-Host "Error: $targetWorksheetName worksheet not found!" -ForegroundColor Red
        Write-Host "Available worksheets:" -ForegroundColor Yellow
        for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
            $wsName = $workbook.Worksheets.Item($i).Name
            Write-Host "   - $wsName" -ForegroundColor White
        }
        throw "Target worksheet not found: $targetWorksheetName"
    }
    
    # Verify PDF was created
    if (Test-Path $pdfPath) {
        $pdfSize = (Get-Item $pdfPath).Length
        Write-Host "PDF created successfully: $pdfPath (Size: $pdfSize bytes)" -ForegroundColor Green
    } else {
        throw "PDF file was not created"
    }
    
} catch {
    Write-Host "Critical error in PDF generation: $_" -ForegroundColor Red
    throw $_
} finally {
    Write-Host "Starting cleanup process..." -ForegroundColor Yellow
    
    # Always try to close Excel properly
    if ($workbook) {
        try {
            Write-Host "Closing workbook..." -ForegroundColor Yellow
            $workbook.Close($false)  # Don't save changes
        } catch {
            Write-Host "Warning: Error closing workbook: $_" -ForegroundColor Yellow
        }
    }
    
    if ($excel) {
        try {
            Write-Host "Quitting Excel..." -ForegroundColor Yellow
            $excel.Quit()
        } catch {
            Write-Host "Warning: Error quitting Excel: $_" -ForegroundColor Yellow
        }
    }
    
    # Force release all COM objects
    try {
        Write-Host "Releasing COM objects..." -ForegroundColor Yellow
        if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
    } catch {
        Write-Host "Warning: Error releasing COM objects: $_" -ForegroundColor Yellow
    }
    
    # Force garbage collection (optimized - single pass)
    try {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {
        Write-Host "Warning: Error in garbage collection: $_" -ForegroundColor Yellow
    }
    
    # Force kill any remaining Excel processes
    try {
        Write-Host "Killing any remaining Excel processes..." -ForegroundColor Yellow
        Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
        Write-Host "Excel processes killed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error killing Excel processes: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Cleanup completed!" -ForegroundColor Green
}

Write-Host "PDF generation completed successfully!" -ForegroundColor Green
exit 0
`;
  }

  /**
   * Auto-populate Excel sheet with OpenSolar data
   */
  async autoPopulateWithOpenSolarData(
    opportunityId: string, 
    openSolarData: any, 
    templateFileName?: string
  ): Promise<{ success: boolean; message: string; data?: any; error?: string }> {
    try {
      this.logger.log(`🤖 Auto-populating Excel with OpenSolar data for opportunity: ${opportunityId}`);
      
      // Get the Excel file path
      const excelFilePath = this.findLatestOpportunityFile(opportunityId, 'off-peak') || path.join(this.OPPORTUNITIES_FOLDER, `Off peak V2.1 Eon SEG-${opportunityId}.xlsm`);
      
      if (!fs.existsSync(excelFilePath)) {
        return {
          success: false,
          message: `Excel file not found: ${excelFilePath}`,
          error: 'File not found'
        };
      }

      // Get all dropdown options from Excel
      const dropdownOptions = await this.getAllDropdownOptionsForFrontend(opportunityId, templateFileName);
      
      if (!dropdownOptions.success || !dropdownOptions.dropdownOptions) {
        return {
          success: false,
          message: 'Failed to get dropdown options from Excel',
          error: dropdownOptions.error
        };
      }

      // Prepare the data to populate
      const populateData: Record<string, string> = {};
      const matchedFields: string[] = [];
      const unmatchedFields: string[] = [];

      // Match panel manufacturer and model
      if (openSolarData.panel_manufacturer) {
        const panelManufacturerOptions = dropdownOptions.dropdownOptions['panel_manufacturer'] || [];
        this.logger.log(`🔍 Panel manufacturer from OpenSolar: "${openSolarData.panel_manufacturer}"`);
        this.logger.log(`🔍 Available Excel options: ${panelManufacturerOptions.join(', ')}`);
        
        // If no Excel options available, try common field names
        let allManufacturerOptions = [...panelManufacturerOptions];
        if (panelManufacturerOptions.length === 0) {
          this.logger.log(`⚠️ No panel_manufacturer options found, trying alternative field names`);
          const alternativeFields = ['panel_manufacturer', 'manufacturer', 'panel_make', 'make'];
          for (const fieldName of alternativeFields) {
            const altOptions = dropdownOptions.dropdownOptions[fieldName] || [];
            if (altOptions.length > 0) {
              this.logger.log(`✅ Found options in field "${fieldName}": ${altOptions.join(', ')}`);
              allManufacturerOptions = [...allManufacturerOptions, ...altOptions];
              break;
            }
          }
        }
        
        const matchedManufacturer = this.findBestMatch(openSolarData.panel_manufacturer, allManufacturerOptions);
        this.logger.log(`🔍 Match result: "${matchedManufacturer}"`);
        
        if (matchedManufacturer) {
          populateData['panel_manufacturer'] = matchedManufacturer;
          matchedFields.push('panel_manufacturer');
          
          // Now match panel model based on manufacturer
          if (openSolarData.panel_model) {
            const panelModelOptions = dropdownOptions.dropdownOptions['panel_model'] || [];
            const matchedModel = this.findBestMatch(openSolarData.panel_model, panelModelOptions);
            
            if (matchedModel) {
              populateData['panel_model'] = matchedModel;
              matchedFields.push('panel_model');
            } else {
              unmatchedFields.push('panel_model');
            }
          }
        } else {
          // Try to extract manufacturer from panel model if available
          if (openSolarData.panel_model) {
            this.logger.log(`🔄 Trying to extract manufacturer from panel model: "${openSolarData.panel_model}"`);
            const extractedManufacturer = this.extractManufacturerFromModel(openSolarData.panel_model);
            if (extractedManufacturer) {
              const retryMatch = this.findBestMatch(extractedManufacturer, allManufacturerOptions);
              if (retryMatch) {
                this.logger.log(`✅ Found manufacturer from model: "${retryMatch}"`);
                populateData['panel_manufacturer'] = retryMatch;
                matchedFields.push('panel_manufacturer');
              } else {
                unmatchedFields.push('panel_manufacturer');
              }
            } else {
              unmatchedFields.push('panel_manufacturer');
            }
          } else {
            unmatchedFields.push('panel_manufacturer');
          }
        }
      }

      // Match battery manufacturer and model
      if (openSolarData.battery_manufacturer) {
        const batteryManufacturerOptions = dropdownOptions.dropdownOptions['battery_manufacturer'] || [];
        const matchedManufacturer = this.findBestMatch(openSolarData.battery_manufacturer, batteryManufacturerOptions);
        
        if (matchedManufacturer) {
          populateData['battery_manufacturer'] = matchedManufacturer;
          matchedFields.push('battery_manufacturer');
          
          // Now match battery model based on manufacturer
          if (openSolarData.battery_model) {
            const batteryModelOptions = dropdownOptions.dropdownOptions['battery_model'] || [];
            const matchedModel = this.findBestMatch(openSolarData.battery_model, batteryModelOptions);
            
            if (matchedModel) {
              populateData['battery_model'] = matchedModel;
              matchedFields.push('battery_model');
            } else {
              unmatchedFields.push('battery_model');
            }
          }
        } else {
          unmatchedFields.push('battery_manufacturer');
        }
      }

      // Match solar inverter manufacturer and model
      if (openSolarData.solar_inverter_manufacturer) {
        const solarInverterManufacturerOptions = dropdownOptions.dropdownOptions['solar_inverter_manufacturer'] || [];
        const matchedManufacturer = this.findBestMatch(openSolarData.solar_inverter_manufacturer, solarInverterManufacturerOptions);
        
        if (matchedManufacturer) {
          populateData['solar_inverter_manufacturer'] = matchedManufacturer;
          matchedFields.push('solar_inverter_manufacturer');
          
          // Now match solar inverter model based on manufacturer
          if (openSolarData.solar_inverter_model) {
            const solarInverterModelOptions = dropdownOptions.dropdownOptions['solar_inverter_model'] || [];
            const matchedModel = this.findBestMatch(openSolarData.solar_inverter_model, solarInverterModelOptions);
            
            if (matchedModel) {
              populateData['solar_inverter_model'] = matchedModel;
              matchedFields.push('solar_inverter_model');
            } else {
              unmatchedFields.push('solar_inverter_model');
            }
          }
        } else {
          unmatchedFields.push('solar_inverter_manufacturer');
        }
      }

      // Match battery inverter manufacturer and model
      if (openSolarData.battery_inverter_manufacturer) {
        const batteryInverterManufacturerOptions = dropdownOptions.dropdownOptions['battery_inverter_manufacturer'] || [];
        const matchedManufacturer = this.findBestMatch(openSolarData.battery_inverter_manufacturer, batteryInverterManufacturerOptions);
        
        if (matchedManufacturer) {
          populateData['battery_inverter_manufacturer'] = matchedManufacturer;
          matchedFields.push('battery_inverter_manufacturer');
          
          // Now match battery inverter model based on manufacturer
          if (openSolarData.battery_inverter_model) {
            const batteryInverterModelOptions = dropdownOptions.dropdownOptions['battery_inverter_model'] || [];
            const matchedModel = this.findBestMatch(openSolarData.battery_inverter_model, batteryInverterModelOptions);
            
            if (matchedModel) {
              populateData['battery_inverter_model'] = matchedModel;
              matchedFields.push('battery_inverter_model');
            } else {
              unmatchedFields.push('battery_inverter_model');
            }
          }
        } else {
          unmatchedFields.push('battery_inverter_manufacturer');
        }
      }

      // Add numeric values directly
      if (openSolarData.panel_quantity) {
        populateData['panel_quantity'] = openSolarData.panel_quantity.toString();
        matchedFields.push('panel_quantity');
      }

      if (openSolarData.panel_wattage) {
        populateData['panel_wattage'] = openSolarData.panel_wattage.toString();
        matchedFields.push('panel_wattage');
      }

      if (openSolarData.system_size_kw) {
        populateData['system_size_kw'] = openSolarData.system_size_kw.toString();
        matchedFields.push('system_size_kw');
      }

      if (openSolarData.battery_capacity) {
        populateData['battery_capacity'] = openSolarData.battery_capacity.toString();
        matchedFields.push('battery_capacity');
      }

      // Save the populated data to Excel
      if (Object.keys(populateData).length > 0) {
        const saveResult = await this.saveDynamicInputs(opportunityId, populateData, templateFileName);

        if (saveResult.success) {
          this.logger.log(`✅ Successfully auto-populated ${Object.keys(populateData).length} fields`);
          this.logger.log(`✅ Matched fields: ${matchedFields.join(', ')}`);
          if (unmatchedFields.length > 0) {
            this.logger.log(`⚠️ Unmatched fields: ${unmatchedFields.join(', ')}`);
          }

          return {
            success: true,
            message: `Auto-populated ${Object.keys(populateData).length} fields with OpenSolar data`,
            data: {
              populatedFields: populateData,
              matchedFields,
              unmatchedFields
            }
          };
        } else {
          return {
            success: false,
            message: 'Failed to save populated data to Excel',
            error: saveResult.error
          };
        }
      } else {
        return {
          success: false,
          message: 'No data could be matched and populated',
          error: 'No matches found'
        };
      }

    } catch (error) {
      this.logger.error(`❌ Error auto-populating Excel with OpenSolar data:`, error.message);
      return {
        success: false,
        message: 'Failed to auto-populate Excel with OpenSolar data',
        error: error.message
      };
    }
  }

  /**
   * Find the best match for a value in a list of options
   * Enhanced to handle any case with multiple matching strategies
   */
  private findBestMatch(value: string, options: string[]): string | null {
    if (!value || !options || options.length === 0) {
      this.logger.log(`❌ findBestMatch: Invalid input - value: "${value}", options length: ${options?.length || 0}`);
      return null;
    }

    const normalizedValue = value.toLowerCase().trim();
    this.logger.log(`🔍 findBestMatch: Normalized value: "${normalizedValue}"`);
    this.logger.log(`🔍 Available options: ${options.slice(0, 10).join(', ')}${options.length > 10 ? '...' : ''}`);
    
    // Strategy 1: Exact match (case-insensitive)
    const exactMatch = options.find(option => 
      option.toLowerCase().trim() === normalizedValue
    );
    if (exactMatch) {
      this.logger.log(`✅ Exact match found: "${exactMatch}"`);
      return exactMatch;
    }

    // Strategy 2: Partial match (contains)
    const partialMatch = options.find(option => 
      option.toLowerCase().includes(normalizedValue) || 
      normalizedValue.includes(option.toLowerCase())
    );
    if (partialMatch) {
      this.logger.log(`✅ Partial match found: "${partialMatch}"`);
      return partialMatch;
    }

    // Strategy 3: Word-based matching (split by spaces, hyphens, etc.)
    const valueWords = normalizedValue.split(/[\s\-_\.]+/).filter(word => word.length > 2);
    if (valueWords.length > 0) {
      this.logger.log(`🔍 Trying word-based matching with words: ${valueWords.join(', ')}`);
      
      for (const word of valueWords) {
        const wordMatch = options.find(option => 
          option.toLowerCase().includes(word)
        );
        if (wordMatch) {
          this.logger.log(`✅ Word-based match found: "${wordMatch}" (matched word: "${word}")`);
          return wordMatch;
        }
      }
    }

    // Strategy 4: Fuzzy matching with similarity scoring
    const fuzzyMatches = options.map(option => {
      const optionLower = option.toLowerCase();
      let score = 0;
      
      // Check for common substrings
      if (optionLower.includes(normalizedValue) || normalizedValue.includes(optionLower)) {
        score += 50;
      }
      
      // Check for word overlap
      const optionWords = optionLower.split(/[\s\-_\.]+/);
      const valueWords = normalizedValue.split(/[\s\-_\.]+/);
      const commonWords = optionWords.filter(word => valueWords.includes(word));
      score += commonWords.length * 20;
      
      // Check for character similarity
      let charMatches = 0;
      for (let i = 0; i < Math.min(normalizedValue.length, optionLower.length); i++) {
        if (normalizedValue[i] === optionLower[i]) charMatches++;
      }
      score += (charMatches / Math.max(normalizedValue.length, optionLower.length)) * 30;
      
      return { option, score };
    }).filter(match => match.score > 20).sort((a, b) => b.score - a.score);
    
    if (fuzzyMatches.length > 0) {
      const bestMatch = fuzzyMatches[0];
      this.logger.log(`✅ Fuzzy match found: "${bestMatch.option}" (score: ${bestMatch.score})`);
      return bestMatch.option;
    }

    // Strategy 5: Manufacturer-specific patterns
    const manufacturerPatterns = [
      { pattern: /jinko/i, matches: ['Jinko', 'Jinko Solar'] },
      { pattern: /trina/i, matches: ['Trina', 'Trina Solar'] },
      { pattern: /canadian\s*solar/i, matches: ['Canadian Solar'] },
      { pattern: /longi/i, matches: ['Longi'] },
      { pattern: /ja\s*solar/i, matches: ['JA Solar'] },
      { pattern: /risen/i, matches: ['Risen'] },
      { pattern: /q\s*cells/i, matches: ['Q Cells', 'Hanwha Q Cells'] },
      { pattern: /sunpower/i, matches: ['SunPower'] },
      { pattern: /lg/i, matches: ['LG', 'LG Chem'] },
      { pattern: /panasonic/i, matches: ['Panasonic'] },
      { pattern: /rec/i, matches: ['REC'] },
      { pattern: /solarworld/i, matches: ['SolarWorld'] },
      { pattern: /first\s*solar/i, matches: ['First Solar'] },
      { pattern: /yingli/i, matches: ['Yingli'] },
      { pattern: /v-tac/i, matches: ['V-TAC'] },
      { pattern: /solis/i, matches: ['Solis'] },
      { pattern: /fronius/i, matches: ['Fronius'] },
      { pattern: /sma/i, matches: ['SMA'] },
      { pattern: /victron/i, matches: ['Victron'] },
      { pattern: /tesla/i, matches: ['Tesla'] },
      { pattern: /sonnen/i, matches: ['Sonnen'] },
      { pattern: /pylontech/i, matches: ['Pylontech'] },
      { pattern: /aiko/i, matches: ['Aiko'] },
      { pattern: /amerisolar/i, matches: ['AmeriSolar'] },
      { pattern: /das\s*solar/i, matches: ['DAS Solar'] },
      { pattern: /dmegc/i, matches: ['DMEGC Solar'] },
      { pattern: /energizer/i, matches: ['Energizer'] },
      { pattern: /eurener/i, matches: ['Eurener'] },
      { pattern: /evolution/i, matches: ['Evolution'] },
      { pattern: /exiom/i, matches: ['Exiom Solutions'] },
      { pattern: /hyundai/i, matches: ['Hyundai'] },
      { pattern: /meyer\s*burger/i, matches: ['Meyer Burger'] },
      { pattern: /perlight/i, matches: ['Perlight Solar'] },
      { pattern: /sharp/i, matches: ['Sharp'] },
      { pattern: /solarwatt/i, matches: ['Solarwatt'] },
      { pattern: /sunket/i, matches: ['Sunket'] },
      { pattern: /sunrise/i, matches: ['Sunrise Energy'] },
      { pattern: /suntech/i, matches: ['Suntech'] },
      { pattern: /tenka/i, matches: ['Tenka Solar'] },
      { pattern: /tongwei/i, matches: ['Tongwei'] },
      { pattern: /uksol/i, matches: ['UKSOL'] },
      { pattern: /ulica/i, matches: ['Ulica'] }
    ];

    for (const { pattern, matches } of manufacturerPatterns) {
      this.logger.log(`🔍 Testing pattern: ${pattern} for matches: ${matches.join(', ')}`);
      if (pattern.test(normalizedValue)) {
        this.logger.log(`✅ Pattern matched! Looking for matches in options`);
        
        // Try each possible match
        for (const match of matches) {
          const manufacturerOption = options.find(option => 
            option.toLowerCase().includes(match.toLowerCase())
          );
          if (manufacturerOption) {
            this.logger.log(`✅ Found manufacturer option: "${manufacturerOption}"`);
            return manufacturerOption;
          }
        }
        
        this.logger.log(`❌ No matching option found for pattern matches`);
      }
    }

    // Strategy 6: Acronym matching (e.g., "LG" from "LG Chem")
    const acronyms = normalizedValue.match(/\b[A-Z]{2,}\b/g);
    if (acronyms) {
      this.logger.log(`🔍 Trying acronym matching: ${acronyms.join(', ')}`);
      for (const acronym of acronyms) {
        const acronymMatch = options.find(option => 
          option.toLowerCase().includes(acronym.toLowerCase())
        );
        if (acronymMatch) {
          this.logger.log(`✅ Acronym match found: "${acronymMatch}" (acronym: "${acronym}")`);
          return acronymMatch;
        }
      }
    }

    this.logger.log(`❌ No match found for "${normalizedValue}"`);
    return null;
  }

  /**
   * Extract manufacturer name from panel model
   * Enhanced to handle any case with multiple extraction strategies
   */
  private extractManufacturerFromModel(modelName: string): string | null {
    if (!modelName) return null;
    
    this.logger.log(`🔍 Extracting manufacturer from model: "${modelName}"`);
    
    // Strategy 1: Direct manufacturer patterns
    const manufacturerPatterns = [
      { pattern: /jinko/i, name: 'Jinko Solar' },
      { pattern: /trina/i, name: 'Trina Solar' },
      { pattern: /canadian\s*solar/i, name: 'Canadian Solar' },
      { pattern: /longi/i, name: 'Longi' },
      { pattern: /ja\s*solar/i, name: 'JA Solar' },
      { pattern: /risen/i, name: 'Risen' },
      { pattern: /q\s*cells/i, name: 'Hanwha Q Cells' },
      { pattern: /sunpower/i, name: 'SunPower' },
      { pattern: /lg/i, name: 'LG' },
      { pattern: /panasonic/i, name: 'Panasonic' },
      { pattern: /rec/i, name: 'REC' },
      { pattern: /solarworld/i, name: 'SolarWorld' },
      { pattern: /first\s*solar/i, name: 'First Solar' },
      { pattern: /yingli/i, name: 'Yingli' },
      { pattern: /v-tac/i, name: 'V-TAC' },
      { pattern: /solis/i, name: 'Solis' },
      { pattern: /fronius/i, name: 'Fronius' },
      { pattern: /sma/i, name: 'SMA' },
      { pattern: /victron/i, name: 'Victron' },
      { pattern: /tesla/i, name: 'Tesla' },
      { pattern: /sonnen/i, name: 'Sonnen' },
      { pattern: /pylontech/i, name: 'Pylontech' },
      { pattern: /aiko/i, name: 'Aiko' },
      { pattern: /amerisolar/i, name: 'AmeriSolar' },
      { pattern: /das\s*solar/i, name: 'DAS Solar' },
      { pattern: /dmegc/i, name: 'DMEGC Solar' },
      { pattern: /energizer/i, name: 'Energizer' },
      { pattern: /eurener/i, name: 'Eurener' },
      { pattern: /evolution/i, name: 'Evolution' },
      { pattern: /exiom/i, name: 'Exiom Solutions' },
      { pattern: /hyundai/i, name: 'Hyundai' },
      { pattern: /meyer\s*burger/i, name: 'Meyer Burger' },
      { pattern: /perlight/i, name: 'Perlight Solar' },
      { pattern: /sharp/i, name: 'Sharp' },
      { pattern: /solarwatt/i, name: 'Solarwatt' },
      { pattern: /sunket/i, name: 'Sunket' },
      { pattern: /sunrise/i, name: 'Sunrise Energy' },
      { pattern: /suntech/i, name: 'Suntech' },
      { pattern: /tenka/i, name: 'Tenka Solar' },
      { pattern: /tongwei/i, name: 'Tongwei' },
      { pattern: /uksol/i, name: 'UKSOL' },
      { pattern: /ulica/i, name: 'Ulica' }
    ];

    for (const { pattern, name } of manufacturerPatterns) {
      if (pattern.test(modelName)) {
        this.logger.log(`✅ Extracted manufacturer: "${name}" from pattern`);
        return name;
      }
    }

    // Strategy 2: Common manufacturer names in model
    const commonManufacturers = [
      'Jinko', 'Trina', 'Canadian Solar', 'Longi', 'JA Solar', 'Risen', 'Q Cells',
      'SunPower', 'LG', 'Panasonic', 'REC', 'SolarWorld', 'First Solar', 'Yingli',
      'V-TAC', 'Solis', 'Fronius', 'SMA', 'Victron', 'Tesla', 'Sonnen', 'LG Chem',
      'Pylontech', 'Huawei', 'Growatt', 'Enphase', 'BYD', 'Aiko', 'AmeriSolar',
      'DAS Solar', 'DMEGC Solar', 'Energizer', 'Eurener', 'Evolution', 'Exiom',
      'Hyundai', 'Meyer Burger', 'Perlight', 'Sharp', 'Solarwatt', 'Sunket',
      'Sunrise', 'Suntech', 'Tenka', 'Tongwei', 'UKSOL', 'Ulica'
    ];

    for (const manufacturer of commonManufacturers) {
      if (modelName.toLowerCase().includes(manufacturer.toLowerCase())) {
        this.logger.log(`✅ Extracted manufacturer: "${manufacturer}" from common names`);
        return manufacturer;
      }
    }

    // Strategy 3: Extract from model prefix (common pattern: MANUFACTURER-MODEL)
    const modelParts = modelName.split(/[\s\-_\.]+/);
    if (modelParts.length > 0) {
      const firstPart = modelParts[0];
      // Check if first part looks like a manufacturer (not just numbers or generic terms)
      if (firstPart.length > 2 && !/^\d+$/.test(firstPart) && !['Tier', 'Model', 'Panel', 'Solar'].includes(firstPart)) {
        this.logger.log(`✅ Extracted manufacturer from model prefix: "${firstPart}"`);
        return firstPart;
      }
    }

    // Strategy 4: Look for manufacturer in model description
    const modelLower = modelName.toLowerCase();
    if (modelLower.includes('tiger neo') || modelLower.includes('tiger')) {
      this.logger.log(`✅ Extracted manufacturer from Tiger Neo: "Jinko Solar"`);
      return 'Jinko Solar';
    }
    if (modelLower.includes('hi-mo') || modelLower.includes('him')) {
      this.logger.log(`✅ Extracted manufacturer from Hi-Mo: "Longi"`);
      return 'Longi';
    }
    if (modelLower.includes('bi-hiku') || modelLower.includes('bihiku')) {
      this.logger.log(`✅ Extracted manufacturer from BiHiKu: "Canadian Solar"`);
      return 'Canadian Solar';
    }

    this.logger.log(`❌ Could not extract manufacturer from model: "${modelName}"`);
    return null;
  }

  /**
   * Write a value to a specific cell in an Excel file
   */
  async writeCellValue(opportunityId: string, cellReference: string, value: any, templateFileName?: string): Promise<boolean> {
    this.logger.log(`Writing value ${value} to cell ${cellReference} for opportunity: ${opportunityId}${templateFileName ? ` using template: ${templateFileName}` : ''}`);

    try {
      // Determine which file to use
      let excelFilePath: string;
      
      if (templateFileName) {
        // Use the specified template file
        excelFilePath = this.getTemplateFilePath(templateFileName);
        this.logger.log(`Using template file: ${excelFilePath}`);
      } else {
        // Try to find opportunity-specific file first
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId, 'off-peak');
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using opportunity file: ${excelFilePath}`);
        } else {
          excelFilePath = this.getTemplateFilePath();
          this.logger.log(`Opportunity file not found, using default template: ${excelFilePath}`);
        }
      }

      // Check if file exists
      if (!fs.existsSync(excelFilePath)) {
        const error = `Excel file not found at: ${excelFilePath}`;
        this.logger.error(error);
        return false;
      }

      // Check if we're on Windows (required for Excel COM automation)
      if (process.platform !== 'win32') {
        const error = 'Excel automation requires Windows platform';
        this.logger.error(error);
        return false;
      }

      // Create PowerShell script
      const psScript = this.createWriteCellScript(cellReference, value, excelFilePath);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-write-cell-${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created temporary PowerShell script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Failed to clean up temporary script file: ${cleanupError}`);
      }

      if (result.success) {
        this.logger.log(`Successfully wrote value ${value} to cell ${cellReference}`);
        return true;
      } else {
        this.logger.error(`Failed to write value to cell: ${result.error}`);
        return false;
      }

    } catch (error) {
      this.logger.error(`Error writing value to cell ${cellReference}:`, error);
      return false;
    }
  }

  /**
   * Create PowerShell script to write a value to a specific cell
   */
  private createWriteCellScript(cellReference: string, value: any, excelFilePath: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    return `
# Write Value to Excel Cell
$ErrorActionPreference = "Stop"

# Configuration
$excelFilePath = "${excelFilePathEscaped}"
$cellReference = "${cellReference}"
$value = "${value}"

Write-Host "Writing value to Excel cell: $cellReference" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook
    Write-Host "Opening Excel workbook..." -ForegroundColor Yellow
    try {
        $workbook = $excel.Workbooks.Open($excelFilePath, 0, $false, 5, "")
        Write-Host "Workbook opened successfully without password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open workbook: $_" -ForegroundColor Red
        throw "Could not open workbook: $excelFilePath"
    }
    
    # Get the first worksheet
    $worksheet = $workbook.Worksheets.Item(1)
    if (!$worksheet) {
        throw "Worksheet not found"
    }
    Write-Host "Found worksheet" -ForegroundColor Green
    
    # Write value to the specified cell
    try {
        $worksheet.Range($cellReference).Value = $value
        Write-Host "Value written successfully to cell $cellReference" -ForegroundColor Green
    } catch {
        Write-Host "Failed to write value to cell $cellReference: $_" -ForegroundColor Red
        throw "Failed to write value to cell"
    }
    
    # Save the workbook
    try {
        $workbook.Save()
        Write-Host "Workbook saved successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to save workbook: $_" -ForegroundColor Red
        throw "Failed to save workbook"
    }
    
    # Close Excel
    try {
        $workbook.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    return $true
} catch {
    Write-Host "Critical error in writing value to cell: $_" -ForegroundColor Red
    return $false
}
`;
  }

  /**
   * Perform complete calculation with user session isolation
   */
  async performCompleteCalculationWithSession(
    userId: string,
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string; pdfPath?: string }> {
    this.logger.log(`Performing complete calculation with session isolation for user: ${userId}, opportunity: ${opportunityId}`);

    try {
      // Queue the request through session management
      const result = await this.sessionManagementService.queueRequest(
        userId,
        'excel_calculation',
        'com',
        {
          opportunityId,
          customerDetails,
          radioButtonSelections,
          dynamicInputs,
          templateFileName
        },
        1 // High priority
      );

      return result;
    } catch (error) {
      this.logger.error(`Session-based calculation failed for user ${userId}:`, error);
      return { 
        success: false, 
        message: 'Calculation failed', 
        error: error.message 
      };
    }
  }

  /**
   * Execute Excel calculation with user isolation (called by session management)
   */
  async executeExcelCalculationWithIsolation(
    userId: string,
    data: {
      opportunityId: string;
      customerDetails: { customerName: string; address: string; postcode: string };
      radioButtonSelections: string[];
      dynamicInputs?: Record<string, string>;
      templateFileName?: string;
    }
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string; pdfPath?: string }> {
    const session = await this.sessionManagementService.createOrGetSession(userId);
    const workingDirectory = session.workingDirectory;

    this.logger.log(`Executing Excel calculation with isolation for user: ${userId}`);

    try {
      // Check if template file exists
      const templateFilePath = this.getTemplateFilePath(data.templateFileName);
      if (!fs.existsSync(templateFilePath)) {
        const error = `Template file not found at: ${templateFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Check if we're on Windows (required for Excel COM automation)
      if (process.platform !== 'win32') {
        const error = 'Excel automation requires Windows platform';
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Create user-specific file paths
      const userExcelDir = path.join(workingDirectory, 'excel');
      const newFilePath = path.join(userExcelDir, `calculation_${data.opportunityId}_${Date.now()}.xlsm`);
      const pdfPath = path.join(workingDirectory, 'pdf', `calculation_${data.opportunityId}_${Date.now()}.pdf`);

      // Create user-specific PowerShell script
      const psScript = this.createIsolatedCalculationScript(
        data.opportunityId,
        data.customerDetails,
        data.radioButtonSelections,
        data.dynamicInputs,
        data.templateFileName,
        templateFilePath,
        newFilePath,
        workingDirectory
      );
      
      // Create temporary script file in user's directory
      const tempScriptPath = path.join(workingDirectory, 'temp', `calculation_${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created isolated PowerShell script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Could not clean up temporary script: ${cleanupError.message}`);
      }

      if (result.success) {
        return {
          success: true,
          message: 'Calculation completed successfully with user isolation',
          filePath: newFilePath,
          pdfPath: pdfPath
        };
      } else {
        return {
          success: false,
          message: 'Calculation failed',
          error: result.error || 'Unknown error'
        };
      }
    } catch (error) {
      this.logger.error(`Isolated calculation failed for user ${userId}:`, error);
      return { 
        success: false, 
        message: 'Calculation failed', 
        error: error.message 
      };
    }
  }

  /**
   * Create isolated calculation script for user session
   */
  private createIsolatedCalculationScript(
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string,
    templatePath?: string,
    newFilePath?: string,
    workingDirectory?: string
  ): string {
    const radioButtonsString = radioButtonSelections.map(selection => `"${this.escapeForPowerShell(selection)}"`).join(',');
    
    let inputsString = '';
    if (dynamicInputs) {
      inputsString = Object.entries(dynamicInputs)
        .map(([key, value]) => `    "${key}" = "${this.escapeForPowerShell(value)}"`)
        .join('\n');
    }
    
    return `
# Isolated Excel Calculation - User Session
$ErrorActionPreference = "Stop"

# Configuration
$templatePath = "${templatePath?.replace(/\\/g, '\\\\')}"
$newFilePath = "${newFilePath?.replace(/\\/g, '\\\\')}"
$workingDirectory = "${workingDirectory?.replace(/\\/g, '\\\\')}"
$password = "${this.PASSWORD}"
$opportunityId = "${opportunityId}"

# Radio button selections
$radioButtonSelections = @(${radioButtonsString})

# Dynamic inputs (if any)
$dynamicInputs = @{
${inputsString}
}

Write-Host "Starting isolated calculation for: $opportunityId" -ForegroundColor Green
Write-Host "Working Directory: $workingDirectory" -ForegroundColor Yellow
Write-Host "Template path: $templatePath" -ForegroundColor Yellow
Write-Host "New file path: $newFilePath" -ForegroundColor Yellow

try {
    # Create user-specific Excel application
    Write-Host "Creating isolated Excel application..." -ForegroundColor Green
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Isolated Excel application created successfully" -ForegroundColor Green
    
    # Copy template to user's directory
    Write-Host "Copying template file..." -ForegroundColor Green
    Copy-Item -Path $templatePath -Destination $newFilePath -Force
    Write-Host "Template file copied successfully" -ForegroundColor Green
    
    # Open the copied file
    Write-Host "Opening workbook..." -ForegroundColor Green
    $workbook = $excel.Workbooks.Open($newFilePath, 0, $false, 5, $password)
    Write-Host "Workbook opened successfully" -ForegroundColor Green
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
    # Unprotect all worksheets
    Write-Host "Unprotecting worksheets..." -ForegroundColor Green
    foreach ($ws in $workbook.Worksheets) {
        try {
            if ($ws.ProtectContents) {
                $ws.Unprotect($password)
                Write-Host "Unprotected worksheet: $($ws.Name)" -ForegroundColor Green
            }
        } catch {
            Write-Host "Warning: Could not unprotect worksheet $($ws.Name)" -ForegroundColor Yellow
        }
    }
    
    # Fill in customer details
    Write-Host "Filling customer details..." -ForegroundColor Green
    $worksheet.Range("H12").Value = "${this.escapeForPowerShell(customerDetails.customerName)}"
    $worksheet.Range("H13").Value = "${this.escapeForPowerShell(customerDetails.address)}"
    $worksheet.Range("H14").Value = "${this.escapeForPowerShell(customerDetails.postcode)}"
    Write-Host "Customer details filled successfully" -ForegroundColor Green
    
    # Save the workbook
    Write-Host "Saving workbook..." -ForegroundColor Green
    $workbook.Save()
    Write-Host "Workbook saved successfully" -ForegroundColor Green
    
    # Close workbook
    $workbook.Close($false)
    Write-Host "Workbook closed successfully" -ForegroundColor Green
    
    Write-Host "Isolated calculation completed successfully!" -ForegroundColor Green
    Write-Host "Output file: $newFilePath" -ForegroundColor Green
    
} catch {
    Write-Host "Error in isolated calculation: $_" -ForegroundColor Red
    throw $_
} finally {
    Write-Host "Cleaning up isolated Excel process..." -ForegroundColor Yellow
    
    if ($excel) {
        try {
            $excel.Quit()
        } catch {
            Write-Host "Warning: Error quitting Excel: $_" -ForegroundColor Yellow
        }
    }
    
    # Force release COM objects
    try {
        if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {
        Write-Host "Warning: Error releasing COM objects: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Isolated Excel cleanup completed" -ForegroundColor Green
}

Write-Host "Isolated calculation process completed!" -ForegroundColor Green
exit 0
`;
  }

  /**
   * Create PowerShell script to retrieve pricing data from Excel
   */
  private createGetPricingDataScript(excelFilePath: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    return `
# Get Pricing Data from Excel
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePathEscaped}"
$password = "${this.PASSWORD}"

Write-Host "Getting pricing data from: $filePath" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        $workbook = $excel.Workbooks.Open($filePath)
        Write-Host "Workbook opened without password" -ForegroundColor Green
    }
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    if (!$worksheet) {
        throw "Inputs worksheet not found"
    }
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
    # Define pricing field mappings
    $pricingFields = @{
        "total_system_cost" = "H80"  # Default to off-peak, will be adjusted based on calculator type
        "deposit" = "H81"
        "interest_rate" = "H82"
        "interest_rate_type" = "H83"
        "payment_term" = "H84"
        "payment_method" = "H85"
    }
    
    # Check if this is an EPVS calculator by looking for EPVS-specific cells
    $isEPVS = $false
    try {
        $epvsCell = $worksheet.Range("H81")
        if ($epvsCell.Value -ne $null) {
            $isEPVS = $true
            Write-Host "Detected EPVS calculator, adjusting cell mappings" -ForegroundColor Cyan
            $pricingFields = @{
                "total_system_cost" = "H81"
                "deposit" = "H82"
                "interest_rate" = "H83"
                "interest_rate_type" = "H84"
                "payment_term" = "H85"
                "payment_method" = "H86"
            }
        }
    } catch {
        Write-Host "Using default off-peak calculator mappings" -ForegroundColor Cyan
    }
    
    Write-Host "Retrieving pricing data..." -ForegroundColor Yellow
    
    # DEBUG: Scan a wider range of cells to see what's actually there
    Write-Host "DEBUG: Scanning pricing-related cells..." -ForegroundColor Magenta
    for ($row = 80; $row -le 90; $row++) {
        for ($col = 7; $col -le 9; $col++) {  # Columns H, I, J
            try {
                $cellRef = [char](65 + $col) + $row  # Convert to A1 notation
                $cell = $worksheet.Range($cellRef)
                $value = $cell.Value
                $text = $cell.Text
                if ($value -ne $null -and $value -ne "" -and $value -ne 0) {
                    Write-Host "DEBUG: Found data in $cellRef - Value: '$value', Text: '$text'" -ForegroundColor Magenta
                }
            } catch {
                # Ignore errors for debug scan
            }
        }
    }
    
    $pricingData = @{}
    
    foreach ($field in $pricingFields.GetEnumerator()) {
        try {
            $cell = $worksheet.Range($field.Value)
            $value = $cell.Value
            $formula = $cell.Formula
            $text = $cell.Text
            
            Write-Host "DEBUG: Cell $($field.Value) - Value: '$value', Formula: '$formula', Text: '$text'" -ForegroundColor Cyan
            
            if ($value -ne $null -and $value -ne "") {
                $pricingData[$field.Key] = $value.ToString()
                Write-Host "Retrieved $($field.Key): $value" -ForegroundColor Green
            } else {
                $pricingData[$field.Key] = ""
                Write-Host "Retrieved $($field.Key): (empty)" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Error retrieving $($field.Key): $_" -ForegroundColor Red
            $pricingData[$field.Key] = ""
        }
    }
    
    # Add calculator type information
    $pricingData["calculator_type"] = if ($isEPVS) { "epvs" } else { "off-peak" }
    
    # Convert to JSON
    $jsonResult = $pricingData | ConvertTo-Json -Depth 2 -Compress
    Write-Host "RESULT: $jsonResult" -ForegroundColor Green
    
    # Close Excel
    try {
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Pricing data retrieval completed successfully!" -ForegroundColor Green
    exit 0
    
} catch {
    Write-Host "Critical error in pricing data retrieval: $_" -ForegroundColor Red
    exit 1
}
`;
  }
}
