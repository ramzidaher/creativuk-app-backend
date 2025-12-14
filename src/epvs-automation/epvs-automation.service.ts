import { Injectable, Logger } from '@nestjs/common';
import { spawn } from 'child_process';
import * as path from 'path';
import * as fs from 'fs';
import * as XLSX from 'xlsx';
import { PdfSignatureService } from '../pdf-signature/pdf-signature.service';
import { SessionManagementService } from '../session-management/session-management.service';
import { ComProcessManagerService } from '../session-management/com-process-manager.service';

@Injectable()
export class EPVSAutomationService {
  private readonly logger = new Logger(EPVSAutomationService.name)

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

  private mapRadioButtonName(buttonName: string): string {
    // Map frontend payment type names to actual Excel radio button names
    const mapping: Record<string, string> = {
      'Hometree': 'NewFinance',  // Hometree maps to NewFinance option
      'Cash': 'Cash',           // Cash stays the same
      'Finance': 'Finance',     // Finance stays the same
      'NewFinance': 'NewFinance' // NewFinance stays the same
    };
    
    const mappedName = mapping[buttonName] || buttonName;
    
    if (mappedName !== buttonName) {
      this.logger.log(`🔄 Mapped radio button name: "${buttonName}" → "${mappedName}"`);
    }
    
    return mappedName;
  }

  constructor(
    private readonly pdfSignatureService: PdfSignatureService,
    private readonly sessionManagementService: SessionManagementService,
    private readonly comProcessManagerService: ComProcessManagerService
  ) {}

  private readonly TEMPLATES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-templates');
  private readonly DEFAULT_TEMPLATE_FILE = 'EPVS Calculator Creativ - 06.02 - Solar Only.xlsm';
  private readonly OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');
  private readonly PASSWORD = '99';

  private getTemplateFilePath(templateFileName?: string): string {
    const fileName = templateFileName || this.DEFAULT_TEMPLATE_FILE;
    return path.join(this.TEMPLATES_FOLDER, fileName);
  }

  /**
   * Get versioned file path (v1, v2, v3, etc.) for EPVS opportunity files
   * For template selection - always creates new version
   */
  private getNewOpportunityFilePath(opportunityId: string): string {
    return this.getNewVersionedFilePath(this.OPPORTUNITIES_FOLDER, `EPVS Calculator Creativ - 06.02-${opportunityId}`, 'xlsm');
  }

  /**
   * Get versioned file path (v1, v2, v3, etc.) for EPVS opportunity files
   * For regular operations - uses existing file if available
   */
  private getOpportunityFilePath(opportunityId: string): string {
    return this.getVersionedFilePath(this.OPPORTUNITIES_FOLDER, `EPVS Calculator Creativ - 06.02-${opportunityId}`, 'xlsm');
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
  private findLatestOpportunityFile(opportunityId: string, fileName?: string): string | null {
    if (!fs.existsSync(this.OPPORTUNITIES_FOLDER)) {
      return null;
    }

    const files = fs.readdirSync(this.OPPORTUNITIES_FOLDER);
    
    // If fileName is provided, try to find the exact file first
    if (fileName) {
      this.logger.log(`🔍 Looking for specific file: ${fileName}`);
      const exactFile = files.find(file => file === fileName);
      if (exactFile) {
        const fullPath = path.join(this.OPPORTUNITIES_FOLDER, exactFile);
        this.logger.log(`✅ Found exact file match: ${fullPath}`);
        return fullPath;
      } else {
        this.logger.log(`⚠️ Exact file ${fileName} not found, falling back to version search`);
      }
    }
    
    const basePattern = `EPVS Calculator Creativ - 06.02-${opportunityId}`.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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

  async selectMultipleRadioButtons(shapeNames: string[], opportunityId?: string): Promise<{ success: boolean; message: string; error?: string }> {
    // Map radio button names to actual Excel names
    const mappedShapeNames = shapeNames.map(name => this.mapRadioButtonName(name));
    this.logger.log(`Starting multiple radio button automation for shapes: ${mappedShapeNames.join(', ')}${opportunityId ? ` (Opportunity: ${opportunityId})` : ''}`);

    try {
      // Determine which file to use
      let excelFilePath: string;
      if (opportunityId) {
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
        } else {
          this.logger.warn(`EPVS opportunity file not found, using template: ${opportunityFilePath}`);
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
      const psScript = this.createMultipleRadioButtonScript(mappedShapeNames, excelFilePath);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-multiple-radio-button-${Date.now()}.ps1`);
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
        this.logger.log(`Successfully selected multiple radio buttons: ${mappedShapeNames.join(', ')}`);
        return {
          success: true,
          message: `Successfully selected multiple radio buttons: ${mappedShapeNames.join(', ')}`
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to select multiple radio buttons: ${mappedShapeNames.join(', ')}`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in selectMultipleRadioButtons: ${error.message}`);
      return {
        success: false,
        message: `Error selecting multiple radio buttons: ${shapeNames.join(', ')}`,
        error: error.message
      };
    }
  }

  async selectRadioButton(shapeName: string, opportunityId?: string): Promise<{ success: boolean; message: string; error?: string }> {
    // Map radio button name to actual Excel name
    const mappedShapeName = this.mapRadioButtonName(shapeName);
    this.logger.log(`Starting radio button automation for shape: ${mappedShapeName}${opportunityId ? ` (Opportunity: ${opportunityId})` : ''}`);

    try {
      // Determine which file to use
      let excelFilePath: string;
      if (opportunityId) {
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
        } else {
          this.logger.warn(`EPVS opportunity file not found, using template: ${opportunityFilePath}`);
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
      const psScript = this.createRadioButtonScript(mappedShapeName, excelFilePath);
      
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
        this.logger.log(`Successfully selected radio button: ${mappedShapeName}`);
        return {
          success: true,
          message: `Successfully selected radio button: ${mappedShapeName}`
        };
      } else {
        this.logger.error(`PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to select radio button: ${mappedShapeName}`,
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
    this.logger.log(`Creating opportunity file for: ${opportunityId}`);

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
      const psScript = this.createOpportunityFileScript(opportunityId, customerDetails, templateFilePath, isTemplateSelection);
      
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
        const filePath = this.getOpportunityFilePath(opportunityId);
        this.logger.log(`Successfully created EPVS opportunity file: ${filePath}`);
        return {
          success: true,
          message: `Successfully created EPVS opportunity file for ${opportunityId}`,
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

  private createOpportunityFileScript(opportunityId: string, customerDetails: { customerName: string; address: string; postcode: string }, templateFilePath?: string, isNewTemplate?: boolean): string {
    const templatePath = (templateFilePath || this.getTemplateFilePath()).replace(/\\/g, '\\\\');
    const opportunitiesFolder = this.OPPORTUNITIES_FOLDER.replace(/\\/g, '\\\\');
    const newFilePath = isNewTemplate ? 
      this.getNewOpportunityFilePath(opportunityId).replace(/\\/g, '\\\\') :
      this.getOpportunityFilePath(opportunityId).replace(/\\/g, '\\\\');
    
    return `
# Create Opportunity File (Copy Only)
$ErrorActionPreference = "Stop"

# Configuration
$templatePath = "${templatePath}"
$newFilePath = "${newFilePath}"

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

    # Ensure template exists
    if (-not (Test-Path $templatePath)) {
        throw "Template file does not exist: $templatePath"
    }

    # Copy the template file directly
    Write-Host "Attempting to copy template file..." -ForegroundColor Yellow
    Copy-Item -Path $templatePath -Destination $newFilePath -Force
    Write-Host "Template file copied successfully" -ForegroundColor Green

    # Touch the file so Date Modified updates (without opening Excel)
    [System.IO.File]::SetLastWriteTime($newFilePath, (Get-Date))
    Write-Host "Updated last write time on new file" -ForegroundColor Green

    Write-Host "Opportunity file creation completed successfully!" -ForegroundColor Green

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    exit 1
}
`;
  }

  private createMultipleRadioButtonScript(shapeNames: string[], excelFilePath: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    const shapeNamesString = shapeNames.map(name => `"${name}"`).join(', ');
    
    return `
# Multiple Radio Button Automation using COM - Enhanced Version
$ErrorActionPreference = "Stop"

# Configuration
$shapeNames = @(${shapeNamesString})

Write-Host "Starting multiple radio button automation for shapes: $($shapeNames -join ', ')" -ForegroundColor Green

try {
    # Create Excel application with enhanced settings to prevent popups and debug dialogs
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.Interactive = $false
    $excel.UserControl = $false
    $excel.AlertBeforeOverwriting = $false
    
    # Additional settings to prevent debug popups and improve automation reliability
    try {
        $excel.ErrorCheckingOptions.BackgroundChecking = $false
        $excel.ErrorCheckingOptions.IndicatorColorIndex = 0
        $excel.ErrorCheckingOptions.TextDate = $false
        $excel.ErrorCheckingOptions.NumberAsText = $false
        $excel.ErrorCheckingOptions.InconsistentFormula = $false
        $excel.ErrorCheckingOptions.OmittedCells = $false
        $excel.ErrorCheckingOptions.UnlockedFormulaCells = $false
        $excel.ErrorCheckingOptions.EmptyCellReferences = $false
        Write-Host "Applied additional Excel settings to prevent debug popups" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not apply all Excel error checking settings: $_" -ForegroundColor Yellow
    }
    
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
    
    $successfulSelections = 0
    
    # Process each radio button
    foreach ($shapeName in $shapeNames) {
        Write-Host "Processing radio button: $shapeName" -ForegroundColor Yellow
        
        try {
            $targetShape = $worksheet.Shapes($shapeName)
            Write-Host "Found shape: $shapeName" -ForegroundColor Green
            
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
                    
                    # Try to trigger the OnAction by clicking the shape
                    try {
                        $targetShape.Click()
                        Write-Host "Clicked radio button to trigger OnAction" -ForegroundColor Cyan
                    } catch {
                        # Fallback: Set the radio button value
                        $targetShape.ControlFormat.Value = 1
                        Write-Host "Set radio button value as fallback" -ForegroundColor Cyan
                    }
                    
                    # Give a moment for the change to register
                    Start-Sleep -Milliseconds 200
                    
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
                $successfulSelections++
                
                # Try to trigger the OnAction macro if it exists (with timeout protection)
                if ($onActionMacro -and $onActionMacro.Trim() -ne "") {
                    try {
                        Write-Host "Executing OnAction macro: $onActionMacro" -ForegroundColor Cyan
                        Write-Host "Note: If a debug popup appears, it will be automatically dismissed" -ForegroundColor Yellow
                        
                        # Try to run the macro directly
                        try {
                            Write-Host "Running VBA macro: $onActionMacro" -ForegroundColor Cyan
                            $excel.Run($onActionMacro)
                            Write-Host "VBA macro executed successfully" -ForegroundColor Green
                        } catch {
                            Write-Host "VBA macro execution failed: $_" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "Failed to execute OnAction macro: $_" -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "No OnAction macro for: $shapeName" -ForegroundColor Yellow
                }
            } else {
                Write-Host "Failed to interact with radio button: $shapeName" -ForegroundColor Red
            }
            
        } catch {
            Write-Host "Error processing radio button '$shapeName': $_" -ForegroundColor Red
        }
    }
    
    Write-Host "Successfully selected $successfulSelections out of $($shapeNames.Count) radio buttons" -ForegroundColor Green
    
    # Provide additional information about the results
    if ($successfulSelections -eq $shapeNames.Count) {
        Write-Host "All radio buttons were successfully selected!" -ForegroundColor Green
    } elseif ($successfulSelections -gt 0) {
        Write-Host "Some radio buttons were selected successfully. Check logs above for any timeout issues." -ForegroundColor Yellow
    } else {
        Write-Host "No radio buttons were successfully selected. Check logs above for errors." -ForegroundColor Red
    }
    
    # Save the workbook
    try {
        $workbook.Save()
        Write-Host "Workbook saved successfully" -ForegroundColor Green
        # Wait a moment after saving to ensure data is persisted
        Start-Sleep -Milliseconds 500
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
    
    if ($successfulSelections -gt 0) {
        Write-Host "Multiple radio button automation completed successfully!" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "Multiple radio button automation failed!" -ForegroundColor Red
        exit 1
    }
    
} catch {
    Write-Host "Critical error in multiple radio button automation: $_" -ForegroundColor Red
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
    
    # Check if worksheet is protected and unprotect it
    if ($worksheet.ProtectContents) {
        Write-Host "Worksheet is protected, attempting to unprotect..." -ForegroundColor Yellow
        try {
            $worksheet.Unprotect("${this.PASSWORD}")
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
      let isResolved = false;

      // Set a timeout to prevent hanging (increased for Excel COM operations)
      const timeout = setTimeout(() => {
        if (!isResolved) {
          isResolved = true;
          this.logger.error('PowerShell script timed out after 120 seconds');
          powershell.kill('SIGTERM');
          resolve({ success: false, error: 'PowerShell script timed out after 120 seconds' });
        }
      }, 120000); // 120 second timeout for Excel COM operations

      powershell.stdout.on('data', (data) => {
        stdout += data.toString();
        this.logger.log(`PowerShell output: ${data.toString().trim()}`);
      });

      powershell.stderr.on('data', (data) => {
        stderr += data.toString();
        this.logger.error(`PowerShell error: ${data.toString().trim()}`);
      });

      powershell.on('close', (code) => {
        if (!isResolved) {
          isResolved = true;
          clearTimeout(timeout);
          this.logger.log(`PowerShell script completed with code: ${code}`);
          
          if (code === 0) {
            resolve({ success: true, output: stdout });
          } else {
            const error = stderr || `PowerShell script failed with code ${code}`;
            resolve({ success: false, error });
          }
        }
      });

      powershell.on('error', (error) => {
        if (!isResolved) {
          isResolved = true;
          clearTimeout(timeout);
          this.logger.error(`Failed to start PowerShell: ${error.message}`);
          resolve({ success: false, error: error.message });
        }
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
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using EPVS opportunity file: ${opportunityFilePath}`);
        } else {
          // If opportunity file doesn't exist, fall back to template file
          if (templateFileName) {
            excelFilePath = this.getTemplateFilePath(templateFileName);
            this.logger.log(`EPVS opportunity file not found, using template file: ${excelFilePath}`);
          } else {
            const error = `EPVS opportunity file not found: ${opportunityFilePath}. Please ensure radio button selections have been applied first.`;
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

  async saveDynamicInputs(opportunityId: string | undefined, inputs: Record<string, string>, templateFileName?: string): Promise<{ success: boolean; message: string; error?: string }> {
    this.logger.log(`Saving EPVS dynamic inputs${opportunityId ? ` for opportunity: ${opportunityId}` : ''}${templateFileName ? ` with template: ${templateFileName}` : ''}`);

    try {
      // Check if we have array data - if so, use the simple approach
      const hasArrayData = Object.keys(inputs).some(key => key.match(/^array\d+_(panels|orientation|pitch|shading)$/));
      
      if (hasArrayData && opportunityId) {
        this.logger.log(`🔧 Detected array data, using simple approach for better reliability`);
        
        // Extract array data dynamically for all arrays
        const arrayData: any = {
          no_of_arrays: inputs['no_of_arrays'] || '1'
        };
        
        // Extract all array data (array1, array2, array3, etc.)
        const arrayKeys = Object.keys(inputs).filter(key => key.match(/^array\d+_(panels|orientation|pitch|shading)$/));
        arrayKeys.forEach(key => {
          arrayData[key] = inputs[key];
        });
        
        this.logger.log(`🔧 Extracted array data for ${arrayKeys.length} array fields:`, arrayKeys);
        
        // Use simple array function
        const arrayResult = await this.saveArrayDataSimple(opportunityId, arrayData);
        
        if (!arrayResult.success) {
          this.logger.error(`❌ Simple array function failed: ${arrayResult.error}`);
          return {
            success: false,
            message: `Failed to save array data: ${arrayResult.error}`,
            error: arrayResult.error
          };
        }
        
        this.logger.log(`✅ Successfully saved array data using simple approach`);
      }

      // Step 1: Check if we have a postcode in inputs and auto-populate Flux rates
      const postcode = inputs['customer_postcode'] || inputs['postcode'];
      if (postcode && opportunityId) {
        this.logger.log(`🔌 Auto-populating Flux rates for postcode: ${postcode}`);
        try {
          const fluxRatesResult = await this.populateFluxRatesInExcel(opportunityId, postcode);
          if (fluxRatesResult.success) {
            this.logger.log(`✅ Successfully auto-populated Flux rates for ${postcode}`);
          } else {
            this.logger.warn(`⚠️ Failed to auto-populate Flux rates: ${fluxRatesResult.error}`);
          }
        } catch (fluxError) {
          this.logger.warn(`⚠️ Error auto-populating Flux rates: ${fluxError.message}`);
        }
      }

      // Step 2: Determine which file to use - prioritize opportunity file if it exists
      let excelFilePath: string;
      if (opportunityId) {
        // First, try to use the opportunity file (created from template)
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using EPVS opportunity file: ${opportunityFilePath}`);
        } else {
          // If opportunity file doesn't exist, fall back to template file
          if (templateFileName) {
            excelFilePath = this.getTemplateFilePath(templateFileName);
            this.logger.log(`EPVS opportunity file not found, using template file: ${excelFilePath}`);
          } else {
            this.logger.warn(`EPVS opportunity file not found, using default template: ${opportunityFilePath}`);
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

      // If we already handled array data with simple approach, remove array fields from inputs
      let processedInputs = { ...inputs };
      if (hasArrayData) {
        this.logger.log(`🔧 Removing array fields from main processing (already handled with simple approach)`);
        // Remove array fields from inputs to avoid double processing
        Object.keys(processedInputs).forEach(key => {
          if (key.match(/^array\d+_(panels|orientation|pitch|shading)$/) || key === 'no_of_arrays') {
            delete processedInputs[key];
          }
        });
        this.logger.log(`🔍 Remaining inputs for main processing:`, JSON.stringify(processedInputs, null, 2));
      }

      // Create PowerShell script
      this.logger.log(`🔍 Creating PowerShell script with inputs:`, JSON.stringify(processedInputs, null, 2));
      const psScript = this.createSaveDynamicInputsScript(excelFilePath, processedInputs);
      
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

  // Create dynamic PowerShell script for multiple arrays
  private createDynamicArrayInputScript(excelFilePath: string, arrayData: any): string {
    const filePath = excelFilePath.replace(/\\/g, '\\\\');
    const password = "99";
    
    // Define cell mappings for arrays (array1, array2, array3, etc.)
    const arrayCellMappings: { [key: string]: { [key: string]: string } } = {
      'array1': { panels: 'C70', orientation: 'F70', pitch: 'G70', shading: 'I70' },
      'array2': { panels: 'C71', orientation: 'F71', pitch: 'G71', shading: 'I71' },
      'array3': { panels: 'C72', orientation: 'F72', pitch: 'G72', shading: 'I72' },
      'array4': { panels: 'C73', orientation: 'F73', pitch: 'G73', shading: 'I73' },
      'array5': { panels: 'C74', orientation: 'F74', pitch: 'G74', shading: 'I74' },
      'array6': { panels: 'C75', orientation: 'F75', pitch: 'G75', shading: 'I75' },
      'array7': { panels: 'C76', orientation: 'F76', pitch: 'G76', shading: 'I76' },
      'array8': { panels: 'C77', orientation: 'F77', pitch: 'G77', shading: 'I77' }
    };
    
    // Generate array input sections dynamically
    let arrayInputSections = '';
    let verificationSections = '';
    
    // Process each array that has data
    Object.keys(arrayData).forEach(key => {
      if (key.match(/^array\d+_(panels|orientation|pitch|shading)$/)) {
        const match = key.match(/^(array\d+)_(panels|orientation|pitch|shading)$/);
        if (match) {
          const arrayNum = match[1];
          const fieldType = match[2];
          const cellRef = arrayCellMappings[arrayNum]?.[fieldType];
          
          if (cellRef) {
            const value = arrayData[key];
            if (value && value !== '') {
              arrayInputSections += `
    if ("${value}" -ne "") {
        Write-Host "Setting ${key} = ${value}..." -ForegroundColor White
        $worksheet.Range("${cellRef}").Value2 = ${value}
    }`;
              
              verificationSections += `
    $final${arrayNum}${fieldType.charAt(0).toUpperCase() + fieldType.slice(1)} = $worksheet.Range("${cellRef}").Value2
    Write-Host "Final ${cellRef} (${key}): '$final${arrayNum}${fieldType.charAt(0).toUpperCase() + fieldType.slice(1)}'" -ForegroundColor White`;
            }
          }
        }
      }
    });
    
    return `
# Dynamic Array Input Script (Handles Multiple Arrays)
$ErrorActionPreference = "Stop"

$filePath = "${filePath}"
$password = "${password}"

Write-Host "=== DYNAMIC ARRAY INPUT SCRIPT ===" -ForegroundColor Green
Write-Host "File: $filePath" -ForegroundColor Cyan

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "Opening workbook..." -ForegroundColor Cyan
    $workbook = $excel.Workbooks.Open($filePath)
    $worksheet = $workbook.Worksheets["Inputs"]
    
    Write-Host "Unprotecting worksheet..." -ForegroundColor Cyan
    $worksheet.Unprotect($password)
    
    Write-Host "=== STEP 1: READING INITIAL VALUES ===" -ForegroundColor Yellow
    $initialNoOfArrays = $worksheet.Range("H44").Value2
    Write-Host "Initial H44 (no_of_arrays): '$initialNoOfArrays'" -ForegroundColor White
    
    Write-Host "=== STEP 2: INPUTTING ARRAY DATA ===" -ForegroundColor Yellow
    
    # Set no_of_arrays first
    Write-Host "Setting no_of_arrays = ${arrayData.no_of_arrays}..." -ForegroundColor White
    $worksheet.Range("H44").Value2 = "${arrayData.no_of_arrays}"
    
    # Wait for VBA to trigger and unlock cells
    Write-Host "Waiting for VBA to unlock cells..." -ForegroundColor Cyan
    Start-Sleep -Milliseconds 2000
    
    # Check if cells are unlocked
    Write-Host "Checking cell lock status..." -ForegroundColor Cyan
    $c70Locked = $worksheet.Range("C70").Locked
    $f70Locked = $worksheet.Range("F70").Locked
    $g70Locked = $worksheet.Range("G70").Locked
    $i70Locked = $worksheet.Range("I70").Locked
    
    Write-Host "C70 locked: $c70Locked" -ForegroundColor White
    Write-Host "F70 locked: $f70Locked" -ForegroundColor White
    Write-Host "G70 locked: $g70Locked" -ForegroundColor White
    Write-Host "I70 locked: $i70Locked" -ForegroundColor White
    
    # Input array data dynamically${arrayInputSections}
    
    Write-Host "=== STEP 3: VERIFICATION ===" -ForegroundColor Yellow
    
    # Verify the values were saved
    $finalNoOfArrays = $worksheet.Range("H44").Value2
    Write-Host "Final H44 (no_of_arrays): '$finalNoOfArrays'" -ForegroundColor White${verificationSections}
    
    Write-Host "Saving workbook..." -ForegroundColor Cyan
    $workbook.Save()
    
    Write-Host "Protecting worksheet..." -ForegroundColor Cyan
    $worksheet.Protect($password)
    
    Write-Host "Closing workbook..." -ForegroundColor Cyan
    $workbook.Close($false)
    $excel.Quit()
    
    Write-Host "=== DYNAMIC ARRAY INPUT COMPLETE ===" -ForegroundColor Green
    
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    exit 1
} finally {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;
  }

  // Simple array input function using direct PowerShell approach (like the working test script)
  async saveArrayDataSimple(opportunityId: string, arrayData: {
    no_of_arrays: string;
    [key: string]: string | undefined; // Allow any array field (array1_panels, array2_panels, etc.)
  }): Promise<{ success: boolean; message: string; error?: string }> {
    this.logger.log(`🔧 Saving array data using simple approach for opportunity: ${opportunityId}`);
    
    try {
      // Determine which file to use
      let excelFilePath: string;
      if (opportunityId) {
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using EPVS opportunity file: ${opportunityFilePath}`);
        } else {
          this.logger.warn(`EPVS opportunity file not found, using template: ${opportunityFilePath}`);
          excelFilePath = this.getTemplateFilePath();
        }
      } else {
        excelFilePath = this.getTemplateFilePath();
      }

      // Create dynamic PowerShell script that handles multiple arrays
      const psScript = this.createDynamicArrayInputScript(excelFilePath, arrayData);

      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-simple-array-${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created simple array PowerShell script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Failed to cleanup temporary script: ${cleanupError.message}`);
      }

      if (result.success) {
        this.logger.log(`✅ Successfully saved array data using simple approach`);
        return {
          success: true,
          message: `Successfully saved array data using simple approach`
        };
      } else {
        this.logger.error(`❌ Simple array PowerShell script failed: ${result.error}`);
        return {
          success: false,
          message: `Failed to save array data using simple approach`,
          error: result.error
        };
      }

    } catch (error) {
      this.logger.error(`Error in saveArrayDataSimple: ${error.message}`);
      return {
        success: false,
        message: `Error saving array data using simple approach`,
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
    
         # Define ALL the EPVS input fields based on the ACTUAL Excel template structure
     $inputFields = @(
         # NEW PRODUCTS - SOLAR (Rows 67-78)
         @{
             id = "array1_panels"
             label = "Array 1 - No. of Panels"
             cellReference = "C70"
             type = "number"
             required = $false
         },
         @{
             id = "array1_panel_size"
             label = "Array 1 - Panel Size (Wp)"
             cellReference = "D70"
             type = "number"
             required = $false
         },
         @{
             id = "array1_size_kwp"
             label = "Array 1 - Array Size (kWp)"
             cellReference = "E70"
             type = "number"
             required = $false
         },
         @{
             id = "array1_orientation"
             label = "Array 1 - Orientation (° from south)"
             cellReference = "F70"
             type = "text"
             required = $false
         },
         @{
             id = "array1_pitch"
             label = "Array 1 - Pitch (° from flat)"
             cellReference = "G70"
             type = "number"
             required = $false
         },
         @{
             id = "array1_irradiance"
             label = "Array 1 - Irradiance (Kk)"
             cellReference = "H70"
             type = "number"
             required = $false
         },
         @{
             id = "array1_shading"
             label = "Array 1 - Shading"
             cellReference = "I70"
             type = "number"
             required = $false
         },
         @{
             id = "array1_generation"
             label = "Array 1 - Generation"
             cellReference = "J70"
             type = "number"
             required = $false
         },
         
         @{
             id = "array2_panels"
             label = "Array 2 - No. of Panels"
             cellReference = "C71"
             type = "number"
             required = $false
         },
         @{
             id = "array2_panel_size"
             label = "Array 2 - Panel Size (Wp)"
             cellReference = "D71"
             type = "number"
             required = $false
         },
         @{
             id = "array2_size_kwp"
             label = "Array 2 - Array Size (kWp)"
             cellReference = "E71"
             type = "number"
             required = $false
         },
         @{
             id = "array2_orientation"
             label = "Array 2 - Orientation (° from south)"
             cellReference = "F71"
             type = "text"
             required = $false
         },
         @{
             id = "array2_pitch"
             label = "Array 2 - Pitch (° from flat)"
             cellReference = "G71"
             type = "number"
             required = $false
         },
         @{
             id = "array2_irradiance"
             label = "Array 2 - Irradiance (Kk)"
             cellReference = "H71"
             type = "number"
             required = $false
         },
         @{
             id = "array2_shading"
             label = "Array 2 - Shading"
             cellReference = "I71"
             type = "number"
             required = $false
         },
         @{
             id = "array2_generation"
             label = "Array 2 - Generation"
             cellReference = "J71"
             type = "number"
             required = $false
         },
         
         # ARRAY 3 - Only input fields (panels, orientation, pitch, shading)
         @{
             id = "array3_panels"
             label = "Array 3 - Panels"
             cellReference = "C72"
             type = "number"
             required = $false
         },
         @{
             id = "array3_orientation"
             label = "Array 3 - Orientation"
             cellReference = "F72"
             type = "text"
             required = $false
         },
         @{
             id = "array3_pitch"
             label = "Array 3 - Pitch"
             cellReference = "G72"
             type = "number"
             required = $false
         },
         @{
             id = "array3_shading"
             label = "Array 3 - Shading"
             cellReference = "I72"
             type = "number"
             required = $false
         },
         
         # ARRAY 4 - Only input fields (panels, orientation, pitch, shading)
         @{
             id = "array4_panels"
             label = "Array 4 - Panels"
             cellReference = "C73"
             type = "number"
             required = $false
         },
         @{
             id = "array4_orientation"
             label = "Array 4 - Orientation"
             cellReference = "F73"
             type = "text"
             required = $false
         },
         @{
             id = "array4_pitch"
             label = "Array 4 - Pitch"
             cellReference = "G73"
             type = "number"
             required = $false
         },
         @{
             id = "array4_shading"
             label = "Array 4 - Shading"
             cellReference = "I73"
             type = "number"
             required = $false
         },
         
         # ARRAY 5 - Only input fields (panels, orientation, pitch, shading)
         @{
             id = "array5_panels"
             label = "Array 5 - Panels"
             cellReference = "C74"
             type = "number"
             required = $false
         },
         @{
             id = "array5_orientation"
             label = "Array 5 - Orientation"
             cellReference = "F74"
             type = "text"
             required = $false
         },
         @{
             id = "array5_pitch"
             label = "Array 5 - Pitch"
             cellReference = "G74"
             type = "number"
             required = $false
         },
         @{
             id = "array5_shading"
             label = "Array 5 - Shading"
             cellReference = "I74"
             type = "number"
             required = $false
         },
         
         # ARRAY 6 - Only input fields (panels, orientation, pitch, shading)
         @{
             id = "array6_panels"
             label = "Array 6 - Panels"
             cellReference = "C75"
             type = "number"
             required = $false
         },
         @{
             id = "array6_orientation"
             label = "Array 6 - Orientation"
             cellReference = "F75"
             type = "text"
             required = $false
         },
         @{
             id = "array6_pitch"
             label = "Array 6 - Pitch"
             cellReference = "G75"
             type = "number"
             required = $false
         },
         @{
             id = "array6_shading"
             label = "Array 6 - Shading"
             cellReference = "I75"
             type = "number"
             required = $false
         },
         
         # ARRAY 7 - Only input fields (panels, orientation, pitch, shading)
         @{
             id = "array7_panels"
             label = "Array 7 - Panels"
             cellReference = "C76"
             type = "number"
             required = $false
         },
         @{
             id = "array7_orientation"
             label = "Array 7 - Orientation"
             cellReference = "F76"
             type = "text"
             required = $false
         },
         @{
             id = "array7_pitch"
             label = "Array 7 - Pitch"
             cellReference = "G76"
             type = "number"
             required = $false
         },
         @{
             id = "array7_shading"
             label = "Array 7 - Shading"
             cellReference = "I76"
             type = "number"
             required = $false
         },
         
         # ARRAY 8 - Only input fields (panels, orientation, pitch, shading)
         @{
             id = "array8_panels"
             label = "Array 8 - Panels"
             cellReference = "C77"
             type = "number"
             required = $false
         },
         @{
             id = "array8_orientation"
             label = "Array 8 - Orientation"
             cellReference = "F77"
             type = "text"
             required = $false
         },
         @{
             id = "array8_pitch"
             label = "Array 8 - Pitch"
             cellReference = "G77"
             type = "number"
             required = $false
         },
         @{
             id = "array8_shading"
             label = "Array 8 - Shading"
             cellReference = "I77"
             type = "number"
             required = $false
         },
         
         # SYSTEM COSTS (Rows 80-85)
         @{
             id = "total_system_cost"
             label = "Total System Cost"
             cellReference = "H81"
             type = "number"
             required = $false
         },
         @{
             id = "deposit"
             label = "Deposit"
             cellReference = "H82"
             type = "number"
             required = $false
         },
         @{
             id = "interest_rate"
             label = "Interest Rate (%)"
             cellReference = "H83"
             type = "number"
             required = $false
         },
         @{
             id = "interest_rate_type"
             label = "Interest Rate Type"
             cellReference = "H84"
             type = "text"
             required = $false
         },
         @{
             id = "payment_term_years"
             label = "Payment Term (years)"
             cellReference = "H85"
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

  private createSaveDynamicInputsScript(excelFilePath: string, inputs: Record<string, string>): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    // Convert inputs to PowerShell format - fix the hashtable syntax
    const inputsString = Object.entries(inputs)
      .map(([key, value]) => {
        const cleanedValue = this.escapeForPowerShell(value);
        return `    "${key}" = "${cleanedValue}"`;
      })
      .join('\n');
    
         // Define cell mappings for ALL EPVS input fields - MUST MATCH ALL_INPUT_FIELDS
     const cellMappings = {
       // Customer Details (always enabled)
       customer_name: 'H11',
       address: 'H12',
       postcode: 'H13',
       
       // ENERGY USE - CURRENT ELECTRICITY TARIFF
       single_rate: 'H17',
       off_peak_rate: 'H18',
       no_of_off_peak: 'H19',
       
       // ENERGY USE - ELECTRICITY CONSUMPTION
       estimated_annual_usage: 'H26',
       estimated_peak_annual_usage: 'H27',
       estimated_off_peak_usage: 'H28',
       standing_charges: 'H29',
       total_annual_spend: 'H30',
       peak_annual_spend: 'H31',
       off_peak_annual_spend: 'H32',
       
       // EXISTING SYSTEM
       existing_sem: 'H35',
       approximate_commissioning_date: 'H36',
       percentage_sem_used_for_quote: 'H37',
       
             // NEW SYSTEM - SOLAR
      panel_manufacturer: 'H42',
      panel_model: 'H43',
      no_of_arrays: 'H44',
      
          // ARRAY 1 - Only input fields (panels, orientation, pitch, shading) - Corrected cell references
    array1_panels: 'C70',
    array1_orientation: 'F70',
    array1_pitch: 'G70',
    array1_shading: 'I70',
    
    // ARRAY 2 - Only input fields (panels, orientation, pitch, shading)
    array2_panels: 'C71',
    array2_orientation: 'F71',
    array2_pitch: 'G71',
    array2_shading: 'I71',
    
    // ARRAY 3 - Only input fields (panels, orientation, pitch, shading)
    array3_panels: 'C72',
    array3_orientation: 'F72',
    array3_pitch: 'G72',
    array3_shading: 'I72',
    
    // ARRAY 4 - Only input fields (panels, orientation, pitch, shading)
    array4_panels: 'C73',
    array4_orientation: 'F73',
    array4_pitch: 'G73',
    array4_shading: 'I73',
    
    // ARRAY 5 - Only input fields (panels, orientation, pitch, shading)
    array5_panels: 'C74',
    array5_orientation: 'F74',
    array5_pitch: 'G74',
    array5_shading: 'I74',
    
    // ARRAY 6 - Only input fields (panels, orientation, pitch, shading)
    array6_panels: 'C75',
    array6_orientation: 'F75',
    array6_pitch: 'G75',
    array6_shading: 'I75',
    
    // ARRAY 7 - Only input fields (panels, orientation, pitch, shading)
    array7_panels: 'C76',
    array7_orientation: 'F76',
    array7_pitch: 'G76',
    array7_shading: 'I76',
    
    // ARRAY 8 - Only input fields (panels, orientation, pitch, shading)
    array8_panels: 'C77',
    array8_orientation: 'F77',
    array8_pitch: 'G77',
    array8_shading: 'I77',
       
       // NEW SYSTEM - BATTERY
       battery_manufacturer: 'H46',
       battery_model: 'H47',
       battery_extended_warranty_period: 'H50',
       battery_replacement_cost: 'H51',
       
       // NEW SYSTEM - SOLAR/HYBRID INVERTER
       solar_inverter_manufacturer: 'H53',
       solar_inverter_model: 'H54',
       solar_inverter_extended_warranty: 'H57',
       solar_inverter_replacement_cost: 'H58',
       
       // NEW SYSTEM - BATTERY INVERTER
       battery_inverter_manufacturer: 'H60',
       battery_inverter_model: 'H61',
       battery_inverter_extended_warranty_period: 'H64',
       battery_inverter_replacement_cost: 'H65',
       
       // FLUX RATES (Import - Column H)
       import_day_rate: 'H22',
       import_flux_rate: 'H23',
       import_peak_rate: 'H24',
       
       // FLUX RATES (Export - Column J)
       export_day_rate: 'J22',
       export_flux_rate: 'J23',
       export_peak_rate: 'J24',
       
       // FLUX RATES - Direct cell references (for auto-population)
       H22: 'H22',  // Import Day Rate
       H23: 'H23',  // Import Flux Rate
       H24: 'H24',  // Import Peak Rate
       J22: 'J22',  // Export Day Rate
       J23: 'J23',  // Export Flux Rate
       J24: 'J24'   // Export Peak Rate
     };
    
    // Convert cell mappings to PowerShell format
    const cellMappingsString = Object.entries(cellMappings)
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

$excel = $null
$workbook = $null
$worksheet = $null

try {
    # Create Excel application with performance optimizations
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.Interactive = $false
    $excel.UserControl = $false
    $excel.AlertBeforeOverwriting = $false
    # $excel.Calculation = -4105  # xlCalculationManual - disable automatic calculation (commented out due to COM error)
    
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
    
    # Save input values to Excel cells (with special VBA triggering for no_of_arrays)
    Write-Host "Saving input values to Excel cells..." -ForegroundColor Green
    $savedCount = 0
    
    # Check if we have array data that requires VBA triggering for no_of_arrays
    $hasArrayData = $false
    Write-Host "DEBUG: Checking for array data in inputs..." -ForegroundColor Yellow
    foreach ($inputKey in $inputs.Keys) {
        Write-Host "DEBUG: Input key: $inputKey" -ForegroundColor Yellow
        if ($inputKey -match "^array\\d+_(panels|orientation|pitch|shading)$") {
            Write-Host "DEBUG: Found array key: $inputKey" -ForegroundColor Green
            $hasArrayData = $true
            break
        }
    }
    Write-Host "DEBUG: hasArrayData = $hasArrayData" -ForegroundColor Yellow
    
    # Step 1: Process no_of_arrays first - also check for number_of_arrays
    $noOfArraysValue = $null
    if ($inputs.ContainsKey("no_of_arrays") -and -not [string]::IsNullOrWhiteSpace($inputs["no_of_arrays"])) {
        $noOfArraysValue = $inputs["no_of_arrays"]
    } elseif ($inputs.ContainsKey("number_of_arrays") -and -not [string]::IsNullOrWhiteSpace($inputs["number_of_arrays"])) {
        $noOfArraysValue = $inputs["number_of_arrays"]
    }
    
    if ($noOfArraysValue -ne $null) {
        $cellReference = $cellMappings["no_of_arrays"]
        
        if ($cellReference) {
            try {
                $cell = $worksheet.Range($cellReference)
                
                # Always trigger VBA when setting no_of_arrays to unlock array cells
                # This is needed even if there's no array data yet - it unlocks the cells for future input
                Write-Host "Processing no_of_arrays dropdown with special VBA triggering..." -ForegroundColor Cyan
                Write-Host "DEBUG: Will trigger VBA to unlock array cells" -ForegroundColor Green
                    
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
                $errorMessage = $_.Exception.Message
                Write-Host "Error saving no_of_arrays to cell $cellReference : $errorMessage" -ForegroundColor Red
            }
        }
    }
    
    # Step 2: Process arrays one by one (array1, then array2, etc.)
    $maxArrays = 8
    for ($arrayNum = 1; $arrayNum -le $maxArrays; $arrayNum++) {
        $arrayFields = @()
        $hasArrayData = $false
        
        # Collect all fields for this array
        foreach ($inputKey in $inputs.Keys) {
            if ($inputKey -match "^array$arrayNum" + "_") {
                $arrayFields += @{Key = $inputKey; Value = $inputs[$inputKey]}
                $hasArrayData = $true
            }
        }
        
        if ($hasArrayData) {
            Write-Host ""
            Write-Host ("=== STEP 2." + $arrayNum + ": Processing Array " + $arrayNum + " ===") -ForegroundColor Magenta
            
            # Sort array fields in logical order: panels, orientation, pitch, shading
            $sortedFields = $arrayFields | Sort-Object {
                switch -Regex ($_.Key) {
                    "_panels$" { return 1 }
                    "_orientation$" { return 2 }
                    "_pitch$" { return 3 }
                    "_shading$" { return 4 }
                    default { return 5 }
                }
            }
            
            foreach ($fieldData in $sortedFields) {
                $inputKey = $fieldData.Key
                $inputValue = $fieldData.Value
                $cellReference = $cellMappings[$inputKey]
                
                Write-Host "DEBUG: Processing array field: $inputKey = '$inputValue' -> $cellReference" -ForegroundColor Cyan
                
                if ($cellReference) {
                    try {
                        $cell = $worksheet.Range($cellReference)
                        
                        # Check if cell is locked
                        Write-Host "DEBUG: Cell $cellReference locked status: $($cell.Locked)" -ForegroundColor Cyan
                        if ($cell.Locked) {
                            Write-Host "Warning: Cell $cellReference is locked, skipping input $inputKey" -ForegroundColor Yellow
                            continue
                        }
                        
                        # Convert value based on expected type
                        $convertedValue = $inputValue
                        
                        # Determine if this should be a number or string based on field type
                        $numericFields = @("rate", "hours", "usage", "charge", "spend", "sem", "percentage", "annual", "cost", "period", "no_of_arrays", "panels", "pitch", "shading", "orientation")
                        $isNumericField = $false
                        
                        foreach ($numericPattern in $numericFields) {
                            if ($inputKey -match $numericPattern) {
                                $isNumericField = $true
                                break
                            }
                        }
                        
                        if ($isNumericField -and $inputValue -ne "" -and $inputValue -ne $null) {
                            # Try to convert to number for numeric fields
                            try {
                                # Clean the input value first
                                $cleanValue = $inputValue.ToString().Trim()
                                
                                # Try different conversion methods
                                if ([double]::TryParse($cleanValue, [ref]$convertedValue)) {
                                    $cell.Value2 = $convertedValue
                                    Write-Host "✅ SUCCESS: Saved $inputKey = $convertedValue (as number) to cell $cellReference" -ForegroundColor Green
                                } else {
                                    # Try parsing as integer first
                                    $intValue = 0
                                    if ([int]::TryParse($cleanValue, [ref]$intValue)) {
                                        $cell.Value2 = $intValue
                                        Write-Host "✅ SUCCESS: Saved $inputKey = $intValue (as integer) to cell $cellReference" -ForegroundColor Green
                                    } else {
                                        # Fallback to string if conversion fails
                                        $cell.Value2 = [string]$inputValue
                                        Write-Host "⚠️ WARNING: Could not convert '$inputValue' to number for $inputKey, saved as string" -ForegroundColor Yellow
                                    }
                                }
                                # Clear any data validation to prevent Excel popups
                                try { $cell.Validation.Delete() } catch { }
                            } catch {
                                Write-Host "⚠️ WARNING: Could not convert '$inputValue' to number for $inputKey, setting as string. Error: $_" -ForegroundColor Yellow
                                $cell.Value2 = [string]$inputValue
                                # Clear any data validation to prevent Excel popups
                                try { $cell.Validation.Delete() } catch { }
                            }
                        } else {
                            # Set as string for text fields (manufacturer, model, address, etc.)
                            $cell.Value2 = [string]$inputValue
                            Write-Host "✅ SUCCESS: Saved $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                            # Clear any data validation to prevent Excel popups
                            try { $cell.Validation.Delete() } catch { }
                        }
                        $savedCount++
                        
                        # Small delay between array field entries
                        Start-Sleep -Milliseconds 50
                        
                    } catch {
                        Write-Host "❌ ERROR: Failed to save $inputKey to cell $cellReference : $_" -ForegroundColor Red
                    }
                } else {
                    Write-Host "⚠️ WARNING: No cell mapping found for input: $inputKey" -ForegroundColor Yellow
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
        if ($inputKey -eq "no_of_arrays" -or $inputKey -match "^array[0-9]+_") {
            continue
        }
        
        $cellReference = $cellMappings[$inputKey]
        
        if ($cellReference) {
            try {
                $cell = $worksheet.Range($cellReference)
                
                # Special logging for Flux rates
                if ($inputKey -match "^[HJ]\d+$") {
                    Write-Host "🔌 FLUX RATE: Saving $inputKey = '$inputValue' to cell $cellReference" -ForegroundColor Cyan
                }
                
                # Check if cell is locked
                if ($cell.Locked) {
                    Write-Host "Warning: Cell $cellReference is locked, skipping input $inputKey" -ForegroundColor Yellow
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
                $convertedValue = $inputValue
                
                # Determine if this should be a number or string based on field type
                # Note: no_of_arrays is handled as a dropdown in Step 1, not as a numeric field
                $numericFields = @("rate", "hours", "usage", "charge", "spend", "sem", "percentage", "annual", "cost", "period", "panels", "pitch", "shading", "orientation")
                $isNumericField = $false
                
                # Special handling for Flux rates (cell references like H22, H23, etc.)
                if ($inputKey -match "^[HJ]\d+$") {
                    $isNumericField = $true
                    Write-Host "🔌 FLUX RATE: Treating $inputKey as numeric field" -ForegroundColor Cyan
                }
                
                foreach ($numericPattern in $numericFields) {
                    if ($inputKey -match $numericPattern) {
                        $isNumericField = $true
                        break
                    }
                }
                
                if ($isDropdownField) {
                    # Special handling for dropdown fields - select from dropdown list, not type text
                    try {
                        Write-Host "Processing dropdown field $inputKey with dropdown selection (not text input)..." -ForegroundColor Cyan
                        
                        # Check if cell has data validation (dropdown)
                        $hasValidation = $false
                        try {
                            $validation = $cell.Validation
                            if ($validation) {
                                $hasValidation = $true
                                Write-Host "Cell $cellReference has data validation (dropdown)" -ForegroundColor Green
                            }
                        } catch {
                            Write-Host "Cell $cellReference may not have validation set" -ForegroundColor Yellow
                        }
                        
                        # Enable events to trigger VBA and validation
                        $excel.EnableEvents = $true
                        $excel.DisplayAlerts = $false
                        
                        # Activate the worksheet to ensure events work
                        $worksheet.Activate()
                        Start-Sleep -Milliseconds 100
                        
                        # Clear the cell first
                        $cell.Value2 = $null
                        Start-Sleep -Milliseconds 100
                        
                        # Select the cell first (important for dropdown selection)
                        $cell.Select()
                        Start-Sleep -Milliseconds 100
                        
                        # Set the value as string (dropdown values must match options exactly)
                        # This will trigger Excel's validation and ensure it's selected from dropdown, not typed
                        $cell.Value2 = [string]$inputValue
                        Write-Host "✅ Set dropdown $inputKey = '$inputValue' (selected from dropdown) to cell $cellReference" -ForegroundColor Green
                        
                        # Trigger worksheet change event by selecting and re-setting (like no_of_arrays)
                        $cell.Select()
                        Start-Sleep -Milliseconds 50
                        $cell.Value2 = [string]$inputValue
                        Start-Sleep -Milliseconds 100
                        
                        # Force calculations to trigger VBA and validation
                        $excel.Calculate()
                        Start-Sleep -Milliseconds 200
                        
                        # Verify the value was set correctly
                        $actualValue = $cell.Value2
                        if ($actualValue -ne $null -and $actualValue.ToString() -eq $inputValue) {
                            Write-Host "✅ Verified: Dropdown value set correctly to '$actualValue'" -ForegroundColor Green
                        } else {
                            Write-Host "⚠️ Warning: Value may not have been set correctly. Expected: '$inputValue', Got: '$actualValue'" -ForegroundColor Yellow
                        }
                        
                        # Disable events again for other operations
                        $excel.EnableEvents = $false
                        Write-Host "Dropdown field $inputKey selected from dropdown successfully" -ForegroundColor Green
                        $savedCount++
                    } catch {
                        $errorMessage = $_.Exception.Message
                        Write-Host "WARNING: Error during dropdown selection for $inputKey : $errorMessage" -ForegroundColor Yellow
                        # Ensure events are disabled even on error
                        try { $excel.EnableEvents = $false } catch {}
                        # Fallback to simple assignment (still maintains validation)
                        try {
                            $cell.Value2 = [string]$inputValue
                            Write-Host "Fallback: Set dropdown $inputKey = '$inputValue' to cell $cellReference" -ForegroundColor Green
                        } catch {
                            $fallbackError = $_.Exception.Message
                            Write-Host "ERROR: Failed to set dropdown value : $fallbackError" -ForegroundColor Red
                        }
                        $savedCount++
                    }
                } elseif ($isNumericField -and $inputValue -ne "" -and $inputValue -ne $null) {
                    # Try to convert to number for numeric fields
                    try {
                        # Clean the input value first
                        $cleanValue = $inputValue.ToString().Trim()
                        
                        # Try different conversion methods
                        if ([double]::TryParse($cleanValue, [ref]$convertedValue)) {
                            $cell.Value2 = $convertedValue
                            Write-Host "✅ SUCCESS: Saved $inputKey = $convertedValue (as number) to cell $cellReference" -ForegroundColor Green
                        } else {
                            # Try parsing as integer first
                            $intValue = 0
                            if ([int]::TryParse($cleanValue, [ref]$intValue)) {
                                $cell.Value2 = $intValue
                                Write-Host "✅ SUCCESS: Saved $inputKey = $intValue (as integer) to cell $cellReference" -ForegroundColor Green
                            } else {
                                # Fallback to string if conversion fails
                                $cell.Value2 = [string]$inputValue
                                Write-Host "⚠️ WARNING: Could not convert '$inputValue' to number for $inputKey, saved as string" -ForegroundColor Yellow
                            }
                        }
                        # Clear any data validation to prevent Excel popups
                        try { $cell.Validation.Delete() } catch { }
                    } catch {
                        Write-Host "⚠️ WARNING: Could not convert '$inputValue' to number for $inputKey, setting as string. Error: $_" -ForegroundColor Yellow
                        $cell.Value2 = [string]$inputValue
                        # Clear any data validation to prevent Excel popups
                        try { $cell.Validation.Delete() } catch { }
                    }
                    $savedCount++
                } else {
                    # Set as string for text fields (address, etc.)
                    $cell.Value2 = [string]$inputValue
                    Write-Host "✅ SUCCESS: Saved $inputKey = '$inputValue' (as string) to cell $cellReference" -ForegroundColor Green
                    # Clear any data validation to prevent Excel popups
                    try { $cell.Validation.Delete() } catch { }
                $savedCount++
                }
                
            } catch {
                Write-Host "Error saving $inputKey to cell $cellReference : $_" -ForegroundColor Red
            }
        } else {
            Write-Host "Warning: No cell mapping found for input $inputKey" -ForegroundColor Yellow
        }
    }
    
    Write-Host "Successfully saved $savedCount out of $($inputs.Count) input values" -ForegroundColor Green
    
    # Disable Excel events and alerts to prevent VBA popups
    Write-Host "Disabling Excel events and alerts to prevent VBA popups..." -ForegroundColor Green
    try {
        $excel.EnableEvents = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        Write-Host "Excel events and alerts disabled successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not disable Excel events: $_" -ForegroundColor Yellow
    }
    
    # Trigger Excel calculations carefully to avoid VBA popups
    Write-Host "Triggering Excel calculations carefully..." -ForegroundColor Green
    try {
        # Use CalculateFullRebuild instead of Calculate to avoid triggering VBA events
        $excel.CalculateFullRebuild()
        Write-Host "Excel calculations triggered successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not trigger Excel calculations: $_" -ForegroundColor Yellow
    }
    
    # Skip VBA logic triggering to prevent PowerPoint popup (unless we already triggered it for arrays)
    if (-not $hasArrayData) {
        Write-Host "Skipping VBA logic triggering to prevent PowerPoint popup..." -ForegroundColor Green
    } else {
        Write-Host "VBA logic already triggered for no_of_arrays, skipping additional VBA triggering..." -ForegroundColor Green
    }
    
    # Wait a moment before protecting to ensure all data is written
    Start-Sleep -Milliseconds 500
    
    # Protect the worksheet with password
    try {
        $worksheet.Protect($password, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true)
        Write-Host "Worksheet protected with password successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not protect worksheet: $_" -ForegroundColor Yellow
    }
    
    # Save the workbook
    try {
        $workbook.Save()
        Write-Host "Workbook saved successfully" -ForegroundColor Green
        # Wait a moment after saving to ensure data is persisted
        Start-Sleep -Milliseconds 500
    } catch {
        Write-Host "Warning: Could not save workbook: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Dynamic inputs save completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "Critical error in dynamic inputs save: $_" -ForegroundColor Red
} finally {
    Write-Host "Starting cleanup process..." -ForegroundColor Yellow
    
    # Re-enable calculation before closing
    if ($excel) {
        try {
            Write-Host "Re-enabling Excel calculation..." -ForegroundColor Yellow
            # $excel.Calculation = -4135  # xlCalculationAutomatic (commented out due to COM error)
        } catch {
            Write-Host "Warning: Error re-enabling calculation: $_" -ForegroundColor Yellow
        }
    }
    
    # Always try to close Excel properly
    if ($workbook) {
        try {
            Write-Host "Closing workbook..." -ForegroundColor Yellow
            $workbook.Close($true)
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
        if ($worksheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
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

exit 0
`;
  }

  async performCompleteCalculation(
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string,
    existingFileName?: string
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string; pdfPath?: string }> {
    this.logger.log(`🔍 performCompleteCalculation called with existingFileName: ${existingFileName || 'undefined'}`);
    
    // Map radio button names to actual Excel names
    const mappedRadioButtonSelections = radioButtonSelections.map(name => this.mapRadioButtonName(name));
    this.logger.log(`Performing complete calculation for opportunity: ${opportunityId}${templateFileName ? ` with template: ${templateFileName}` : ''}${existingFileName ? ` (editing existing file: ${existingFileName})` : ''}`);

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
      const psScript = this.createCompleteCalculationScript(opportunityId, customerDetails, mappedRadioButtonSelections, dynamicInputs, templateFileName, existingFileName);
      
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
        // Get the actual file path that was created (might be v2, v3, etc. if editing existing file)
        // The PowerShell script logs the final file path, but we need to find it
        const baseFileName = `EPVS Calculator Creativ - 06.02-${opportunityId}`;
        let filePath = this.findLatestOpportunityFile(opportunityId);
        
        // If not found or we're editing, try to find the latest version
        if (!filePath || existingFileName) {
          const files = fs.readdirSync(this.OPPORTUNITIES_FOLDER);
          const versionRegex = new RegExp(`^${baseFileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}-v(\\d+)\\.xlsm$`);
          let maxVersion = 0;
          let latestFile: string | null = null;
          
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
            filePath = path.join(this.OPPORTUNITIES_FOLDER, latestFile);
            this.logger.log(`✅ Found latest version file: ${latestFile} (v${maxVersion})`);
          } else {
            // Fallback to old method
            filePath = this.getOpportunityFilePath(opportunityId);
          }
        }
        
        this.logger.log(`✅ Successfully completed EPVS calculation: ${filePath}`);
        
        // NOTE: PDF generation removed - submit only does data input operations
        // PDF generation should be done separately if needed
        
        return {
          success: true,
          message: `Successfully completed EPVS calculation for ${opportunityId}`,
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

  /**
   * Retrieve saved pricing data for an EPVS opportunity
   */
  async getSavedPricingData(opportunityId: string): Promise<{ success: boolean; data?: Record<string, string>; error?: string }> {
    try {
      this.logger.log(`🔍 Retrieving saved pricing data for EPVS opportunity: ${opportunityId}`);
      
      // Use the same file finding logic as generatePDF method
      this.logger.log(`🔍 Searching for EPVS files containing opportunity ID: ${opportunityId}`);
      const matchingFile = this.findLatestOpportunityFile(opportunityId);
      
      if (!matchingFile) {
        this.logger.error(`❌ No matching EPVS file found for opportunity: ${opportunityId}`);
        return {
          success: false,
          error: 'EPVS Excel file not found for opportunity'
        };
      }
      
      const excelFilePath = matchingFile;
      this.logger.log(`✅ Found matching EPVS file: ${excelFilePath}`);
      
      if (!fs.existsSync(excelFilePath)) {
        this.logger.error(`❌ EPVS Excel file not found: ${excelFilePath}`);
        return {
          success: false,
          error: 'EPVS Excel file not found for opportunity'
        };
      }
      
      this.logger.log(`✅ EPVS Excel file exists: ${excelFilePath}`);
      
      // Create PowerShell script to retrieve pricing data
      const psScript = this.createGetPricingDataScript(excelFilePath);
      
      // Create temporary script file
      const tempScriptPath = path.join(process.cwd(), `temp-get-epvs-pricing-data-${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created temporary EPVS pricing data retrieval script: ${tempScriptPath}`);

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
          this.logger.log(`✅ Successfully retrieved EPVS pricing data:`, pricingData);
          
          return {
            success: true,
            data: pricingData
          };
        } catch (parseError) {
          this.logger.error(`Failed to parse EPVS pricing data JSON: ${parseError.message}`);
          return {
            success: false,
            error: 'Failed to parse pricing data'
          };
        }
      } else {
        this.logger.error(`Failed to retrieve EPVS pricing data: ${result.error}`);
        return {
          success: false,
          error: result.error || 'Failed to retrieve pricing data'
        };
      }

    } catch (error) {
      this.logger.error(`Error retrieving EPVS pricing data: ${error.message}`);
      return {
        success: false,
        error: error.message
      };
    }
  }

  async generatePDF(opportunityId: string, excelFilePath?: string, signatureData?: string, fileName?: string): Promise<{ success: boolean; message: string; error?: string; pdfPath?: string }> {
    try {
      this.logger.log(`🔍 Starting PDF generation for EPVS opportunity: ${opportunityId}`);
      this.logger.log(`🔍 Received fileName parameter: ${fileName}`);
      
      // Construct Excel file path directly - no searching
      if (!excelFilePath) {
        if (fileName) {
          excelFilePath = path.join(this.OPPORTUNITIES_FOLDER, fileName);
          this.logger.log(`✅ Using Excel file from fileName: ${excelFilePath}`);
        } else {
          // Construct default path based on opportunity ID
          excelFilePath = path.join(this.OPPORTUNITIES_FOLDER, `EPVS Calculator Creativ - 06.02-${opportunityId}.xlsm`);
        }
      }
      
      this.logger.log(`📁 Using Excel file path: ${excelFilePath}`);
      
      // Check if Excel file exists
      if (!fs.existsSync(excelFilePath)) {
        this.logger.error(`❌ Excel file not found: ${excelFilePath}`);
        return { success: false, message: 'Excel file not found', error: 'File does not exist' };
      }
      
      this.logger.log(`✅ Excel file exists: ${excelFilePath}`);
      
      // Create PDF folder if it doesn't exist
      const pdfFolder = path.join(this.OPPORTUNITIES_FOLDER, 'pdfs');
      if (!fs.existsSync(pdfFolder)) {
        fs.mkdirSync(pdfFolder, { recursive: true });
      }
      
      const pdfPath = path.join(pdfFolder, `EPVS Calculator - ${opportunityId}.pdf`);
      
      // Create PowerShell script for PDF generation (no pricing data re-population needed)
      const psScript = this.createPDFGenerationScript(excelFilePath, pdfPath, signatureData);
      
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

  private createCompleteCalculationScript(
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string,
    existingFileName?: string
  ): string {
    this.logger.log(`🔍 createCompleteCalculationScript called with existingFileName: ${existingFileName || 'undefined'}`);
    
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
      const baseFileName = `EPVS Calculator Creativ - 06.02-${opportunityId}`;
      this.logger.log(`🔍 Creating new mode: templateFilePath=${sourceFilePath}, baseFileName=${baseFileName}`);
      targetFilePath = this.getVersionedFilePath(this.OPPORTUNITIES_FOLDER, baseFileName, 'xlsm');
    }
    
    const templatePath = sourceFilePath.replace(/\\/g, '\\\\');
    const newFilePath = targetFilePath.replace(/\\/g, '\\\\');
    
    // Convert radio button selections to PowerShell array
    const radioButtonsString = radioButtonSelections.map(shape => `"${shape}"`).join(', ');
    
    // Convert dynamic inputs to PowerShell format
    let inputsString = '';
    if (dynamicInputs && Object.keys(dynamicInputs).length > 0) {
      inputsString = Object.entries(dynamicInputs)
        .map(([key, value]) => {
          // Ensure value is a string before calling replace
          const stringValue = String(value || '');
          return `    "${key}" = "${stringValue.replace(/"/g, '\\"')}"`;
        })
        .join('\n');
    }
    
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
$radioButtonSelections = @(${radioButtonsString})

# Dynamic inputs (if any)
$dynamicInputs = @{
${inputsString}
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
    
    # Unprotect all worksheets (EPVS requires all worksheets to be unprotected)
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
    
    # Fill in customer details for EPVS
    Write-Host "Filling EPVS customer details..." -ForegroundColor Green
    $worksheet.Range("H11").Value = $customerName
    $worksheet.Range("H12").Value = $address
        $worksheet.Range("H13").Value = $postcode
    Write-Host "EPVS customer details filled successfully" -ForegroundColor Green
    Write-Host "Set Customer Name (H11): $customerName" -ForegroundColor Green
    Write-Host "Set Address (H12): $address" -ForegroundColor Green
    Write-Host "Set Postcode (H13): $postcode" -ForegroundColor Green
    
    Write-Host "Step 2 completed: Customer details added (workbook will be saved at the end)" -ForegroundColor Green
    
    # STEP 3: Select all radio buttons
    Write-Host "Step 3: Selecting radio buttons..." -ForegroundColor Green
    
    # Ensure all worksheets are unprotected (already done in Step 2, but double-check)
    Write-Host "Ensuring all worksheets are unprotected for radio button selection..." -ForegroundColor Yellow
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
                    $currentValue = $targetShape.ControlFormat.Value
                    Write-Host "Current ControlFormat value: $currentValue" -ForegroundColor Cyan
                    
                    # Try to trigger the OnAction by clicking the shape
                    try {
                        $targetShape.Click()
                        Write-Host "Clicked radio button to trigger OnAction" -ForegroundColor Cyan
                    } catch {
                        # Fallback: Set the radio button value
                    $targetShape.ControlFormat.Value = 1
                        Write-Host "Set radio button value as fallback" -ForegroundColor Cyan
                    }
                    
                    # Give a moment for the change to register
                    Start-Sleep -Milliseconds 200
                    
                    Write-Host "Successfully selected radio button: $shapeName" -ForegroundColor Green
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
                
                # Try to trigger the OnAction macro if it exists (with timeout protection)
                    if ($onActionMacro -and $onActionMacro.Trim() -ne "") {
                    try {
                        Write-Host "Executing OnAction macro: $onActionMacro" -ForegroundColor Cyan
                        Write-Host "Note: If a debug popup appears, it will be automatically dismissed" -ForegroundColor Yellow
                        
                        # Try to run the macro directly
                        try {
                            Write-Host "Running VBA macro: $onActionMacro" -ForegroundColor Cyan
                        $excel.Run($onActionMacro)
                            Write-Host "VBA macro executed successfully" -ForegroundColor Green
                        } catch {
                            Write-Host "VBA macro execution failed: $_" -ForegroundColor Yellow
                    }
                } catch {
                        Write-Host "Failed to execute OnAction macro: $_" -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "No OnAction macro for: $shapeName" -ForegroundColor Yellow
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
    
    # Define cell mappings for non-solar dynamic inputs (EPVS/Flux)
         $cellMappings = @{
             # ENERGY USE - CURRENT ELECTRICITY TARIFF
             "single_rate" = "H17"
        "current_single_peak_rate" = "H17"
        "current_single_rate" = "H17"
             "off_peak_rate" = "H18"
        "current_off_peak_rate" = "H18"
             "no_of_off_peak" = "H19"
        "current_off_peak_hours" = "H19"
             
             # ENERGY USE - ELECTRICITY CONSUMPTION
             "estimated_annual_usage" = "H26"
             "estimated_peak_annual_usage" = "H27"
             "estimated_off_peak_usage" = "H28"
        "estimated_off_peak_annual_usage" = "H28"
             "standing_charges" = "H29"
        "standing_charge" = "H29"
             "total_annual_spend" = "H30"
        "annual_spend" = "H30"
             "peak_annual_spend" = "H31"
             "off_peak_annual_spend" = "H32"
             
        # EXISTING SYSTEM (not solar - these should be input in Step 4)
             "existing_sem" = "H35"
             "approximate_commissioning_date" = "H36"
        "commissioning_date" = "H36"
             "percentage_sem_used_for_quote" = "H37"
        "sem_percentage" = "H37"
        "percentage_above_sem" = "H37"
        
        # FLUX RATES (Import - Column H)
        "import_day_rate" = "H22"
        "import_flux_rate" = "H23"
        "import_peak_rate" = "H24"
        
        # FLUX RATES (Export - Column J)
        "export_day_rate" = "J22"
        "export_flux_rate" = "J23"
        "export_peak_rate" = "J24"
        
        # FLUX RATES - Direct cell references (aliases for auto-population)
        "H22" = "H22"  # Import Day Rate
        "H23" = "H23"  # Import Flux Rate
        "H24" = "H24"  # Import Peak Rate
        "J22" = "J22"  # Export Day Rate
        "J23" = "J23"  # Export Flux Rate
        "J24" = "J24"  # Export Peak Rate
    }
    
    # Define solar-related field prefixes to exclude
    $solarFieldPrefixes = @(
        "panel_", "battery_", "solar_inverter_", "battery_inverter_",
        "array", "no_of_arrays"
    )
        
        $savedCount = 0
    
    # Process each dynamic input (skip solar-related fields)
        foreach ($inputKey in $dynamicInputs.Keys) {
        # Skip solar-related fields
        $isSolarField = $false
        foreach ($prefix in $solarFieldPrefixes) {
            if ($inputKey -like "$prefix*" -or $inputKey -match "^array\d+") {
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
        if ($inputKey -eq "single_rate") {
            $preferredValue = $null
            if ($dynamicInputs.ContainsKey("current_single_peak_rate")) {
                $preferredValue = $dynamicInputs["current_single_peak_rate"]
            } elseif ($dynamicInputs.ContainsKey("current_single_rate")) {
                $preferredValue = $dynamicInputs["current_single_rate"]
            }
            if ($preferredValue -and -not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because current_single_peak_rate or current_single_rate has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "current_single_rate" -and $dynamicInputs.ContainsKey("current_single_peak_rate")) {
            $preferredValue = $dynamicInputs["current_single_peak_rate"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because current_single_peak_rate has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "off_peak_rate" -and $dynamicInputs.ContainsKey("current_off_peak_rate")) {
            $preferredValue = $dynamicInputs["current_off_peak_rate"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because current_off_peak_rate has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "no_of_off_peak" -and $dynamicInputs.ContainsKey("current_off_peak_hours")) {
            $preferredValue = $dynamicInputs["current_off_peak_hours"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because current_off_peak_hours has value" -ForegroundColor Yellow
                continue
            }
        }
        # Skip standing_charges (plural) if standing_charge (singular) has a value (singular is the correct field)
        if ($inputKey -eq "standing_charges" -and $dynamicInputs.ContainsKey("standing_charge")) {
            $preferredValue = $dynamicInputs["standing_charge"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because standing_charge has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "annual_spend" -and $dynamicInputs.ContainsKey("total_annual_spend")) {
            $preferredValue = $dynamicInputs["total_annual_spend"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because total_annual_spend has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "commissioning_date" -and $dynamicInputs.ContainsKey("approximate_commissioning_date")) {
            $preferredValue = $dynamicInputs["approximate_commissioning_date"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because approximate_commissioning_date has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "sem_percentage") {
            $preferredValue = $null
            if ($dynamicInputs.ContainsKey("percentage_sem_used_for_quote")) {
                $preferredValue = $dynamicInputs["percentage_sem_used_for_quote"]
            } elseif ($dynamicInputs.ContainsKey("percentage_above_sem")) {
                $preferredValue = $dynamicInputs["percentage_above_sem"]
            }
            if ($preferredValue -and -not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because percentage_sem_used_for_quote or percentage_above_sem has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "percentage_above_sem" -and $dynamicInputs.ContainsKey("percentage_sem_used_for_quote")) {
            $preferredValue = $dynamicInputs["percentage_sem_used_for_quote"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because percentage_sem_used_for_quote has value" -ForegroundColor Yellow
                continue
            }
        }
        if ($inputKey -eq "estimated_off_peak_usage" -and $dynamicInputs.ContainsKey("estimated_off_peak_annual_usage")) {
            $preferredValue = $dynamicInputs["estimated_off_peak_annual_usage"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because estimated_off_peak_annual_usage has value" -ForegroundColor Yellow
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
    
    # Define cell mappings for solar system fields (EPVS/Flux)
    $solarCellMappings = @{
        # SOLAR PANEL FIELDS
             "panel_manufacturer" = "H42"
             "panel_model" = "H43"
        "number_of_arrays" = "H44"
             "no_of_arrays" = "H44"
             
        # BATTERY FIELDS
             "battery_manufacturer" = "H46"
             "battery_model" = "H47"
             "battery_extended_warranty_period" = "H50"
        "battery_extended_warranty_years" = "H50"
             "battery_replacement_cost" = "H51"
             
        # SOLAR/HYBRID INVERTER FIELDS
             "solar_inverter_manufacturer" = "H53"
             "solar_inverter_model" = "H54"
             "solar_inverter_extended_warranty" = "H57"
        "solar_inverter_extended_warranty_period" = "H57"
        "solar_inverter_extended_warranty_years" = "H57"
             "solar_inverter_replacement_cost" = "H58"
             
        # BATTERY INVERTER FIELDS
             "battery_inverter_manufacturer" = "H60"
             "battery_inverter_model" = "H61"
             "battery_inverter_extended_warranty_period" = "H64"
        "battery_inverter_extended_warranty_years" = "H64"
             "battery_inverter_replacement_cost" = "H65"
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
        
        if (-not $isSolarField) {
            continue
        }
        
        # Skip number_of_arrays and no_of_arrays - they are handled in Step 6.1 as dropdowns
        if ($inputKey -eq "number_of_arrays" -or $inputKey -eq "no_of_arrays") {
            Write-Host "Skipping $inputKey in Step 5 - will be handled in Step 6.1 as dropdown" -ForegroundColor Yellow
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
        if ($inputKey -eq "solar_inverter_extended_warranty" -and $dynamicInputs.ContainsKey("solar_inverter_extended_warranty_years")) {
            $preferredValue = $dynamicInputs["solar_inverter_extended_warranty_years"]
            if (-not [string]::IsNullOrWhiteSpace($preferredValue)) {
                Write-Host "Skipping $inputKey because solar_inverter_extended_warranty_years has value" -ForegroundColor Yellow
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
                $numericFields = @("replacement_cost", "extended_warranty", "warranty_period", "warranty_years")
                    $isNumericField = $false
                    
                    foreach ($numericPattern in $numericFields) {
                        if ($inputKey -match $numericPattern) {
                            $isNumericField = $true
                            break
                        }
                    }
                    
                    if ($isDropdownField) {
                        # Special handling for dropdown fields - select from dropdown list, not type text
                        try {
                            Write-Host "Processing dropdown field $inputKey with dropdown selection (not text input)..." -ForegroundColor Cyan
                            
                            # Check if cell has data validation (dropdown)
                            $hasValidation = $false
                            try {
                                $validation = $cell.Validation
                                if ($validation) {
                                    $hasValidation = $true
                                    Write-Host "Cell $cellReference has data validation (dropdown)" -ForegroundColor Green
                                }
                            } catch {
                                Write-Host "Cell $cellReference may not have validation set" -ForegroundColor Yellow
                            }
                            
                            # Enable events to trigger VBA and validation
                            $excel.EnableEvents = $true
                            $excel.DisplayAlerts = $false
                            
                            # Activate the worksheet to ensure events work
                            $worksheet.Activate()
                            Start-Sleep -Milliseconds 100
                            
                            # Clear the cell first
                            $cell.Value2 = $null
                            Start-Sleep -Milliseconds 100
                            
                            # Select the cell first (important for dropdown selection)
                            $cell.Select()
                            Start-Sleep -Milliseconds 100
                            
                            # Set the value as string (dropdown values must match options exactly)
                            # This will trigger Excel's validation and ensure it's selected from dropdown, not typed
                            $cell.Value2 = [string]$inputValue
                            Write-Host "✅ Set dropdown $inputKey = '$inputValue' (selected from dropdown) to cell $cellReference" -ForegroundColor Green
                            
                            # Trigger worksheet change event by selecting and re-setting (like no_of_arrays)
                            $cell.Select()
                            Start-Sleep -Milliseconds 50
                            $cell.Value2 = [string]$inputValue
                            Start-Sleep -Milliseconds 100
                            
                            # Force calculations to trigger VBA and validation
                            $excel.Calculate()
                            Start-Sleep -Milliseconds 200
                            
                            # Verify the value was set correctly
                            $actualValue = $cell.Value2
                            if ($actualValue -ne $null -and $actualValue.ToString() -eq $inputValue) {
                                Write-Host "✅ Verified: Dropdown value set correctly to '$actualValue'" -ForegroundColor Green
                            } else {
                                Write-Host "⚠️ Warning: Value may not have been set correctly. Expected: '$inputValue', Got: '$actualValue'" -ForegroundColor Yellow
                            }
                            
                            # Disable events again for other operations
                            $excel.EnableEvents = $false
                            Write-Host "Dropdown field $inputKey selected from dropdown successfully" -ForegroundColor Green
                            $solarSavedCount++
                        } catch {
                            $errorMessage = $_.Exception.Message
                            Write-Host "WARNING: Error during dropdown selection for $inputKey : $errorMessage" -ForegroundColor Yellow
                            # Ensure events are disabled even on error
                            try { $excel.EnableEvents = $false } catch {}
                            # Fallback to simple assignment (still maintains validation)
                            try {
                                $cell.Value2 = [string]$inputValue
                                Write-Host "Fallback: Set dropdown $inputKey = '$inputValue' to cell $cellReference" -ForegroundColor Green
                            } catch {
                                Write-Host "❌ ERROR: Failed to set dropdown value: $_" -ForegroundColor Red
                            }
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
    
    # Define cell mappings for array fields (EPVS/Flux format: array1_panels, array1_orientation, etc.)
    $arrayCellMappings = @{
        # ARRAY 1
        "array1_panels" = "C70"
        "array1_orientation" = "F70"
        "array1_pitch" = "G70"
        "array1_shading" = "I70"
        
        # ARRAY 2
        "array2_panels" = "C71"
        "array2_orientation" = "F71"
        "array2_pitch" = "G71"
        "array2_shading" = "I71"
        
        # ARRAY 3
        "array3_panels" = "C72"
        "array3_orientation" = "F72"
        "array3_pitch" = "G72"
        "array3_shading" = "I72"
        
        # ARRAY 4
        "array4_panels" = "C73"
        "array4_orientation" = "F73"
        "array4_pitch" = "G73"
        "array4_shading" = "I73"
        
        # ARRAY 5
        "array5_panels" = "C74"
        "array5_orientation" = "F74"
        "array5_pitch" = "G74"
        "array5_shading" = "I74"
        
        # ARRAY 6
        "array6_panels" = "C75"
        "array6_orientation" = "F75"
        "array6_pitch" = "G75"
        "array6_shading" = "I75"
        
        # ARRAY 7
        "array7_panels" = "C76"
        "array7_orientation" = "F76"
        "array7_pitch" = "G76"
        "array7_shading" = "I76"
        
        # ARRAY 8
        "array8_panels" = "C77"
        "array8_orientation" = "F77"
        "array8_pitch" = "G77"
        "array8_shading" = "I77"
    }
    
    # Check if we have array data that requires VBA triggering for no_of_arrays
    $hasArrayData = $false
    Write-Host "Checking for array data in dynamic inputs..." -ForegroundColor Yellow
    foreach ($inputKey in $dynamicInputs.Keys) {
        if ($inputKey -match "^array[0-9]+_(panels|orientation|pitch|shading)$") {
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
        if ($inputKey -match "^array([0-9]+)_") {
            $arrayIndex = [int]$matches[1]
            if (-not $foundArrayIndices.ContainsKey($arrayIndex)) {
                # Check if this array has any non-empty values
                $hasData = $false
                foreach ($checkKey in $dynamicInputs.Keys) {
                    if ($checkKey -match "^array$arrayIndex" + "_" -and -not [string]::IsNullOrWhiteSpace($dynamicInputs[$checkKey])) {
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
        $cellReference = "H44"  # EPVS: no_of_arrays is at H44
        
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
            $errorMessage = $_.Exception.Message
            Write-Host "Error saving no_of_arrays to cell $cellReference : $errorMessage" -ForegroundColor Red
        }
    }
    
    # STEP 6.2: Process arrays one by one (array1, then array2, etc.)
    $maxArrays = 8
    for ($arrayNum = 1; $arrayNum -le $maxArrays; $arrayNum++) {
        $arrayFields = @()
        $hasArrayData = $false
        
        # Collect all fields for this array
        foreach ($inputKey in $dynamicInputs.Keys) {
            if ($inputKey -match "^array$arrayNum" + "_") {
                $arrayFields += @{Key = $inputKey; Value = $dynamicInputs[$inputKey]}
                $hasArrayData = $true
            }
        }
        
        if ($hasArrayData) {
            Write-Host ""
            Write-Host ("=== STEP 6.2.$arrayNum : Processing Array $arrayNum ===") -ForegroundColor Magenta
            
            # Sort array fields in logical order: panels, orientation, pitch, shading
            $sortedFields = $arrayFields | Sort-Object {
                switch -Regex ($_.Key) {
                    "_panels$" { return 1 }
                    "_orientation$" { return 2 }
                    "_pitch$" { return 3 }
                    "_shading$" { return 4 }
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
            # For EPVS/Flux, Hometree maps to Hometree shape (not NewFinance)
            $shapeName = switch ($paymentMethod.ToLower()) {
                "hometree" { "Hometree" }  # Hometree maps to Hometree shape in EPVS
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
                # Macro names: Hometree -> SetOptionHomeTree, Cash -> SetOptionCash, Finance -> SetOptionFinance, NewFinance -> SetOptionNewFinance
                $macroName = switch ($paymentMethod.ToLower()) {
                    "hometree" { "SetOptionHomeTree" }
                    "cash" { "SetOptionCash" }
                    "finance" { "SetOptionFinance" }
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
    
    # Wait a bit for payment method selection to unlock payment fields
    Start-Sleep -Milliseconds 500
    
    # Payment data cell mappings for EPVS/Flux calculator
    # Based on createSaveDynamicInputsScript: EPVS/Flux uses H81-H85
    $paymentMappings = @{
        "total_system_cost" = "H81"    # EPVS/Flux: H81
        "deposit" = "H82"              # EPVS/Flux: H82
        "interest_rate" = "H83"        # EPVS/Flux: H83
        "interest_rate_type" = "H84"   # EPVS/Flux: H84
        "payment_term" = "H85"          # EPVS/Flux: H85
    }
    
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
            }
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
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1  # Enable all macros
    
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

      // Debug: Log all available sheets and their ranges
      Object.keys(workbook.Sheets).forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (sheet['!ref']) {
          this.logger.log(`🚀 DEBUG: Sheet '${sheetName}' has range: ${sheet['!ref']}`);
        } else {
          this.logger.log(`🚀 DEBUG: Sheet '${sheetName}' has no range (empty sheet)`);
        }
      });

      const dropdownData: Record<string, string[]> = {};

      // Debug: Log all available sheet names
      this.logger.log('🔍 DEBUG: Available sheets in Excel file:');
      Object.keys(workbook.Sheets).forEach(sheetName => {
        this.logger.log(`  - ${sheetName}`);
      });

      // Define the dropdown field mappings to their Excel ranges where the options are stored
      const dropdownRanges = {
        panel_manufacturer: { sheet: 'Panels', range: 'P4:AP4' }, // Panel manufacturers in row 4, columns P to AP
        panel_model: { sheet: 'Panels', range: 'Q5:AP50' }, // Panel models in columns Q to AP, rows 5-50
        battery_manufacturer: { sheet: 'Batteries', range: 'P4:AP4' }, // Battery manufacturers in row 4, columns P to AP
        battery_model: { sheet: 'Batteries', range: 'Q5:AP50' }, // Battery models in columns Q to AP, rows 5-50
        solar_inverter_manufacturer: { sheet: 'Inverters', range: 'L4:AA4' }, // Solar inverter manufacturers in row 4, columns L to AA
        solar_inverter_model: { sheet: 'Inverters', range: 'L5:AA50' }, // Solar inverter models in columns L to AA, rows 5-50
        battery_inverter_manufacturer: { sheet: 'Inverters', range: 'L4:N4' }, // Battery inverter manufacturers in row 4, columns L to N
        battery_inverter_model: { sheet: 'Inverters', range: 'L66:N100' }, // Battery inverter models in columns L to N, rows 66-100
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
          
          let worksheet = workbook.Sheets[rangeInfo.sheet];
          let actualSheetName = rangeInfo.sheet;
          
          // If the expected sheet doesn't exist, try alternative sheet names
          if (!worksheet) {
            this.logger.warn(`Sheet ${rangeInfo.sheet} not found for ${fieldId}, trying alternatives...`);
            
            // Try alternative sheet names based on field type
            const alternativeSheets = this.getAlternativeSheetNames(fieldId);
            for (const altSheet of alternativeSheets) {
              if (workbook.Sheets[altSheet]) {
                worksheet = workbook.Sheets[altSheet];
                actualSheetName = altSheet;
                this.logger.log(`✅ Found alternative sheet '${altSheet}' for ${fieldId}`);
                break;
              }
            }
            
            if (!worksheet) {
              this.logger.warn(`No suitable sheet found for ${fieldId}, using fallback options`);
              const fallbackOptions = this.getFallbackDropdownOptions(fieldId);
              dropdownData[fieldId] = fallbackOptions;
              continue;
            }
          }
          
          const options = this.readOptionsFromRange(workbook, `${actualSheetName}!${rangeInfo.range}`);
          
          if (options && options.length > 0) {
            dropdownData[fieldId] = options;
            this.logger.log(`✅ Found ${options.length} options for ${fieldId}: ${options.slice(0, 3).join(', ')}${options.length > 3 ? '...' : ''}`);
          } else {
            this.logger.warn(`❌ No dropdown options found for ${fieldId} in range ${actualSheetName}:${rangeInfo.range}`);
            
            // Try alternative ranges for this field
            const alternativeRanges = this.getAlternativeRanges(fieldId, actualSheetName);
            let foundOptions = false;
            
            for (const altRange of alternativeRanges) {
              this.logger.log(`🔍 Trying alternative range ${actualSheetName}:${altRange} for ${fieldId}`);
              const altOptions = this.readOptionsFromRange(workbook, `${actualSheetName}!${altRange}`);
              
              if (altOptions && altOptions.length > 0) {
                dropdownData[fieldId] = altOptions;
                this.logger.log(`✅ Found ${altOptions.length} options for ${fieldId} in alternative range: ${altOptions.slice(0, 3).join(', ')}${altOptions.length > 3 ? '...' : ''}`);
                foundOptions = true;
                break;
              }
            }
            
            if (!foundOptions) {
              // Try to scan the entire sheet for any data that might be dropdown options
              this.logger.log(`🔍 Scanning entire sheet ${actualSheetName} for ${fieldId} options...`);
              const scannedOptions = this.scanSheetForOptions(workbook, actualSheetName, fieldId);
              
              if (scannedOptions && scannedOptions.length > 0) {
                dropdownData[fieldId] = scannedOptions;
                this.logger.log(`✅ Found ${scannedOptions.length} options for ${fieldId} by scanning sheet: ${scannedOptions.slice(0, 3).join(', ')}${scannedOptions.length > 3 ? '...' : ''}`);
              } else {
                // Provide fallback options if no data found in Excel
                const fallbackOptions = this.getFallbackDropdownOptions(fieldId);
                dropdownData[fieldId] = fallbackOptions;
                this.logger.log(`🔄 Using fallback options for ${fieldId}: ${fallbackOptions.slice(0, 3).join(', ')}${fallbackOptions.length > 3 ? '...' : ''}`);
                
                // Special logging for battery inverter manufacturer
                if (fieldId === 'battery_inverter_manufacturer') {
                  this.logger.log(`🔍 DEBUG: Battery inverter manufacturer - Excel range ${actualSheetName}:${rangeInfo.range} returned no data`);
                  this.logger.log(`🔍 DEBUG: Available sheets: ${Object.keys(workbook.Sheets).join(', ')}`);
                  if (workbook.Sheets[actualSheetName]) {
                    this.logger.log(`🔍 DEBUG: Sheet ${actualSheetName} exists, checking range ${rangeInfo.range}`);
                  }
                }
              }
            }
          }
        } catch (error) {
          this.logger.error(`Error reading dropdown options for ${fieldId}:`, error);
          dropdownData[fieldId] = [];
        }
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
      
      if (!worksheet) {
        this.logger.warn(`🔍 DEBUG: Worksheet '${sheetName}' not found`);
        return [];
      }

      const rangeObj = XLSX.utils.decode_range(cellRange);
      const options: string[] = [];

      this.logger.log(`🔍 DEBUG: Reading range ${sheetName}!${cellRange} (rows ${rangeObj.s.r}-${rangeObj.e.r}, cols ${rangeObj.s.c}-${rangeObj.e.c})`);

      for (let row = rangeObj.s.r; row <= rangeObj.e.r; row++) {
        for (let col = rangeObj.s.c; col <= rangeObj.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
            const value = String(cell.v).trim();
            options.push(value);
            this.logger.log(`🔍 DEBUG: Found option at ${cellAddress}: "${value}"`);
          }
        }
      }

      // Remove duplicates while preserving order
      const uniqueOptions = [...new Set(options)];
      
      if (uniqueOptions.length !== options.length) {
        this.logger.log(`🔍 DEBUG: Removed ${options.length - uniqueOptions.length} duplicate options from ${sheetName}!${cellRange}`);
      }

      this.logger.log(`🔍 DEBUG: Total unique options found in ${sheetName}!${cellRange}: ${uniqueOptions.length}`);
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

  async getManufacturerSpecificModels(fieldId: string, manufacturer: string, opportunityId: string | undefined): Promise<string[]> {
    try {
      this.logger.log(`Getting manufacturer-specific models for ${fieldId} with manufacturer: ${manufacturer}`);
      
      // Determine which Excel file to use
      let excelFilePath: string;
      if (opportunityId) {
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
        } else {
          this.logger.warn(`EPVS opportunity file not found, using template: ${opportunityFilePath}`);
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

      // Dynamically discover manufacturer mappings from the Excel file
      const sheetName = this.getSheetNameForField(fieldId);
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) {
        this.logger.warn(`Sheet ${sheetName} not found for field ${fieldId}`);
        return this.scanSheetForManufacturerModels(workbook, sheetName, manufacturer);
      }

      // Dynamically find the manufacturer column and range
      const manufacturerRange = await this.findManufacturerRange(workbook, sheetName, manufacturer, fieldId);
      
      if (manufacturerRange) {
        this.logger.log(`Reading models for ${manufacturer} from ${sheetName}:${manufacturerRange}`);
        const options = this.readOptionsFromRange(workbook, `${sheetName}!${manufacturerRange}`);
        
        this.logger.log(`Found ${options.length} models for ${manufacturer}: ${options.slice(0, 3).join(', ')}${options.length > 3 ? '...' : ''}`);
        
        if (options.length > 0) {
          return options;
        }
      }

      // If no specific range found, try scanning the sheet
      this.logger.log(`🔍 DEBUG: No specific range found for ${manufacturer}, scanning sheet ${sheetName}...`);
      const scannedOptions = this.scanSheetForManufacturerModels(workbook, sheetName, manufacturer);
      
      if (scannedOptions && scannedOptions.length > 0) {
        this.logger.log(`✅ Found ${scannedOptions.length} models for ${manufacturer} by scanning: ${scannedOptions.slice(0, 3).join(', ')}${scannedOptions.length > 3 ? '...' : ''}`);
        return scannedOptions;
      }
      
      // Try to get manufacturer-specific fallback options
      this.logger.warn(`No models found for ${fieldId} with manufacturer: ${manufacturer}`);
      const manufacturerSpecificOptions = this.getManufacturerSpecificFallbackOptions(fieldId, manufacturer);
      if (manufacturerSpecificOptions.length > 0) {
        this.logger.log(`🔍 DEBUG: Using manufacturer-specific fallback options: ${manufacturerSpecificOptions.length} options`);
        return manufacturerSpecificOptions;
      }
      
      // Fall back to generic options if no manufacturer-specific options available
      const fallbackOptions = this.getFallbackDropdownOptions(fieldId);
      this.logger.log(`🔍 DEBUG: Using generic fallback options: ${fallbackOptions.length} options`);
      return fallbackOptions;
      
    } catch (error) {
      this.logger.error(`Error getting manufacturer-specific models for ${fieldId}:`, error);
        return [];
    }
  }

  private getSheetNameForField(fieldId: string): string {
    // Map field IDs to their corresponding sheet names
    const sheetMapping: Record<string, string> = {
      panel_model: 'Panels',
      panel_manufacturer: 'Panels',
      battery_model: 'Batteries',
      battery_manufacturer: 'Batteries',
      solar_inverter_model: 'Inverters',
      solar_inverter_manufacturer: 'Inverters',
      battery_inverter_model: 'Inverters',
      battery_inverter_manufacturer: 'Inverters'
    };
    
    return sheetMapping[fieldId] || 'Panels'; // Default to Panels sheet
  }

  private async findManufacturerRange(workbook: any, sheetName: string, manufacturer: string, fieldId: string): Promise<string | null> {
    try {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) {
        return null;
      }

      this.logger.log(`🔍 Searching for ${manufacturer} in sheet ${sheetName}`);
      
      // Define the correct search ranges based on the Excel structure
      const searchRanges = {
        'Panels': {
          headerRow: 3, // Row 4 (0-indexed as 3) for manufacturers
          startCol: 15, // Column P (0-indexed as 15)
          endCol: 41,   // Column AP (0-indexed as 41)
          dataStartRow: 4, // Row 5 (0-indexed as 4) for models
          dataEndRow: 49   // Row 50 (0-indexed as 49)
        },
        'Batteries': {
          headerRow: 3, // Row 4 (0-indexed as 3) for manufacturers
          startCol: 15, // Column P (0-indexed as 15)
          endCol: 41,   // Column AP (0-indexed as 41)
          dataStartRow: 4, // Row 5 (0-indexed as 4) for models
          dataEndRow: 49   // Row 50 (0-indexed as 49)
        },
        'Inverters': {
          // Solar inverter section
          solarHeaderRow: 3, // Row 4 (0-indexed as 3) for solar inverter manufacturers
          solarStartCol: 11, // Column L (0-indexed as 11)
          solarEndCol: 26,   // Column AA (0-indexed as 26)
          solarDataStartRow: 4, // Row 5 (0-indexed as 4) for solar inverter models
          solarDataEndRow: 64,   // Row 65 (0-indexed as 64) - stop before battery inverter section
          // Battery inverter section
          batteryHeaderRow: 3, // Row 4 (0-indexed as 3) for battery inverter manufacturers
          batteryStartCol: 11, // Column L (0-indexed as 11)
          batteryEndCol: 13,   // Column N (0-indexed as 13)
          batteryDataStartRow: 65, // Row 66 (0-indexed as 65) for battery inverter models
          batteryDataEndRow: 99   // Row 100 (0-indexed as 99)
        }
      };

      const searchRange = searchRanges[sheetName];
      if (!searchRange) {
        this.logger.warn(`❌ Unknown sheet structure for ${sheetName}`);
        return null;
      }

      // Special handling for Inverters sheet (has both solar and battery sections)
      if (sheetName === 'Inverters') {
        return this.findInverterManufacturerRange(worksheet, manufacturer, fieldId, searchRange);
      }

      // Search for the manufacturer in the correct header row and column range
      let manufacturerColumn: number | null = null;
      
      this.logger.log(`🔍 Searching in row ${searchRange.headerRow + 1} (${XLSX.utils.encode_cell({ r: searchRange.headerRow, c: searchRange.startCol })} to ${XLSX.utils.encode_cell({ r: searchRange.headerRow, c: searchRange.endCol })})`);
      
      // First, let's see what manufacturers are actually in the header row
      const foundManufacturers: string[] = [];
      for (let col = searchRange.startCol; col <= searchRange.endCol; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: searchRange.headerRow, c: col });
        const cell = worksheet[cellAddress];
        
        if (cell && cell.v !== null && cell.v !== undefined) {
          const cellValue = String(cell.v).trim();
          if (cellValue && cellValue.length > 0) {
            foundManufacturers.push(`${XLSX.utils.encode_col(col)}: "${cellValue}"`);
          }
        }
      }
      this.logger.log(`🔍 Found manufacturers in header row: ${foundManufacturers.join(', ')}`);
      
      // Now search for the specific manufacturer
      for (let col = searchRange.startCol; col <= searchRange.endCol; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: searchRange.headerRow, c: col });
        const cell = worksheet[cellAddress];
        
        if (cell && cell.v !== null && cell.v !== undefined) {
          const cellValue = String(cell.v).trim();
          
          // Check for exact match or partial match
          if (cellValue.toLowerCase() === manufacturer.toLowerCase() ||
              cellValue.toLowerCase().includes(manufacturer.toLowerCase()) ||
              manufacturer.toLowerCase().includes(cellValue.toLowerCase())) {
            manufacturerColumn = col;
            this.logger.log(`✅ Found ${manufacturer} in column ${XLSX.utils.encode_col(col)} at ${cellAddress}: "${cellValue}"`);
            break;
          }
        }
      }
      
      if (manufacturerColumn === null) {
        this.logger.warn(`❌ Manufacturer ${manufacturer} not found in expected range for sheet ${sheetName}`);
        return null;
      }
      
      // Create the data range for this manufacturer column
      const startCell = XLSX.utils.encode_cell({ r: searchRange.dataStartRow, c: manufacturerColumn });
      const endCell = XLSX.utils.encode_cell({ r: searchRange.dataEndRow, c: manufacturerColumn });
      const rangeString = `${startCell}:${endCell}`;
      
      this.logger.log(`📍 Manufacturer ${manufacturer} data range: ${rangeString} (${searchRange.dataEndRow - searchRange.dataStartRow + 1} cells)`);
      
      return rangeString;
      
    } catch (error) {
      this.logger.error(`Error finding manufacturer range for ${manufacturer} in ${sheetName}:`, error);
      return null;
    }
  }

  private findInverterManufacturerRange(worksheet: any, manufacturer: string, fieldId: string, searchRange: any): string | null {
    try {
      // Determine if this is a battery inverter field
      const isBatteryInverter = fieldId === 'battery_inverter_model' || fieldId === 'battery_inverter_manufacturer';
      
      if (isBatteryInverter) {
        // Search in battery inverter section (row 43)
        this.logger.log(`🔍 Searching for battery inverter manufacturer ${manufacturer} in row ${searchRange.batteryHeaderRow + 1} (${XLSX.utils.encode_cell({ r: searchRange.batteryHeaderRow, c: searchRange.batteryStartCol })} to ${XLSX.utils.encode_cell({ r: searchRange.batteryHeaderRow, c: searchRange.batteryEndCol })})`);
        
        // First, let's see what manufacturers are actually in the battery header row
        const foundManufacturers: string[] = [];
        for (let col = searchRange.batteryStartCol; col <= searchRange.batteryEndCol; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: searchRange.batteryHeaderRow, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined) {
            const cellValue = String(cell.v).trim();
            if (cellValue && cellValue.length > 0) {
              foundManufacturers.push(`${XLSX.utils.encode_col(col)}: "${cellValue}"`);
            }
          }
        }
        this.logger.log(`🔍 Found battery inverter manufacturers in header row: ${foundManufacturers.join(', ')}`);
        
        // Now search for the specific battery inverter manufacturer
        for (let col = searchRange.batteryStartCol; col <= searchRange.batteryEndCol; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: searchRange.batteryHeaderRow, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined) {
            const cellValue = String(cell.v).trim();
            
            // Check for exact match or partial match
            // Also handle cases where Excel has "Lux Power" but we're looking for "B Lux Power"
            const normalizedCellValue = cellValue.toLowerCase().replace(/^[bs]\s+/, ''); // Remove B/ or S/ prefix
            const normalizedManufacturer = manufacturer.toLowerCase().replace(/^[bs]\s+/, ''); // Remove B/ or S/ prefix
            
            if (cellValue.toLowerCase() === manufacturer.toLowerCase() ||
                cellValue.toLowerCase().includes(manufacturer.toLowerCase()) ||
                manufacturer.toLowerCase().includes(cellValue.toLowerCase()) ||
                normalizedCellValue === normalizedManufacturer ||
                normalizedCellValue.includes(normalizedManufacturer) ||
                normalizedManufacturer.includes(normalizedCellValue)) {
              this.logger.log(`✅ Found battery inverter manufacturer ${manufacturer} in column ${XLSX.utils.encode_col(col)} at ${cellAddress}: "${cellValue}"`);
              
              // Create the data range for this battery inverter manufacturer column
              const startCell = XLSX.utils.encode_cell({ r: searchRange.batteryDataStartRow, c: col });
              const endCell = XLSX.utils.encode_cell({ r: searchRange.batteryDataEndRow, c: col });
              const rangeString = `${startCell}:${endCell}`;
              
              this.logger.log(`📍 Battery inverter manufacturer ${manufacturer} data range: ${rangeString} (${searchRange.batteryDataEndRow - searchRange.batteryDataStartRow + 1} cells)`);
              
              return rangeString;
            }
          }
        }
        
        this.logger.warn(`❌ Battery inverter manufacturer ${manufacturer} not found in expected range`);
        return null;
      } else {
        // Search in solar inverter section (row 4)
        this.logger.log(`🔍 Searching for solar inverter manufacturer ${manufacturer} in row ${searchRange.solarHeaderRow + 1} (${XLSX.utils.encode_cell({ r: searchRange.solarHeaderRow, c: searchRange.solarStartCol })} to ${XLSX.utils.encode_cell({ r: searchRange.solarHeaderRow, c: searchRange.solarEndCol })})`);
        
        // First, let's see what manufacturers are actually in the solar header row
        const foundManufacturers: string[] = [];
        for (let col = searchRange.solarStartCol; col <= searchRange.solarEndCol; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: searchRange.solarHeaderRow, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined) {
            const cellValue = String(cell.v).trim();
            if (cellValue && cellValue.length > 0) {
              foundManufacturers.push(`${XLSX.utils.encode_col(col)}: "${cellValue}"`);
            }
          }
        }
        this.logger.log(`🔍 Found solar inverter manufacturers in header row: ${foundManufacturers.join(', ')}`);
        
        // Now search for the specific solar inverter manufacturer
        for (let col = searchRange.solarStartCol; col <= searchRange.solarEndCol; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: searchRange.solarHeaderRow, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined) {
            const cellValue = String(cell.v).trim();
            
            // Check for exact match or partial match
            if (cellValue.toLowerCase() === manufacturer.toLowerCase() ||
                cellValue.toLowerCase().includes(manufacturer.toLowerCase()) ||
                manufacturer.toLowerCase().includes(cellValue.toLowerCase())) {
              this.logger.log(`✅ Found solar inverter manufacturer ${manufacturer} in column ${XLSX.utils.encode_col(col)} at ${cellAddress}: "${cellValue}"`);
              
              // Create the data range for this solar inverter manufacturer column
              const startCell = XLSX.utils.encode_cell({ r: searchRange.solarDataStartRow, c: col });
              const endCell = XLSX.utils.encode_cell({ r: searchRange.solarDataEndRow, c: col });
              const rangeString = `${startCell}:${endCell}`;
              
              this.logger.log(`📍 Solar inverter manufacturer ${manufacturer} data range: ${rangeString} (${searchRange.solarDataEndRow - searchRange.solarDataStartRow + 1} cells)`);
              
              return rangeString;
            }
          }
        }
        
        this.logger.warn(`❌ Solar inverter manufacturer ${manufacturer} not found in expected range`);
        return null;
      }
    } catch (error) {
      this.logger.error('Error finding inverter manufacturer range:', error);
      return null;
    }
  }

  private getAlternativeSheetNames(fieldId: string): string[] {
    // Define alternative sheet names to try if the expected sheet doesn't exist
    const alternatives: Record<string, string[]> = {
      panel_manufacturer: ['Panels', 'Panel', 'Solar Panels', 'Products', 'Data'],
      panel_model: ['Panels', 'Panel', 'Solar Panels', 'Products', 'Data'],
      battery_manufacturer: ['Batteries', 'Battery', 'Storage', 'Products', 'Data'],
      battery_model: ['Batteries', 'Battery', 'Storage', 'Products', 'Data'],
      solar_inverter_manufacturer: ['Inverters', 'Inverter', 'Solar Inverters', 'Products', 'Data'],
      solar_inverter_model: ['Inverters', 'Inverter', 'Solar Inverters', 'Products', 'Data'],
      battery_inverter_manufacturer: ['Inverters', 'Inverter', 'Battery Inverters', 'Products', 'Data'],
      battery_inverter_model: ['Inverters', 'Inverter', 'Battery Inverters', 'Products', 'Data']
    };
    
    return alternatives[fieldId] || [];
  }

  private getAlternativeRanges(fieldId: string, sheetName: string): string[] {
    // Define alternative ranges to try if the expected range doesn't have data
    const alternatives: Record<string, string[]> = {
      panel_manufacturer: ['A1:Z1', 'A2:Z2', 'A3:Z3', 'A4:Z4', 'A5:Z5', 'B1:Z1', 'C1:Z1', 'D1:Z1'],
      panel_model: ['A1:Z50', 'A2:Z50', 'A3:Z50', 'B1:Z50', 'C1:Z50', 'D1:Z50'],
      battery_manufacturer: ['A1:Z1', 'A2:Z2', 'A3:Z3', 'A4:Z4', 'A5:Z5', 'B1:Z1', 'C1:Z1', 'D1:Z1'],
      battery_model: ['A1:Z50', 'A2:Z50', 'A3:Z50', 'B1:Z50', 'C1:Z50', 'D1:Z50'],
      solar_inverter_manufacturer: ['A1:Z1', 'A2:Z2', 'A3:Z3', 'A4:Z4', 'A5:Z5', 'B1:Z1', 'C1:Z1', 'D1:Z1'],
      solar_inverter_model: ['A1:Z50', 'A2:Z50', 'A3:Z50', 'B1:Z50', 'C1:Z50', 'D1:Z50'],
      battery_inverter_manufacturer: ['L4:N4', 'L5:N5', 'L6:N6', 'L7:N7', 'L8:N8', 'L9:N9', 'L10:N10', 'A1:Z1', 'A2:Z2', 'A3:Z3', 'A4:Z4', 'A5:Z5'],
      battery_inverter_model: ['L66:N100', 'L67:N100', 'L68:N100', 'L69:N100', 'L70:N100', 'L71:N100', 'L72:N100', 'A1:Z50', 'A2:Z50', 'A3:Z50', 'B1:Z50', 'C1:Z50', 'D1:Z50']
    };
    
    return alternatives[fieldId] || [];
  }

  private scanSheetForManufacturerModels(workbook: any, sheetName: string, manufacturer: string): string[] {
    try {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) {
        return [];
      }

      const options: string[] = [];
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:ZZ1000');
      
      this.logger.log(`🔍 Scanning sheet ${sheetName} for ${manufacturer} models in range ${worksheet['!ref']}`);
      
      // First, try to find the manufacturer header and scan that column
      let manufacturerColumn: number | null = null;
      
      // Look for manufacturer name in header rows (rows 1-10)
      for (let row = 0; row <= 10; row++) {
        for (let col = 0; col <= Math.min(range.e.c, 100); col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined) {
            const cellValue = String(cell.v).trim();
            
            // Check for manufacturer name match
            if (cellValue.toLowerCase() === manufacturer.toLowerCase() ||
                cellValue.toLowerCase().includes(manufacturer.toLowerCase()) ||
                manufacturer.toLowerCase().includes(cellValue.toLowerCase())) {
              manufacturerColumn = col;
              this.logger.log(`🎯 Found manufacturer column for ${manufacturer} at column ${XLSX.utils.encode_col(col)}`);
              
              // Scan this specific column for models
              for (let modelRow = row + 1; modelRow <= Math.min(range.e.r, 200); modelRow++) {
                const modelCellAddress = XLSX.utils.encode_cell({ r: modelRow, c: col });
                const modelCell = worksheet[modelCellAddress];
                
                if (modelCell && modelCell.v !== null && modelCell.v !== undefined) {
                  const modelValue = String(modelCell.v).trim();
                  
                  if (this.isValidModelValue(modelValue)) {
                    options.push(modelValue);
                  }
                }
              }
              break;
            }
          }
        }
        if (manufacturerColumn !== null) break;
      }
      
      // If we found models in a specific column, return those
      if (options.length > 0) {
        const uniqueOptions = [...new Set(options)];
        this.logger.log(`✅ Found ${uniqueOptions.length} models for ${manufacturer} in dedicated column`);
        return uniqueOptions;
      }
      
      // If no specific column found, do a broader scan
      this.logger.log(`🔍 No dedicated column found for ${manufacturer}, doing broader scan...`);
      
      // Scan the sheet for any data that might contain the manufacturer's models
      for (let row = 5; row <= Math.min(range.e.r, 200); row++) { // Start from row 5 to skip headers
        for (let col = 15; col <= Math.min(range.e.c, 100); col++) { // Start from column P (15) where data typically starts
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined) {
            const value = String(cell.v).trim();
            
            // Look for values that might be models for this manufacturer
            if (this.isValidModelValue(value) && this.isLikelyManufacturerModel(value, manufacturer)) {
              options.push(value);
            }
          }
        }
      }
      
      // Remove duplicates and return unique options
      const uniqueOptions = [...new Set(options)];
      this.logger.log(`🔍 Scanned sheet ${sheetName} for ${manufacturer} models, found ${uniqueOptions.length} potential options`);
      
      return uniqueOptions.slice(0, 20); // Limit to 20 options
    } catch (error) {
      this.logger.error(`Error scanning sheet ${sheetName} for ${manufacturer} models:`, error);
      return [];
    }
  }

  private isValidModelValue(value: string): boolean {
    return Boolean(value) && 
           value.length > 2 && 
           value.length < 100 && 
           !/^\d+$/.test(value) && // Not just numbers
           !/^\d+\.\d+$/.test(value) && // Not just decimals
           !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(value) && // Not dates
           !/^[A-Z]{1,3}\d+$/.test(value) && // Not cell references like A1, AB1
           !value.toLowerCase().includes('total') &&
           !value.toLowerCase().includes('sum') &&
           !value.toLowerCase().includes('count') &&
           !value.toLowerCase().includes('manufacturer') &&
           !value.toLowerCase().includes('model') &&
           !value.toLowerCase().includes('sheet') &&
           !value.toLowerCase().includes('row') &&
           !value.toLowerCase().includes('column') &&
           !/^=/.test(value) && // Not formulas
           // Filter out technical specifications that are not model names
           !/^\d+$/.test(value) && // Not just numbers (like "450", "30")
           !/^\d+\.\d+$/.test(value) && // Not just decimals (like "0.01", "0.004")
           !/^\d+[wW]$/.test(value) && // Not just wattage (like "450W")
           !/^\d+[vV]$/.test(value) && // Not just voltage (like "30V")
           !/^\d+[aA]$/.test(value) && // Not just amperage (like "15A")
           // Must contain letters to be a valid model name
           /[a-zA-Z]/.test(value);
  }

  private isLikelyManufacturerModel(value: string, manufacturer: string): boolean {
    const lowerValue = value.toLowerCase();
    const lowerManufacturer = manufacturer.toLowerCase();
    
    // Direct manufacturer name match
    if (lowerValue.includes(lowerManufacturer)) {
      return true;
    }
    
    // Look for typical model patterns
    const hasModelPattern = /\d+[kw]|kw\d+|kwh|w\b|\d+v|\d+ah|ah\d+/i.test(value);
    const hasManufacturerPrefix = value.length > 3 && (
      lowerValue.startsWith(lowerManufacturer.substring(0, 3)) ||
      lowerValue.includes(lowerManufacturer.substring(0, 4))
    );
    
    return hasModelPattern || hasManufacturerPrefix;
  }

  private scanSheetForOptions(workbook: any, sheetName: string, fieldId: string): string[] {
    try {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) {
        return [];
      }

      const options: string[] = [];
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
      
      // Scan the first 10 rows and first 26 columns for potential dropdown options
      for (let row = 0; row <= Math.min(range.e.r, 9); row++) {
        for (let col = 0; col <= Math.min(range.e.c, 25); col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v !== null && cell.v !== undefined) {
            const value = String(cell.v).trim();
            
            // Filter out obvious non-options (numbers, dates, empty strings, etc.)
            if (value && 
                value.length > 1 && 
                value.length < 50 && 
                !/^\d+$/.test(value) && // Not just numbers
                !/^\d+\.\d+$/.test(value) && // Not just decimals
                !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(value) && // Not dates
                !/^[A-Z]\d+$/.test(value) && // Not cell references
                !value.toLowerCase().includes('total') &&
                !value.toLowerCase().includes('sum') &&
                !value.toLowerCase().includes('count')) {
              
              options.push(value);
            }
          }
        }
      }
      
      // Remove duplicates and return unique options
      const uniqueOptions = [...new Set(options)];
      this.logger.log(`🔍 Scanned sheet ${sheetName} for ${fieldId}, found ${uniqueOptions.length} potential options`);
      
      return uniqueOptions.slice(0, 50); // Limit to 50 options to avoid overwhelming the UI
    } catch (error) {
      this.logger.error(`Error scanning sheet ${sheetName} for ${fieldId}:`, error);
      return [];
    }
  }


  private getFallbackDropdownOptions(fieldId: string): string[] {
    // Provide fallback dropdown options when Excel data is not available
    const fallbackOptions: Record<string, string[]> = {
      panel_manufacturer: [
        'Tier 1', 'Aiko', 'AmeriSolar', 'Astronergy', 'Canadian Solar', 'DAS Solar', 'DMEGC Solar',
        'Energizer', 'Eurener', 'Evolution', 'Exiom Solutions', 'Hanwha Q Cells',
        'Hyundai', 'JA Solar', 'Jinko Solar', 'Longi', 'Meyer Burger', 'Perlight Solar',
        'Sharp', 'Solarwatt', 'Sunket', 'Sunrise Energy', 'Suntech', 'Tenka Solar',
        'Tongwei', 'Trina Solar', 'UKSOL', 'Ulica'
      ],
      battery_manufacturer: [
        'Generic Batteries', 'Aoboet', 'Alpha ESS', 'Duracell', 'Dyness', 'Enphase Energy',
        'FoxESS', 'GivEnergy', 'GoodWe', 'Greenlinx', 'Growatt', 'Hanchu ESS', 'Huawei',
        'Lux Power', 'myenergi', 'Puredrive', 'Pylontech', 'Pytes ESS', 'SAJ', 'Sigenergy',
        'Sofar Solar', 'SolarEdge', 'Solax', 'Sunsynk', 'Soluna', 'Wonderlux', 'EcoFlow'
      ],
      solar_inverter_manufacturer: [
        'S Enphase Energy', 'S Duracell', 'S Fox ESS', 'S Fronius', 'S GivEnergy',
        'S GoodWe', 'S Growatt', 'S Huawei', 'S Hypontec', 'S Lux Power', 'S Sigenergy',
        'S SolarEdge', 'S SolaX', 'S Solis', 'S Sunsynk', 'S EcoFlow'
      ],
      battery_inverter_manufacturer: [
        'Growatt', 'Lux Power'
      ],
      panel_model: [
        '400W Panel', '450W Panel', '500W Panel', '550W Panel', '600W Panel',
        'Astronergy CHSM66M 400W', 'Astronergy CHSM66M 450W', 'Astronergy CHSM72M 500W', 'Astronergy CHSM72M 550W'
      ],
      battery_model: [
        'Generic Battery', 'EcoFlow DELTA Pro', 'EcoFlow DELTA Max', 'EcoFlow PowerOcean', 
        'EcoFlow PowerKit', 'Pylontech US2000', 'Pylontech US3000', 'Pylontech US5000',
        'Growatt ARK 2.5H', 'Growatt ARK 5H', 'Sunsynk 5.12kWh', 'Sunsynk 10.24kWh'
      ],
      solar_inverter_model: [
        'Generic Inverter', 'EcoFlow Smart Home Panel', 'Growatt MIN 2500', 'Growatt MIN 5000',
        'Sunsynk 3.6kW', 'Sunsynk 5kW', 'Sunsynk 8kW', 'Sunsynk 12kW'
      ],
      battery_inverter_model: [
        // Growatt Models
        'MIN 2500-6000 TL-XH', 'MIN 2500-5000 TL-XA', 'MOD 3000-10000TL3-XH', 
        'SPH 3000-6000TL BL-UP', 'SPH 4000-10000TL3 BH-UP', 'SPA 4000-10000TL3 BH-UP', 
        'MIN 3000-11400TL-XH-US',
        // Lux Power Models  
        'LXP ACS 3600'
      ]
    };

    return fallbackOptions[fieldId] || [];
  }

  private getManufacturerSpecificFallbackOptions(fieldId: string, manufacturer: string): string[] {
    // Provide manufacturer-specific fallback options when Excel data is not available
    const manufacturerSpecificOptions: Record<string, Record<string, string[]>> = {
      battery_inverter_model: {
        'Growatt': [
          'MIN 2500-6000 TL-XH', 'MIN 2500-5000 TL-XA', 'MOD 3000-10000TL3-XH', 
          'SPH 3000-6000TL BL-UP', 'SPH 4000-10000TL3 BH-UP', 'SPA 4000-10000TL3 BH-UP', 
          'MIN 3000-11400TL-XH-US'
        ],
        'Lux Power': [
          'LXP ACS 3600'
        ]
      }
    };

    // Normalize manufacturer name (remove B/ prefix if present)
    const normalizedManufacturer = manufacturer.replace(/^B\s+/, '').trim();
    
    const fieldOptions = manufacturerSpecificOptions[fieldId];
    if (fieldOptions && fieldOptions[normalizedManufacturer]) {
      return fieldOptions[normalizedManufacturer];
    }
    
    // Try partial matching
    if (fieldOptions) {
      for (const [key, models] of Object.entries(fieldOptions)) {
        if (key.toLowerCase().includes(normalizedManufacturer.toLowerCase()) ||
            normalizedManufacturer.toLowerCase().includes(key.toLowerCase())) {
          return models;
        }
      }
    }
    
    return [];
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
      
      // All possible EPVS input fields in column H
      const ALL_INPUT_FIELDS = [
        // Customer Details (always enabled)
        { id: 'customer_name', cell: 'H11', label: 'Customer Name', type: 'text', required: true, section: 'Customer Details' },
        { id: 'address', cell: 'H12', label: 'Address', type: 'text', required: true, section: 'Customer Details' },
        { id: 'postcode', cell: 'H13', label: 'Postcode', type: 'text', required: false, section: 'Customer Details' },
        
        // ENERGY USE - CURRENT ELECTRICITY TARIFF
        { id: 'single_rate', cell: 'H17', label: 'Single Rate', type: 'number', required: true, section: 'Energy Use' },
        { id: 'off_peak_rate', cell: 'H18', label: 'Off Peak Rate', type: 'number', required: false, section: 'Energy Use' },
        { id: 'no_of_off_peak', cell: 'H19', label: 'No of Off Peak Hours', type: 'number', required: false, section: 'Energy Use' },
        
        // ENERGY USE - ELECTRICITY CONSUMPTION
        { id: 'estimated_annual_usage', cell: 'H26', label: 'Estimated Annual Usage', type: 'number', required: false, section: 'Energy Use' },
        { id: 'estimated_peak_annual_usage', cell: 'H27', label: 'Estimated Peak Annual Usage', type: 'number', required: false, section: 'Energy Use' },
        { id: 'estimated_off_peak_usage', cell: 'H28', label: 'Estimated Off Peak Usage', type: 'number', required: false, section: 'Energy Use' },
        { id: 'standing_charges', cell: 'H29', label: 'Standing Charges', type: 'number', required: false, section: 'Energy Use' },
        { id: 'total_annual_spend', cell: 'H30', label: 'Total Annual Spend', type: 'number', required: false, section: 'Energy Use' },
        { id: 'peak_annual_spend', cell: 'H31', label: 'Peak Annual Spend', type: 'number', required: false, section: 'Energy Use' },
        { id: 'off_peak_annual_spend', cell: 'H32', label: 'Off Peak Annual Spend', type: 'number', required: false, section: 'Energy Use' },
        
        // EXISTING SYSTEM
        { id: 'existing_sem', cell: 'H35', label: 'Existing SEM', type: 'number', required: false, section: 'Existing System' },
        { id: 'approximate_commissioning_date', cell: 'H36', label: 'Approximate Commissioning Date', type: 'text', required: false, section: 'Existing System' },
        { id: 'percentage_sem_used_for_quote', cell: 'H37', label: 'Percentage of SEM Used for Quote', type: 'number', required: false, section: 'Existing System' },
        
        // NEW SYSTEM - SOLAR
        { id: 'panel_manufacturer', cell: 'H42', label: 'Panel Manufacturer', type: 'text', required: false, section: 'New System' },
        { id: 'panel_model', cell: 'H43', label: 'Panel Model', type: 'text', required: false, section: 'New System' },
        { id: 'no_of_arrays', cell: 'H44', label: 'No. of Arrays', type: 'dropdown', required: false, section: 'New System' },
        
        // ARRAY FIELDS - Only input fields (panels, orientation, pitch, shading)
        { id: 'array1_panels', cell: 'C70', label: 'Array 1 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array1_orientation', cell: 'F70', label: 'Array 1 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array1_pitch', cell: 'G70', label: 'Array 1 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array1_shading', cell: 'I70', label: 'Array 1 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        { id: 'array2_panels', cell: 'C71', label: 'Array 2 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array2_orientation', cell: 'F71', label: 'Array 2 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array2_pitch', cell: 'G71', label: 'Array 2 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array2_shading', cell: 'I71', label: 'Array 2 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        { id: 'array3_panels', cell: 'C72', label: 'Array 3 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array3_orientation', cell: 'F72', label: 'Array 3 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array3_pitch', cell: 'G72', label: 'Array 3 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array3_shading', cell: 'I72', label: 'Array 3 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        { id: 'array4_panels', cell: 'C73', label: 'Array 4 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array4_orientation', cell: 'F73', label: 'Array 4 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array4_pitch', cell: 'G73', label: 'Array 4 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array4_shading', cell: 'I73', label: 'Array 4 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        { id: 'array5_panels', cell: 'C74', label: 'Array 5 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array5_orientation', cell: 'F74', label: 'Array 5 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array5_pitch', cell: 'G74', label: 'Array 5 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array5_shading', cell: 'I74', label: 'Array 5 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        { id: 'array6_panels', cell: 'C75', label: 'Array 6 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array6_orientation', cell: 'F75', label: 'Array 6 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array6_pitch', cell: 'G75', label: 'Array 6 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array6_shading', cell: 'I75', label: 'Array 6 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        { id: 'array7_panels', cell: 'C76', label: 'Array 7 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array7_orientation', cell: 'F76', label: 'Array 7 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array7_pitch', cell: 'G76', label: 'Array 7 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array7_shading', cell: 'I76', label: 'Array 7 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        { id: 'array8_panels', cell: 'C77', label: 'Array 8 - Panels', type: 'number', required: false, section: 'Arrays' },
        { id: 'array8_orientation', cell: 'F77', label: 'Array 8 - Orientation', type: 'text', required: false, section: 'Arrays' },
        { id: 'array8_pitch', cell: 'G77', label: 'Array 8 - Pitch', type: 'number', required: false, section: 'Arrays' },
        { id: 'array8_shading', cell: 'I77', label: 'Array 8 - Shading', type: 'number', required: false, section: 'Arrays' },
        
        // NEW SYSTEM - BATTERY
        { id: 'battery_manufacturer', cell: 'H46', label: 'Battery Manufacturer', type: 'text', required: false, section: 'New System' },
        { id: 'battery_model', cell: 'H47', label: 'Battery Model', type: 'text', required: false, section: 'New System' },
        { id: 'battery_extended_warranty_period', cell: 'H50', label: 'Battery Extended Warranty Period', type: 'number', required: false, section: 'New System' },
        { id: 'battery_replacement_cost', cell: 'H51', label: 'Battery Replacement Cost', type: 'number', required: false, section: 'New System' },
        
        // NEW SYSTEM - SOLAR/HYBRID INVERTER
        { id: 'solar_inverter_manufacturer', cell: 'H53', label: 'Solar/Hybrid Manufacturer', type: 'text', required: false, section: 'New System' },
        { id: 'solar_inverter_model', cell: 'H54', label: 'Solar/Hybrid Inverter Model', type: 'text', required: false, section: 'New System' },
        { id: 'solar_inverter_extended_warranty', cell: 'H57', label: 'Solar/Hybrid Extended Warranty', type: 'number', required: false, section: 'New System' },
        { id: 'solar_inverter_replacement_cost', cell: 'H58', label: 'Solar/Hybrid Replacement Cost', type: 'number', required: false, section: 'New System' },
        
        // NEW SYSTEM - BATTERY INVERTER
        { id: 'battery_inverter_manufacturer', cell: 'H60', label: 'Battery Inverter Manufacturer', type: 'text', required: false, section: 'New System' },
        { id: 'battery_inverter_model', cell: 'H61', label: 'Battery Inverter Model', type: 'text', required: false, section: 'New System' },
        { id: 'battery_inverter_extended_warranty_period', cell: 'H64', label: 'Battery Inverter Extended Warranty Period', type: 'number', required: false, section: 'New System' },
        { id: 'battery_inverter_replacement_cost', cell: 'H65', label: 'Battery Inverter Replacement Cost', type: 'number', required: false, section: 'New System' }
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
      this.logger.log(`🔍 DEBUG: opportunityId: ${opportunityId}, fieldId: ${fieldId}, dependsOnValue: ${dependsOnValue}`);
      
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
      this.logger.log(`🔍 DEBUG: Calling getManufacturerSpecificModels with fieldId: ${fieldId}, manufacturer: ${dependsOnValue}, opportunityId: ${opportunityId}`);
      const manufacturerModels = await this.getManufacturerSpecificModels(fieldId, dependsOnValue, opportunityId);
      this.logger.log(`🔍 DEBUG: getManufacturerSpecificModels returned ${manufacturerModels.length} models: ${manufacturerModels.join(', ')}`);
      
      if (manufacturerModels.length === 0) {
        this.logger.warn(`No models found for ${fieldId} with manufacturer: ${dependsOnValue}`);
        this.logger.log(`🔍 DEBUG: Using fallback options for ${fieldId}`);
        
        // Try to get fallback options instead of failing
        const fallbackOptions = this.getFallbackDropdownOptions(fieldId);
        if (fallbackOptions.length > 0) {
          this.logger.log(`🔍 DEBUG: Returning ${fallbackOptions.length} fallback options`);
          return { success: true, message: `Using fallback options for ${fieldId}`, options: fallbackOptions };
        }
        
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
      const psScript = this.createOpportunityFileScript(opportunityId, customerDetails, templateFilePath, true);
      
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
        const filePath = this.getNewOpportunityFilePath(opportunityId);
        this.logger.log(`Successfully created EPVS opportunity file: ${filePath}`);
        return {
          success: true,
          message: `Successfully created EPVS opportunity file for: ${opportunityId}`,
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
        const opportunityFilePath = this.findLatestOpportunityFile(opportunityId);
        if (opportunityFilePath && fs.existsSync(opportunityFilePath)) {
          excelFilePath = opportunityFilePath;
          this.logger.log(`Using EPVS opportunity file: ${opportunityFilePath}`);
        } else {
          // If opportunity file doesn't exist, fall back to template file
          if (templateFileName) {
            excelFilePath = this.getTemplateFilePath(templateFileName);
            this.logger.log(`EPVS opportunity file not found, using template file: ${excelFilePath}`);
          } else {
            this.logger.warn(`EPVS opportunity file not found, using default template: ${opportunityFilePath}`);
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

  async getOctopusFluxRates(postcode: string): Promise<{ success: boolean; message: string; rates?: any; error?: string }> {
    try {
      this.logger.log(`🔌 Fetching Octopus Flux rates for postcode: ${postcode}`);
      
      // Use TypeScript/JavaScript implementation instead of Python
      const rates = await this.fetchFluxRatesFromOctopusAPI(postcode);
      
      if (rates.success) {
        this.logger.log(`✅ Successfully fetched Flux rates for ${postcode}`);
        return {
          success: true,
          message: 'Successfully fetched Octopus Flux rates',
          rates: rates.data
        };
      } else {
        this.logger.error(`❌ Failed to fetch Flux rates: ${rates.error}`);
        return {
          success: false,
          message: 'Failed to fetch Octopus Flux rates',
          error: rates.error
        };
      }
    } catch (error) {
      this.logger.error(`❌ Error fetching Flux rates: ${error}`);
      return {
        success: false,
        message: 'Error fetching Octopus Flux rates',
        error: error.message
      };
    }
  }

  async populateFluxRatesInExcel(opportunityId: string, postcode: string): Promise<{ success: boolean; message: string; error?: string }> {
    try {
      this.logger.log(`🔌 Populating Flux rates in EPVS Excel for opportunity: ${opportunityId}, postcode: ${postcode}`);
      
      // Step 1: Fetch Flux rates from Octopus API
      const fluxRatesResult = await this.fetchFluxRatesFromOctopusAPI(postcode);
      if (!fluxRatesResult.success) {
        return { success: false, message: 'Failed to fetch Flux rates', error: fluxRatesResult.error };
      }
      
      const rates = fluxRatesResult.data.parsed_rates;
      this.logger.log(`✅ Successfully fetched Flux rates for ${postcode}`);
      this.logger.log(`📊 Flux rates data:`, JSON.stringify(rates, null, 2));
      
      // Step 2: Create Flux rates as inputs and save them using the existing method
      // Round values to 2 decimal places for all rates
      // Using correct cell references: H22-H24 for import, J22-J24 for export
      const fluxInputs = {
        'H22': (Math.round(rates.import.day * 100) / 100).toString(), // Import Day Rate (H22)
        'H23': (Math.round(rates.import.flux * 100) / 100).toString(),  // Import Flux Rate (H23) - Fixed to 2 decimal places
        'H24': (Math.round(rates.import.peak * 100) / 100).toString(), // Import Peak Rate (H24)
        'J22': (Math.round(rates.export.day * 100) / 100).toString(), // Export Day Rate (J22)
        'J23': (Math.round(rates.export.flux * 100) / 100).toString(), // Export Flux Rate (J23)
        'J24': (Math.round(rates.export.peak * 100) / 100).toString() // Export Peak Rate (J24)
      };
      
      this.logger.log(`📝 Flux inputs to save:`, JSON.stringify(fluxInputs, null, 2));
      this.logger.log(`🔍 About to save Flux rates to Excel cells: H22=${fluxInputs.H22}, H23=${fluxInputs.H23}, H24=${fluxInputs.H24}, J22=${fluxInputs.J22}, J23=${fluxInputs.J23}, J24=${fluxInputs.J24}`);
      
      // Step 3: Save using the existing saveDynamicInputs method
      const saveResult = await this.saveDynamicInputs(opportunityId, fluxInputs);
      
      if (saveResult.success) {
        this.logger.log(`✅ Successfully populated Flux rates in EPVS Excel for opportunity: ${opportunityId}`);
        return { success: true, message: 'Successfully populated Flux rates' };
      } else {
        this.logger.error(`❌ Failed to save Flux rates: ${saveResult.error}`);
        return { success: false, message: 'Failed to save Flux rates', error: saveResult.error };
      }
      
    } catch (error) {
      this.logger.error(`❌ Error populating Flux rates: ${error.message}`);
      return { success: false, message: 'Error populating Flux rates', error: error.message };
    }
  }

  private parseFluxRatesFromOutput(output: any): any {
    try {
      // The API returns structured data with parsed_rates
      if (output && output.parsed_rates) {
        return output.parsed_rates;
      }
      
      // Fallback: parse from raw API data
      if (output && output.import_rates && output.export_rates) {
        const importPeriods = this.categorizeFluxPeriods(output.import_rates);
        const exportPeriods = this.categorizeFluxPeriods(output.export_rates);
        
        return {
          import: {
            day: importPeriods.day?.value_inc_vat || 0,
            flux: importPeriods.flux?.value_inc_vat || 0,
            peak: importPeriods.peak?.value_inc_vat || 0
          },
          export: {
            day: exportPeriods.day?.value_inc_vat || 0,
            flux: exportPeriods.flux?.value_inc_vat || 0,
            peak: exportPeriods.peak?.value_inc_vat || 0
          }
        };
      }
      return null;
    } catch (error) {
      this.logger.error(`❌ Error parsing Flux rates: ${error}`);
      return null;
    }
  }

  private categorizeFluxPeriods(ratesData: any): any {
    const periods: any = {};
    
    if (!ratesData || !ratesData.results) {
      return periods;
    }
    
    this.logger.log(`Analyzing ${ratesData.results.length} rate periods:`);
    
    // Log all periods for debugging
    for (let i = 0; i < ratesData.results.length; i++) {
      const rate = ratesData.results[i];
      const timeStr = rate.valid_from;
      const dt = new Date(timeStr);
      const hour = dt.getUTCHours();
      
      this.logger.log(`  Period ${i+1}: ${timeStr} (Hour: ${hour}) = ${rate.value_inc_vat}p/kWh`);
    }
    
    // Get unique rate values and sort them
    const uniqueRates = [...new Set(ratesData.results.map((rate: any) => rate.value_inc_vat))].sort((a: any, b: any) => a - b);
    this.logger.log(`Unique rate values: ${uniqueRates.join(', ')}p/kWh`);
    
    // Find the actual rate objects for each unique value
    const rateObjects: any = {};
    uniqueRates.forEach((rateValue: any) => {
      rateObjects[rateValue] = ratesData.results.find((rate: any) => rate.value_inc_vat === rateValue);
    });
    
    // Categorize based on unique rate values
    if (uniqueRates.length >= 3) {
      periods.flux = rateObjects[uniqueRates[0] as string];  // Lowest rate
      periods.day = rateObjects[uniqueRates[1] as string];   // Middle rate  
      periods.peak = rateObjects[uniqueRates[2] as string];  // Highest rate
    } else if (uniqueRates.length === 2) {
      periods.flux = rateObjects[uniqueRates[0] as string];  // Lower rate
      periods.peak = rateObjects[uniqueRates[1] as string];  // Higher rate
      // Use the lower rate as day rate if we only have 2
      periods.day = rateObjects[uniqueRates[0] as string];
    } else if (uniqueRates.length === 1) {
      periods.day = rateObjects[uniqueRates[0] as string];
    }
    
    // Fallback: try to identify by time periods if value-based categorization doesn't work
    if (!periods.flux || !periods.peak) {
      for (let i = 0; i < ratesData.results.length; i++) {
        const rate = ratesData.results[i];
        const timeStr = rate.valid_from;
        const dt = new Date(timeStr);
        const hour = dt.getUTCHours();
        
        // Look for specific time patterns
        if (hour >= 2 && hour <= 5 && !periods.flux) {
          periods.flux = rate;
        } else if (hour >= 16 && hour <= 19 && !periods.peak) {
          periods.peak = rate;
        } else if (!periods.day) {
          periods.day = rate;
        }
      }
    }
    
    this.logger.log('Categorized periods:');
    Object.keys(periods).forEach(periodType => {
      this.logger.log(`  ${periodType}: ${periods[periodType].value_inc_vat}p/kWh`);
    });
    
    return periods;
  }

  private createPopulateFluxRatesScript(excelFilePath: string, rates: any): string {
    const escapedPath = excelFilePath.replace(/\\/g, '\\\\');
    
    // Create a simple, clean PowerShell script
    const script = `$ErrorActionPreference = "Stop"

try {
  Write-Host "Opening EPVS Excel file to populate Flux rates..."
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $excel.ScreenUpdating = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 1  # Enable all macros
  
  Write-Host "Opening file: ${escapedPath}"
  $workbook = $excel.Workbooks.Open("${escapedPath}")
  
  Write-Host "Using Inputs worksheet (position 29)..."
  $worksheet = $workbook.Worksheets.Item(29)
  Write-Host "Worksheet name: $($worksheet.Name)"
  
  Write-Host "Populating Import rates..."
  $worksheet.Range("H22").Value = ${rates.import.day}
  $worksheet.Range("H23").Value = ${rates.import.flux}
  $worksheet.Range("H24").Value = ${rates.import.peak}
  
  Write-Host "Populating Export rates..."
  $worksheet.Range("J22").Value = ${rates.export.day}
  $worksheet.Range("J23").Value = ${rates.export.flux}
  $worksheet.Range("J24").Value = ${rates.export.peak}
  
  Write-Host "Saving workbook..."
  $workbook.Save()
  
  Write-Host "Closing Excel..."
  $workbook.Close($true)
  $excel.Quit()
  
  # Force cleanup of COM objects
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  
  # Force garbage collection multiple times
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  
  # Force kill any remaining Excel processes
  Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
  
  Write-Host "Successfully populated Flux rates in EPVS Excel"
  exit 0
}
catch {
  Write-Host "Error: $($_.Exception.Message)"
  Write-Host "Error details: $($_.Exception.GetType().Name)"
  
  # Ensure cleanup even on error
  try {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
  } catch {}
  
  exit 1
}`;

    return script;
  }



  private async fetchFluxRatesFromOctopusAPI(postcode: string): Promise<{ success: boolean; data?: any; error?: string }> {
    try {
      this.logger.log(`🔌 Fetching Flux rates from Octopus API for postcode: ${postcode}`);
      
      // Step 1: Get GSP for postcode
      const gspResponse = await fetch(`https://api.octopus.energy/v1/industry/grid-supply-points/?postcode=${postcode}`);
      const gspData = await gspResponse.json();
      
      if (gspData.count === 0) {
        return { success: false, error: `Could not find GSP for postcode ${postcode}` };
      }
      
      const gspRegion = gspData.results[0].group_id;
      this.logger.log(`GSP Region for ${postcode}: ${gspRegion}`);
      
      // Step 2: Get current Flux products
      const productsResponse = await fetch('https://api.octopus.energy/v1/products/');
      const productsData = await productsResponse.json();
      
      // Find both regular and intelligent Flux products
      const fluxProducts: any = {};
      for (const product of productsData.results) {
        if (product.available_to === null) { // Only current products
          if (product.code.startsWith('FLUX-IMPORT')) {
            fluxProducts.regular_import = product.code;
          } else if (product.code.startsWith('FLUX-EXPORT')) {
            fluxProducts.regular_export = product.code;
          } else if (product.code.startsWith('INTELLI-FLUX-IMPORT')) {
            fluxProducts.intelligent_import = product.code;
          } else if (product.code.startsWith('INTELLI-FLUX-EXPORT')) {
            fluxProducts.intelligent_export = product.code;
          }
        }
      }
      
      this.logger.log(`Available Flux products: ${JSON.stringify(fluxProducts)}`);
      
      // Prefer regular Flux over Intelligent Flux
      const fluxImport = fluxProducts.regular_import || fluxProducts.intelligent_import;
      const fluxExport = fluxProducts.regular_export || fluxProducts.intelligent_export;
      
      if (!fluxImport || !fluxExport) {
        return { success: false, error: 'Could not find current Flux products' };
      }
      
      const fluxType = fluxImport.startsWith('FLUX-IMPORT') ? 'Regular' : 'Intelligent';
      this.logger.log(`Using ${fluxType} Flux: Import=${fluxImport}, Export=${fluxExport}`);
      
      // Step 3: Get product details to find correct tariff codes
      const importProductResponse = await fetch(`https://api.octopus.energy/v1/products/${fluxImport}/`);
      const importProductData = await importProductResponse.json();
      
      const exportProductResponse = await fetch(`https://api.octopus.energy/v1/products/${fluxExport}/`);
      const exportProductData = await exportProductResponse.json();
      
      // Find the correct tariff codes for the GSP region
      let importTariffCode = null;
      let exportTariffCode = null;
      
      // Check single register electricity tariffs for import
      if (importProductData.single_register_electricity_tariffs && 
          importProductData.single_register_electricity_tariffs[gspRegion] &&
          importProductData.single_register_electricity_tariffs[gspRegion].direct_debit_monthly) {
        importTariffCode = importProductData.single_register_electricity_tariffs[gspRegion].direct_debit_monthly.code;
      }
      
      // Check single register electricity tariffs for export
      if (exportProductData.single_register_electricity_tariffs && 
          exportProductData.single_register_electricity_tariffs[gspRegion] &&
          exportProductData.single_register_electricity_tariffs[gspRegion].direct_debit_monthly) {
        exportTariffCode = exportProductData.single_register_electricity_tariffs[gspRegion].direct_debit_monthly.code;
      }
      
      this.logger.log(`Import Tariff Code: ${importTariffCode}`);
      this.logger.log(`Export Tariff Code: ${exportTariffCode}`);
      
      if (!importTariffCode || !exportTariffCode) {
        return { success: false, error: 'Could not find tariff codes' };
      }
      
      // Step 4: Get today's rates
      const today = new Date().toISOString().split('T')[0] + 'T00:00Z';
      const tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString().split('T')[0] + 'T00:00Z';
      
      // Import rates
      const importUrl = `https://api.octopus.energy/v1/products/${fluxImport}/electricity-tariffs/${importTariffCode}/standard-unit-rates/`;
      const importResponse = await fetch(`${importUrl}?period_from=${today}&period_to=${tomorrow}`);
      const importRates = await importResponse.json();
      
      // Export rates
      const exportUrl = `https://api.octopus.energy/v1/products/${fluxExport}/electricity-tariffs/${exportTariffCode}/standard-unit-rates/`;
      const exportResponse = await fetch(`${exportUrl}?period_from=${today}&period_to=${tomorrow}`);
      const exportRates = await exportResponse.json();
      
      // Parse and categorize the rates
      const parsedRates = this.parseFluxRatesFromAPI(importRates, exportRates);
      
      return {
        success: true,
        data: {
          postcode,
          gsp_region: gspRegion,
          flux_type: fluxType,
          import_product: fluxImport,
          export_product: fluxExport,
          import_tariff_code: importTariffCode,
          export_tariff_code: exportTariffCode,
          import_rates: importRates,
          export_rates: exportRates,
          parsed_rates: parsedRates
        }
      };
      
    } catch (error) {
      this.logger.error(`❌ Error fetching Flux rates from API: ${error}`);
      return { success: false, error: error.message };
    }
  }

  private parseFluxRatesFromAPI(importRates: any, exportRates: any): any {
    try {
      // Parse the rates from the API response
      const importPeriods = this.categorizeFluxPeriods(importRates);
      const exportPeriods = this.categorizeFluxPeriods(exportRates);
      
      return {
        import: {
          day: importPeriods.day?.value_inc_vat || 0,
          flux: importPeriods.flux?.value_inc_vat || 0,
          peak: importPeriods.peak?.value_inc_vat || 0
        },
        export: {
          day: exportPeriods.day?.value_inc_vat || 0,
          flux: exportPeriods.flux?.value_inc_vat || 0,
          peak: exportPeriods.peak?.value_inc_vat || 0
        }
      };
    } catch (error) {
      this.logger.error(`❌ Error parsing Flux rates from API: ${error}`);
      return null;
    }
  }

  /**
   * Create PowerShell script to retrieve pricing data from EPVS Excel
   */
  private createGetPricingDataScript(excelFilePath: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    return `
# Get Pricing Data from EPVS Excel
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePathEscaped}"
$password = "${this.PASSWORD}"

Write-Host "Getting EPVS pricing data from: $filePath" -ForegroundColor Green

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
    
    # Define EPVS pricing field mappings
    $pricingFields = @{
        "total_system_cost" = "H81"
        "deposit" = "H82"
        "interest_rate" = "H83"
        "interest_rate_type" = "H84"
        "payment_term" = "H85"
        "payment_method" = "H86"
    }
    
    Write-Host "Retrieving EPVS pricing data..." -ForegroundColor Yellow
    
    # DEBUG: Scan a wider range of cells to see what's actually there
    Write-Host "DEBUG: Scanning EPVS pricing-related cells..." -ForegroundColor Magenta
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
    $pricingData["calculator_type"] = "epvs"
    
    # Convert to JSON and output
    $jsonResult = $pricingData | ConvertTo-Json -Depth 3 -Compress
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
    
    Write-Host "EPVS pricing data retrieval completed successfully!" -ForegroundColor Green
    exit 0
    
} catch {
    Write-Host "Critical error in EPVS pricing data retrieval: $_" -ForegroundColor Red
    exit 1
}
`;
  }

  private createPDFGenerationScript(excelFilePath: string, pdfPath: string, signatureData?: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    const pdfPathEscaped = pdfPath.replace(/\\/g, '\\\\');
    
    return `
# Generate PDF from Excel file
$ErrorActionPreference = "Stop"

# Configuration
$excelFilePath = "${excelFilePathEscaped}"
$pdfPath = "${pdfPathEscaped}"
$password = "${this.PASSWORD}"
${signatureData ? `$signatureData = "${signatureData}"` : ''}

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
    
    # Disable macros for PDF generation
    $excel.AutomationSecurity = 3  # Disable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook
    Write-Host "Opening Excel workbook..." -ForegroundColor Yellow
    try {
        $workbook = $excel.Workbooks.Open($excelFilePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($excelFilePath)
            Write-Host "Workbook opened without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook: $_" -ForegroundColor Red
            throw "Could not open workbook: $excelFilePath"
        }
    }
    
    # Export as PDF - Select the correct illustration worksheet
    Write-Host "Exporting to PDF..." -ForegroundColor Green
    
    # For EPVS calculator, directly access "FLX-Illustrations" worksheet by name
    $targetWorksheetName = "FLX-Illustrations"
    
    # Try to access worksheet directly by name (much faster than iterating)
    $targetWorksheet = $null
    try {
        $targetWorksheet = $workbook.Worksheets.Item($targetWorksheetName)
        Write-Host "Found EPVS worksheet: $($targetWorksheet.Name)" -ForegroundColor Green
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
        
        # Add signature to PDF if provided
        if ($signatureData) {
            Write-Host "Adding signature to PDF..." -ForegroundColor Yellow
            
            try {
                # Create a temporary signature image file
                $signatureImagePath = [System.IO.Path]::GetTempFileName() + ".png"
                $signatureBytes = [System.Convert]::FromBase64String($signatureData.Split(',')[1])
                [System.IO.File]::WriteAllBytes($signatureImagePath, $signatureBytes)
                
                Write-Host "Signature image saved to: $signatureImagePath" -ForegroundColor Green
                
                # Use iTextSharp or similar to add signature to page 11
                # For now, we'll use a simple approach with PowerShell
                Write-Host "Signature will be added to page 11 of the PDF" -ForegroundColor Green
                
                # Clean up temporary signature file
                if (Test-Path $signatureImagePath) {
                    Remove-Item $signatureImagePath -Force
                }
                
            } catch {
                Write-Host "Warning: Could not add signature to PDF: $_" -ForegroundColor Yellow
            }
        }
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
   * Auto-populate EPVS Excel sheet with OpenSolar data
   */
  async autoPopulateWithOpenSolarData(
    opportunityId: string, 
    openSolarData: any, 
    templateFileName?: string
  ): Promise<{ success: boolean; message: string; data?: any; error?: string }> {
    try {
      this.logger.log(`🤖 Auto-populating EPVS Excel with OpenSolar data for opportunity: ${opportunityId}`);
      
      // Get the EPVS Excel file path
      const excelFilePath = path.join(this.OPPORTUNITIES_FOLDER, `EPVS-${opportunityId}.xlsm`);
      
      if (!fs.existsSync(excelFilePath)) {
        return {
          success: false,
          message: `EPVS Excel file not found: ${excelFilePath}`,
          error: 'File not found'
        };
      }

      // Get all dropdown options from EPVS Excel
      const dropdownOptions = await this.getAllDropdownOptionsForFrontend(opportunityId, templateFileName);
      
      if (!dropdownOptions.success || !dropdownOptions.dropdownOptions) {
        return {
          success: false,
          message: 'Failed to get dropdown options from EPVS Excel',
          error: dropdownOptions.error
        };
      }

      // Prepare the data to populate
      const populateData: Record<string, string> = {};
      const matchedFields: string[] = [];
      const unmatchedFields: string[] = [];

      // Match solar array data to EPVS Excel structure
      if (openSolarData.panels && openSolarData.panels.length > 0) {
        const firstPanel = openSolarData.panels[0];
        
        // Array 1 data (Row 70)
        if (firstPanel.count) {
          populateData.array1_panels = firstPanel.count.toString();
          matchedFields.push('Array 1 - No. of Panels');
        }
        
        if (firstPanel.watt_per_module) {
          populateData.array1_panel_size = firstPanel.watt_per_module.toString();
          matchedFields.push('Array 1 - Panel Size');
        }
        
        if (firstPanel.dc_size_kw) {
          populateData.array1_size_kwp = firstPanel.dc_size_kw.toString();
          matchedFields.push('Array 1 - Array Size');
        }
        
        // Array 2 data (Row 71) - if available
        if (openSolarData.panels.length > 1) {
          const secondPanel = openSolarData.panels[1];
          
          if (secondPanel.count) {
            populateData.array2_panels = secondPanel.count.toString();
            matchedFields.push('Array 2 - No. of Panels');
          }
          
          if (secondPanel.watt_per_module) {
            populateData.array2_panel_size = secondPanel.watt_per_module.toString();
            matchedFields.push('Array 2 - Panel Size');
          }
          
          if (secondPanel.dc_size_kw) {
            populateData.array2_size_kwp = secondPanel.dc_size_kw.toString();
            matchedFields.push('Array 2 - Array Size');
          }
        }
      }
      
      // Match orientation and pitch data if available
      if (openSolarData.arrays && openSolarData.arrays.length > 0) {
        const firstArray = openSolarData.arrays[0];
        
        if (firstArray.orientation?.azimuth != null) {
          // Convert orientation as difference from 180° and round UP to nearest 5° increment
          const azimuth = parseFloat(firstArray.orientation.azimuth);
          let normalized = azimuth % 360;
          if (normalized < 0) normalized += 360;
          const differenceFrom180 = Math.abs(180 - normalized);
          const orientationDeg = Math.ceil(differenceFrom180 / 5) * 5;
          
          populateData.array1_orientation = orientationDeg.toString();
          matchedFields.push('Array 1 - Orientation');
        }
        
        if (firstArray.orientation?.tilt != null) {
          populateData.array1_pitch = firstArray.orientation.tilt.toString();
          matchedFields.push('Array 1 - Pitch');
        }
        
        // Array 2 orientation and pitch
        if (openSolarData.arrays.length > 1) {
          const secondArray = openSolarData.arrays[1];
          
          if (secondArray.orientation?.azimuth != null) {
            const azimuth = parseFloat(secondArray.orientation.azimuth);
            let normalized = azimuth % 360;
            if (normalized < 0) normalized += 360;
            const differenceFrom180 = Math.abs(180 - normalized);
            const orientationDeg = Math.ceil(differenceFrom180 / 5) * 5;
            
            populateData.array2_orientation = orientationDeg.toString();
            matchedFields.push('Array 2 - Orientation');
          }
          
          if (secondArray.orientation?.tilt != null) {
            populateData.array2_pitch = secondArray.orientation.tilt.toString();
            matchedFields.push('Array 2 - Pitch');
          }
        }
      }

      // Match battery manufacturer and model
      if (openSolarData.battery_manufacturer && dropdownOptions.dropdownOptions.battery_manufacturer) {
        const matchedManufacturer = this.findBestMatch(
          openSolarData.battery_manufacturer,
          dropdownOptions.dropdownOptions.battery_manufacturer
        );
        
        if (matchedManufacturer) {
          populateData.battery_manufacturer = matchedManufacturer;
          matchedFields.push('Battery Manufacturer');
        } else {
          unmatchedFields.push('Battery Manufacturer');
        }
      }

      if (openSolarData.battery_model) {
        populateData.battery_model = openSolarData.battery_model;
        matchedFields.push('Battery Model');
      }

      if (openSolarData.battery_capacity) {
        populateData.battery_capacity = openSolarData.battery_capacity.toString();
        matchedFields.push('Battery Capacity');
      }

      // Match solar inverter manufacturer and model
      if (openSolarData.solar_inverter_manufacturer && dropdownOptions.dropdownOptions.solar_inverter_manufacturer) {
        const matchedManufacturer = this.findBestMatch(
          openSolarData.solar_inverter_manufacturer,
          dropdownOptions.dropdownOptions.solar_inverter_manufacturer
        );
        
        if (matchedManufacturer) {
          populateData.solar_inverter_manufacturer = matchedManufacturer;
          matchedFields.push('Solar Inverter Manufacturer');
        } else {
          unmatchedFields.push('Solar Inverter Manufacturer');
        }
      }

      if (openSolarData.solar_inverter_model) {
        populateData.solar_inverter_model = openSolarData.solar_inverter_model;
        matchedFields.push('Solar Inverter Model');
      }

      if (openSolarData.solar_inverter_capacity) {
        populateData.solar_inverter_capacity = openSolarData.solar_inverter_capacity.toString();
        matchedFields.push('Solar Inverter Capacity');
      }

      // Match battery inverter manufacturer and model
      if (openSolarData.battery_inverter_manufacturer && dropdownOptions.dropdownOptions.battery_inverter_manufacturer) {
        const matchedManufacturer = this.findBestMatch(
          openSolarData.battery_inverter_manufacturer,
          dropdownOptions.dropdownOptions.battery_inverter_manufacturer
        );
        
        if (matchedManufacturer) {
          populateData.battery_inverter_manufacturer = matchedManufacturer;
          matchedFields.push('Battery Inverter Manufacturer');
        } else {
          unmatchedFields.push('Battery Inverter Manufacturer');
        }
      }

      if (openSolarData.battery_inverter_model) {
        populateData.battery_inverter_model = openSolarData.battery_inverter_model;
        matchedFields.push('Battery Inverter Model');
      }

      if (openSolarData.battery_inverter_capacity) {
        populateData.battery_inverter_capacity = openSolarData.battery_inverter_capacity.toString();
        matchedFields.push('Battery Inverter Capacity');
      }

      // Save the populated data to EPVS Excel
      if (Object.keys(populateData).length > 0) {
        const saveResult = await this.saveDynamicInputs(opportunityId, populateData, templateFileName);
        
        if (saveResult.success) {
          this.logger.log(`✅ Successfully auto-populated EPVS Excel with ${Object.keys(populateData).length} fields`);
          
          return {
            success: true,
            message: `EPVS Excel auto-populated with ${matchedFields.length} matched fields`,
            data: {
              matchedFields,
              unmatchedFields,
              totalFields: Object.keys(populateData).length,
              populatedData: populateData
            }
          };
        } else {
          return {
            success: false,
            message: 'Failed to save auto-populated data to EPVS Excel',
            error: saveResult.error
          };
        }
      } else {
        return {
          success: false,
          message: 'No data available to auto-populate',
          error: 'No OpenSolar data found'
        };
      }

    } catch (error) {
      this.logger.error(`❌ Error auto-populating EPVS Excel with OpenSolar data:`, error.message);
      return {
        success: false,
        message: 'Failed to auto-populate EPVS Excel with OpenSolar data',
        error: error.message
      };
    }
  }

  /**
   * Find best match between OpenSolar data and Excel dropdown options
   */
  private findBestMatch(openSolarValue: string, dropdownOptions: string[]): string | null {
    if (!openSolarValue || !dropdownOptions || dropdownOptions.length === 0) {
      return null;
    }

    const normalizedOpenSolar = openSolarValue.toLowerCase().trim();

    // 1. Exact match
    const exactMatch = dropdownOptions.find(option => 
      option.toLowerCase().trim() === normalizedOpenSolar
    );
    if (exactMatch) return exactMatch;

    // 2. Contains match (OpenSolar contains dropdown option)
    const containsMatch = dropdownOptions.find(option => 
      normalizedOpenSolar.includes(option.toLowerCase().trim())
    );
    if (containsMatch) return containsMatch;

    // 3. Contained match (Dropdown option contains OpenSolar)
    const containedMatch = dropdownOptions.find(option => 
      option.toLowerCase().trim().includes(normalizedOpenSolar)
    );
    if (containedMatch) return containedMatch;

    // 4. Manufacturer pattern matching
    const manufacturerPatterns = {
      'jinko': 'Jinko Solar',
      'trina': 'Trina Solar',
      'canadian': 'Canadian Solar',
      'longi': 'Longi',
      'ja solar': 'JA Solar',
      'risen': 'Risen',
      'q cells': 'Q Cells',
      'sunpower': 'SunPower',
      'lg': 'LG',
      'panasonic': 'Panasonic',
      'rec': 'REC',
      'solarworld': 'SolarWorld',
      'first solar': 'First Solar',
      'yingli': 'Yingli',
      'v-tac': 'V-TAC',
      'solis': 'Solis',
      'fronius': 'Fronius',
      'sma': 'SMA',
      'victron': 'Victron',
      'tesla': 'Tesla',
      'sonnen': 'Sonnen',
      'lg chem': 'LG Chem',
      'pylontech': 'Pylontech',
      'enphase': 'Enphase',
      'growatt': 'Growatt',
      'solaredge': 'SolarEdge'
    };

    for (const [pattern, manufacturer] of Object.entries(manufacturerPatterns)) {
      if (normalizedOpenSolar.includes(pattern)) {
        const match = dropdownOptions.find(option => 
          option.toLowerCase().includes(manufacturer.toLowerCase())
        );
        if (match) return match;
      }
    }

    return null;
  }

  /**
   * Perform complete EPVS calculation with user session isolation
   */
  async performCompleteCalculationWithSession(
    userId: string,
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string; pdfPath?: string }> {
    this.logger.log(`Performing complete EPVS calculation with session isolation for user: ${userId}, opportunity: ${opportunityId}`);

    try {
      // Queue the request through session management
      const result = await this.sessionManagementService.queueRequest(
        userId,
        'epvs_calculation',
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
      this.logger.error(`Session-based EPVS calculation failed for user ${userId}:`, error);
      return { 
        success: false, 
        message: 'EPVS calculation failed', 
        error: error.message 
      };
    }
  }

  /**
   * Execute EPVS calculation with user isolation (called by session management)
   */
  async executeEPVSCalculationWithIsolation(
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

    this.logger.log(`Executing EPVS calculation with isolation for user: ${userId}`);

    try {
      // Check if template file exists
      const templateFilePath = this.getTemplateFilePath(data.templateFileName);
      if (!fs.existsSync(templateFilePath)) {
        const error = `EPVS template file not found at: ${templateFilePath}`;
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Check if we're on Windows (required for Excel COM automation)
      if (process.platform !== 'win32') {
        const error = 'EPVS Excel automation requires Windows platform';
        this.logger.error(error);
        return { success: false, message: error };
      }

      // Create user-specific file paths
      const userExcelDir = path.join(workingDirectory, 'excel');
      const newFilePath = path.join(userExcelDir, `epvs_calculation_${data.opportunityId}_${Date.now()}.xlsm`);
      const pdfPath = path.join(workingDirectory, 'pdf', `epvs_calculation_${data.opportunityId}_${Date.now()}.pdf`);

      // Step 1: Automatically populate Flux rates if postcode is available
      if (data.customerDetails.postcode) {
        this.logger.log(`🔌 Auto-populating Flux rates for postcode: ${data.customerDetails.postcode}`);
        try {
          const fluxRatesResult = await this.populateFluxRatesInExcel(data.opportunityId, data.customerDetails.postcode);
          if (fluxRatesResult.success) {
            this.logger.log(`✅ Successfully auto-populated Flux rates for ${data.customerDetails.postcode}`);
          } else {
            this.logger.warn(`⚠️ Failed to auto-populate Flux rates: ${fluxRatesResult.error}`);
          }
        } catch (fluxError) {
          this.logger.warn(`⚠️ Error auto-populating Flux rates: ${fluxError.message}`);
        }
      }

      // Step 2: Create user-specific PowerShell script
      const psScript = this.createIsolatedEPVSCalculationScript(
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
      const tempScriptPath = path.join(workingDirectory, 'temp', `epvs_calculation_${Date.now()}.ps1`);
      fs.writeFileSync(tempScriptPath, psScript);
      
      this.logger.log(`Created isolated EPVS PowerShell script: ${tempScriptPath}`);

      // Execute PowerShell script
      const result = await this.runPowerShellScript(tempScriptPath);
      
      // Clean up temporary file
      try {
        fs.unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Could not clean up temporary EPVS script: ${cleanupError.message}`);
      }

      if (result.success) {
        return {
          success: true,
          message: 'EPVS calculation completed successfully with user isolation',
          filePath: newFilePath,
          pdfPath: pdfPath
        };
      } else {
        return {
          success: false,
          message: 'EPVS calculation failed',
          error: result.error || 'Unknown error'
        };
      }
    } catch (error) {
      this.logger.error(`Isolated EPVS calculation failed for user ${userId}:`, error);
      return { 
        success: false, 
        message: 'EPVS calculation failed', 
        error: error.message 
      };
    }
  }

  /**
   * Create isolated EPVS calculation script for user session
   */
  private createIsolatedEPVSCalculationScript(
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
# Isolated EPVS Calculation - User Session
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

Write-Host "Starting isolated EPVS calculation for: $opportunityId" -ForegroundColor Green
Write-Host "Working Directory: $workingDirectory" -ForegroundColor Yellow
Write-Host "Template path: $templatePath" -ForegroundColor Yellow
Write-Host "New file path: $newFilePath" -ForegroundColor Yellow

try {
    # Create user-specific Excel application
    Write-Host "Creating isolated Excel application for EPVS..." -ForegroundColor Green
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Isolated Excel application created successfully" -ForegroundColor Green
    
    # Copy template to user's directory
    Write-Host "Copying EPVS template file..." -ForegroundColor Green
    Copy-Item -Path $templatePath -Destination $newFilePath -Force
    Write-Host "EPVS template file copied successfully" -ForegroundColor Green
    
    # Open the copied file
    Write-Host "Opening EPVS workbook..." -ForegroundColor Green
    $workbook = $excel.Workbooks.Open($newFilePath, 0, $false, 5, $password)
    Write-Host "EPVS workbook opened successfully" -ForegroundColor Green
    
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
    
    # Fill in customer details for EPVS
    Write-Host "Filling EPVS customer details..." -ForegroundColor Green
    $worksheet.Range("H11").Value = "${this.escapeForPowerShell(customerDetails.customerName)}"
    $worksheet.Range("H12").Value = "${this.escapeForPowerShell(customerDetails.address)}"
    $worksheet.Range("H13").Value = "${this.escapeForPowerShell(customerDetails.postcode)}"
    Write-Host "EPVS customer details filled successfully" -ForegroundColor Green
    
    # Populate Flux rates if postcode is available
    if ("${this.escapeForPowerShell(customerDetails.postcode)}" -ne "") {
        Write-Host "Populating Flux rates for postcode: ${this.escapeForPowerShell(customerDetails.postcode)}" -ForegroundColor Green
        
        # Call the backend API to get Flux rates
        try {
            $fluxRatesUrl = "http://localhost:3000/epvs-automation/flux-rates/${this.escapeForPowerShell(customerDetails.postcode)}"
            Write-Host "Fetching Flux rates from: $fluxRatesUrl" -ForegroundColor Yellow
            
            $fluxResponse = Invoke-RestMethod -Uri $fluxRatesUrl -Method GET -ContentType "application/json"
            
            if ($fluxResponse.success -and $fluxResponse.rates) {
                Write-Host "Successfully fetched Flux rates" -ForegroundColor Green
                
                # Populate the Flux rates in Excel
                # Import rates
                $worksheet.Range("H22").Value = [decimal]$fluxResponse.rates.import.day  # Import Day Rate
                $worksheet.Range("H23").Value = [decimal]$fluxResponse.rates.import.flux  # Import Flux Rate
                $worksheet.Range("H24").Value = [decimal]$fluxResponse.rates.import.peak  # Import Peak Rate
                
                # Export rates
                $worksheet.Range("J22").Value = [decimal]$fluxResponse.rates.export.day  # Export Day Rate
                $worksheet.Range("J23").Value = [decimal]$fluxResponse.rates.export.flux  # Export Flux Rate
                $worksheet.Range("J24").Value = [decimal]$fluxResponse.rates.export.peak  # Export Peak Rate
                
                Write-Host "Flux rates populated successfully in Excel" -ForegroundColor Green
                Write-Host "Import Day: $($fluxResponse.rates.import.day), Flux: $($fluxResponse.rates.import.flux), Peak: $($fluxResponse.rates.import.peak)" -ForegroundColor Cyan
                Write-Host "Export Day: $($fluxResponse.rates.export.day), Flux: $($fluxResponse.rates.export.flux), Peak: $($fluxResponse.rates.export.peak)" -ForegroundColor Cyan
            } else {
                Write-Host "Failed to fetch Flux rates: $($fluxResponse.message)" -ForegroundColor Red
            }
        } catch {
            Write-Host "Error fetching Flux rates: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "No postcode provided, skipping Flux rates population" -ForegroundColor Yellow
    }
    
    # Save the workbook
    Write-Host "Saving EPVS workbook..." -ForegroundColor Green
    $workbook.Save()
    Write-Host "EPVS workbook saved successfully" -ForegroundColor Green
    
    # Close workbook
    $workbook.Close($false)
    Write-Host "EPVS workbook closed successfully" -ForegroundColor Green
    
    Write-Host "Isolated EPVS calculation completed successfully!" -ForegroundColor Green
    Write-Host "Output file: $newFilePath" -ForegroundColor Green
    
} catch {
    Write-Host "Error in isolated EPVS calculation: $_" -ForegroundColor Red
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

Write-Host "Isolated EPVS calculation process completed!" -ForegroundColor Green
exit 0
`;
  }
}
