import { Injectable, Logger } from '@nestjs/common';
import { EnhancedFileManagerService } from './enhanced-file-manager.service';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as path from 'path';
import * as fs from 'fs';

const execAsync = promisify(exec);

@Injectable()
export class EnhancedExcelAutomationService {
  private readonly logger = new Logger(EnhancedExcelAutomationService.name);
  private readonly PASSWORD = 'CreativSolar2024!';
  private readonly OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
  private readonly EPVS_OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');

  /**
   * Get versioned file path (v1, v2, v3, etc.) for opportunity files
   * Only creates new versions when no file exists for the opportunity
   */
  private getVersionedFilePath(directory: string, baseFileName: string, extension: string): string {
    // Ensure directory exists
    if (!fs.existsSync(directory)) {
      fs.mkdirSync(directory, { recursive: true });
    }

    // First check if any file exists for this opportunity (versioned or not)
    const files = fs.readdirSync(directory);
    const basePattern = baseFileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // Escape special regex characters
    
    // Check for existing files (both versioned and non-versioned)
    for (const file of files) {
      // Check for versioned files: baseFileName-v1.xlsm, baseFileName-v2.xlsm, etc.
      const versionRegex = new RegExp(`^${basePattern}-v(\\d+)\\.${extension}$`);
      const versionMatch = file.match(versionRegex);
      if (versionMatch) {
        // File already exists, return the existing file path
        return path.join(directory, file);
      }
      
      // Check for non-versioned files: baseFileName.xlsm
      if (file === `${baseFileName}.${extension}`) {
        // File already exists, return the existing file path
        return path.join(directory, file);
      }
    }

    // No file exists, create the first version (v1)
    return path.join(directory, `${baseFileName}-v1.${extension}`);
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
    
    let maxVersion = 0;
    for (const file of files) {
      const match = file.match(versionRegex);
      if (match) {
        const version = parseInt(match[1], 10);
        if (version > maxVersion) {
          maxVersion = version;
        }
      }
    }

    // Return the next version
    const nextVersion = maxVersion + 1;
    this.logger.log(`🎯 Template selection: Creating new version v${nextVersion} (max existing: v${maxVersion})`);
    return path.join(directory, `${baseFileName}-v${nextVersion}.${extension}`);
  }

  constructor(
    private readonly fileManager: EnhancedFileManagerService
  ) {}

  /**
   * Enhanced opportunity file creation with comprehensive error handling
   */
  async createOpportunityFileEnhanced(
    opportunityId: string, 
    customerDetails: { customerName: string; address: string; postcode: string }, 
    templateFileName?: string
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string }> {
    
    this.logger.log(`🚀 Creating enhanced opportunity file for: ${opportunityId}`);
    
    try {
      // Step 1: Pre-flight checks and cleanup
      await this.fileManager.forceCleanupExcelProcesses();
      
      // Step 2: Determine template path
      const templatePath = this.getTemplateFilePath(templateFileName);
      this.logger.log(`📋 Using template: ${templatePath}`);
      
      // Step 3: Ensure template file is accessible
      if (!await this.fileManager.ensureFileAccess(templatePath)) {
        throw new Error(`Template file not accessible: ${templatePath}`);
      }
      
      // Step 4: Create target file path with versioning
      const newFilePath = this.getNewVersionedFilePath(this.OPPORTUNITIES_FOLDER, `Off peak V2.1 Eon SEG-${opportunityId}`, 'xlsm');
      this.logger.log(`🎯 Target file: ${newFilePath}`);
      
      // Step 5: Create safe copy of template
      if (!await this.fileManager.createSafeFileCopy(templatePath, newFilePath)) {
        throw new Error('Failed to create safe copy of template file');
      }
      
      // Step 6: Open and populate the copied file
      const openResult = await this.fileManager.openExcelFileSafely(newFilePath, this.PASSWORD);
      if (!openResult.success) {
        throw new Error(`Failed to open Excel file: ${openResult.error}`);
      }
      
      // Step 7: Populate customer details
      await this.populateCustomerDetails(newFilePath, customerDetails);
      
      this.logger.log(`✅ Enhanced opportunity file created successfully: ${newFilePath}`);
      
      return {
        success: true,
        message: `Successfully created opportunity file for ${opportunityId}`,
        filePath: newFilePath
      };
      
    } catch (error) {
      this.logger.error(`❌ Enhanced opportunity file creation failed: ${error.message}`);
      
      // Cleanup on failure
      await this.fileManager.forceCleanupExcelProcesses();
      
      return {
        success: false,
        message: `Failed to create opportunity file: ${error.message}`,
        error: error.message
      };
    }
  }

  /**
   * Enhanced complete calculation with retry mechanism
   */
  async performCompleteCalculationEnhanced(
    opportunityId: string,
    customerDetails: { customerName: string; address: string; postcode: string },
    radioButtonSelections: string[],
    dynamicInputs?: Record<string, string>,
    templateFileName?: string
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string; pdfPath?: string }> {
    
    this.logger.log(`🧮 Starting enhanced complete calculation for: ${opportunityId}`);
    
    const maxRetries = 3;
    let lastError: string = '';
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        this.logger.log(`🔄 Attempt ${attempt}/${maxRetries} for calculation`);
        
        // Step 1: Comprehensive cleanup
        await this.fileManager.forceCleanupExcelProcesses();
        
        // Step 2: Create opportunity file
        const createResult = await this.createOpportunityFileEnhanced(opportunityId, customerDetails, templateFileName);
        if (!createResult.success) {
          throw new Error(`Opportunity file creation failed: ${createResult.error}`);
        }
        
        // Step 3: Perform calculation
        const calculationResult = await this.performCalculationWithRetry(
          opportunityId,
          radioButtonSelections,
          dynamicInputs,
          createResult.filePath!
        );
        
        if (calculationResult.success) {
          this.logger.log(`✅ Enhanced calculation completed successfully on attempt ${attempt}`);
          return calculationResult;
        } else {
          throw new Error(`Calculation failed: ${calculationResult.error}`);
        }
        
      } catch (error) {
        lastError = error.message;
        this.logger.warn(`❌ Attempt ${attempt} failed: ${lastError}`);
        
        if (attempt < maxRetries) {
          const waitTime = Math.pow(2, attempt) * 2000; // Exponential backoff
          this.logger.log(`⏳ Waiting ${waitTime}ms before retry...`);
          await new Promise(resolve => setTimeout(resolve, waitTime));
        }
      }
    }
    
    this.logger.error(`❌ All ${maxRetries} attempts failed. Last error: ${lastError}`);
    
    return {
      success: false,
      message: `Failed to complete calculation after ${maxRetries} attempts`,
      error: lastError
    };
  }

  /**
   * Perform calculation with retry mechanism
   */
  private async performCalculationWithRetry(
    opportunityId: string,
    radioButtonSelections: string[],
    dynamicInputs: Record<string, string> | undefined,
    filePath: string
  ): Promise<{ success: boolean; message: string; error?: string; filePath?: string; pdfPath?: string }> {
    
    try {
      // Create enhanced calculation script
      const script = this.createEnhancedCalculationScript(
        opportunityId,
        radioButtonSelections,
        dynamicInputs,
        filePath
      );
      
      const scriptPath = path.join(process.cwd(), 'temp', `enhanced_calculation_${Date.now()}.ps1`);
      
      // Ensure temp directory exists
      const tempDir = path.dirname(scriptPath);
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }
      
      fs.writeFileSync(scriptPath, script);
      
      try {
        // Execute the script
        const { stdout, stderr } = await execAsync(`powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`);
        
        if (stderr) {
          this.logger.warn(`PowerShell stderr: ${stderr}`);
        }
        
        // Parse results from stdout
        const result = this.parseCalculationResults(stdout);
        
        if (result.success) {
          return {
            success: true,
            message: `Successfully completed calculation for ${opportunityId}`,
            filePath: result.filePath,
            pdfPath: result.pdfPath
          };
        } else {
          throw new Error(result.error || 'Unknown calculation error');
        }
        
      } finally {
        // Clean up script file
        if (fs.existsSync(scriptPath)) {
          fs.unlinkSync(scriptPath);
        }
      }
      
    } catch (error) {
      this.logger.error(`❌ Calculation execution failed: ${error.message}`);
      return {
        success: false,
        message: `Calculation execution failed: ${error.message}`,
        error: error.message
      };
    }
  }

  /**
   * Create enhanced calculation script with comprehensive error handling
   */
  private createEnhancedCalculationScript(
    opportunityId: string,
    radioButtonSelections: string[],
    dynamicInputs: Record<string, string> | undefined,
    filePath: string
  ): string {
    
    const escapedPath = filePath.replace(/\\/g, '\\\\');
    const radioButtonsString = radioButtonSelections.map(shape => `"${shape}"`).join(', ');
    let inputsString = '';
    
    if (dynamicInputs && Object.keys(dynamicInputs).length > 0) {
      inputsString = Object.entries(dynamicInputs)
        .map(([key, value]) => `    "${key}" = "${value.replace(/"/g, '\\"')}"`)
        .join('\n');
    }
    
    return `
$ErrorActionPreference = "Stop"

# Enhanced Calculation Script with Comprehensive Error Handling
$opportunityId = "${opportunityId}"
$filePath = "${escapedPath}"
$password = "${this.PASSWORD}"

# Radio button selections
$radioButtonSelections = @(${radioButtonsString})

# Dynamic inputs
$dynamicInputs = @{
${inputsString}
}

Write-Host "🚀 Starting enhanced calculation for: $opportunityId" -ForegroundColor Green
Write-Host "📁 File path: $filePath" -ForegroundColor Yellow
Write-Host "🔘 Radio buttons: $($radioButtonSelections.Count)" -ForegroundColor Cyan
Write-Host "📝 Dynamic inputs: $($dynamicInputs.Count)" -ForegroundColor Cyan

try {
    # Step 1: Comprehensive cleanup
    Write-Host "🧹 Step 1: Comprehensive cleanup..." -ForegroundColor Green
    
    # Kill all Excel processes
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    
    # Clean up COM objects
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    Write-Host "✅ Cleanup completed" -ForegroundColor Green
    
    # Step 2: Verify file access
    Write-Host "🔍 Step 2: Verifying file access..." -ForegroundColor Green
    
    if (-not (Test-Path $filePath)) {
        throw "File does not exist: $filePath"
    }
    
    # Test file permissions
    try {
        $fileInfo = Get-Item $filePath
        Write-Host "✅ File accessible: $($fileInfo.Length) bytes" -ForegroundColor Green
    } catch {
        throw "File access denied: $filePath"
    }
    
    # Step 3: Create Excel application
    Write-Host "📊 Step 3: Creating Excel application..." -ForegroundColor Green
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1
    
    Write-Host "✅ Excel application created" -ForegroundColor Green
    
    # Step 4: Open workbook with retry mechanism
    Write-Host "📂 Step 4: Opening workbook..." -ForegroundColor Green
    
    $workbook = $null
    $maxOpenAttempts = 3
    
    for ($attempt = 1; $attempt -le $maxOpenAttempts; $attempt++) {
        try {
            Write-Host "🔄 Open attempt $attempt/$maxOpenAttempts" -ForegroundColor Yellow
            
            if ($attempt -eq 1) {
                # Try with password first
                $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
            } else {
                # Try without password
                $workbook = $excel.Workbooks.Open($filePath)
            }
            
            Write-Host "✅ Workbook opened successfully" -ForegroundColor Green
            break
            
        } catch {
            Write-Host "❌ Open attempt $attempt failed: $_" -ForegroundColor Red
            
            if ($attempt -eq $maxOpenAttempts) {
                throw "Failed to open workbook after $maxOpenAttempts attempts: $_"
            }
            
            Start-Sleep -Seconds 2
        }
    }
    
    # Step 5: Process worksheets
    Write-Host "📋 Step 5: Processing worksheets..." -ForegroundColor Green
    
    $inputsSheet = $workbook.Worksheets.Item("Inputs")
    if (-not $inputsSheet) {
        throw "Inputs worksheet not found"
    }
    
    Write-Host "✅ Found Inputs worksheet" -ForegroundColor Green
    
    # Unprotect all worksheets
    foreach ($ws in $workbook.Worksheets) {
        try {
            if ($ws.ProtectContents) {
                $ws.Unprotect($password)
                Write-Host "🔓 Unprotected worksheet: $($ws.Name)" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "⚠️ Could not unprotect worksheet: $($ws.Name)" -ForegroundColor Yellow
        }
    }
    
    # Step 6: Set radio button selections
    Write-Host "🔘 Step 6: Setting radio button selections..." -ForegroundColor Green
    
    foreach ($selection in $radioButtonSelections) {
        try {
            $shape = $inputsSheet.Shapes.Item($selection)
            if ($shape) {
                $shape.ControlFormat.Value = 1
                Write-Host "✅ Set radio button: $selection" -ForegroundColor Green
            } else {
                Write-Host "⚠️ Radio button not found: $selection" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "⚠️ Error setting radio button $selection : $_" -ForegroundColor Yellow
        }
    }
    
    # Step 7: Set dynamic inputs
    Write-Host "📝 Step 7: Setting dynamic inputs..." -ForegroundColor Green
    
    foreach ($input in $dynamicInputs.GetEnumerator()) {
        try {
            $cell = $inputsSheet.Range($input.Key)
            if ($cell) {
                $cell.Value = $input.Value
                Write-Host "✅ Set input $($input.Key) = $($input.Value)" -ForegroundColor Green
            }
        } catch {
            Write-Host "⚠️ Error setting input $($input.Key): $_" -ForegroundColor Yellow
        }
    }
    
    # Step 8: Save workbook
    Write-Host "💾 Step 8: Saving workbook..." -ForegroundColor Green
    
    $workbook.Save()
    Write-Host "✅ Workbook saved successfully" -ForegroundColor Green
    
    # Step 9: Generate PDF if possible
    Write-Host "📄 Step 9: Generating PDF..." -ForegroundColor Green
    
    $pdfPath = $null
    try {
        $pdfPath = $filePath -replace '\\.xlsm$', '.pdf'
        $workbook.ExportAsFixedFormat(0, $pdfPath, 0, $false, $false, $false, $false, $false)
        Write-Host "✅ PDF generated: $pdfPath" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ PDF generation failed: $_" -ForegroundColor Yellow
        $pdfPath = $null
    }
    
    # Step 10: Output results
    Write-Host "📊 Step 10: Calculation completed successfully!" -ForegroundColor Green
    Write-Host "RESULT:SUCCESS" -ForegroundColor Green
    Write-Host "FILE_PATH:$filePath" -ForegroundColor Green
    if ($pdfPath) {
        Write-Host "PDF_PATH:$pdfPath" -ForegroundColor Green
    }
    
} catch {
    Write-Host "❌ Critical error in calculation: $_" -ForegroundColor Red
    Write-Host "RESULT:ERROR" -ForegroundColor Red
    Write-Host "ERROR:$_" -ForegroundColor Red
    throw $_
} finally {
    Write-Host "🧹 Starting cleanup process..." -ForegroundColor Yellow
    
    # Cleanup workbook
    if ($workbook) {
        try {
            $workbook.Close($false)
            Write-Host "✅ Workbook closed" -ForegroundColor Green
        } catch {
            Write-Host "⚠️ Error closing workbook: $_" -ForegroundColor Yellow
        }
    }
    
    # Cleanup Excel
    if ($excel) {
        try {
            $excel.Quit()
            Write-Host "✅ Excel closed" -ForegroundColor Green
        } catch {
            Write-Host "⚠️ Error closing Excel: $_" -ForegroundColor Yellow
        }
    }
    
    # Release COM objects
    try {
        if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "✅ COM objects released" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Error releasing COM objects: $_" -ForegroundColor Yellow
    }
    
    # Final cleanup
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    
    Write-Host "✅ Cleanup completed" -ForegroundColor Green
}

Write-Host "🏁 Enhanced calculation script completed" -ForegroundColor Green
`;
  }

  /**
   * Parse calculation results from PowerShell output
   */
  private parseCalculationResults(output: string): { success: boolean; filePath?: string; pdfPath?: string; error?: string } {
    const lines = output.split('\n');
    let success = false;
    let filePath: string | undefined;
    let pdfPath: string | undefined;
    let error: string | undefined;
    
    for (const line of lines) {
      if (line.includes('RESULT:SUCCESS')) {
        success = true;
      } else if (line.includes('RESULT:ERROR')) {
        success = false;
      } else if (line.includes('FILE_PATH:')) {
        filePath = line.split('FILE_PATH:')[1]?.trim();
      } else if (line.includes('PDF_PATH:')) {
        pdfPath = line.split('PDF_PATH:')[1]?.trim();
      } else if (line.includes('ERROR:')) {
        error = line.split('ERROR:')[1]?.trim();
      }
    }
    
    return { success, filePath, pdfPath, error };
  }

  /**
   * Populate customer details in Excel file
   */
  private async populateCustomerDetails(filePath: string, customerDetails: { customerName: string; address: string; postcode: string }): Promise<void> {
    this.logger.log(`👤 Populating customer details in: ${filePath}`);
    
    const script = `
$ErrorActionPreference = "Stop"

$filePath = "${filePath.replace(/\\/g, '\\\\')}"
$password = "${this.PASSWORD}"

$customerName = "${customerDetails.customerName.replace(/"/g, '\\"')}"
$address = "${customerDetails.address.replace(/"/g, '\\"')}"
$postcode = "${customerDetails.postcode.replace(/"/g, '\\"')}"

Write-Host "Populating customer details..." -ForegroundColor Green

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1
    
    $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
    $worksheet = $workbook.Worksheets.Item("Inputs")
    
    # Unprotect worksheet
    if ($worksheet.ProtectContents) {
        $worksheet.Unprotect($password)
    }
    
    # Set customer details
    $worksheet.Range("H12").Value = $customerName
    $worksheet.Range("H13").Value = $address
    $worksheet.Range("H14").Value = $postcode
    
    $workbook.Save()
    Write-Host "Customer details populated successfully" -ForegroundColor Green
    
} catch {
    Write-Host "Error populating customer details: $_" -ForegroundColor Red
    throw $_
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    
    # Release COM objects
    if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;
    
    const scriptPath = path.join(process.cwd(), 'temp', `populate_customer_${Date.now()}.ps1`);
    
    try {
      fs.writeFileSync(scriptPath, script);
      await execAsync(`powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`);
      this.logger.log(`✅ Customer details populated successfully`);
    } finally {
      if (fs.existsSync(scriptPath)) {
        fs.unlinkSync(scriptPath);
      }
    }
  }

  /**
   * Get template file path
   */
  private getTemplateFilePath(templateFileName?: string): string {
    if (templateFileName) {
      return path.join(process.cwd(), 'src', 'excel-file-calculator', 'templates', templateFileName);
    }
    return path.join(process.cwd(), 'src', 'excel-file-calculator', 'templates', 'Off peak V2.1 Eon SEG Template.xlsm');
  }
}
