import { Injectable, Logger } from '@nestjs/common';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as fs from 'fs';
import * as path from 'path';

const execAsync = promisify(exec);

@Injectable()
export class EnhancedFileManagerService {
  private readonly logger = new Logger(EnhancedFileManagerService.name);

  /**
   * Enhanced file access with retry mechanism and permission checking
   */
  async ensureFileAccess(filePath: string, maxRetries: number = 3): Promise<boolean> {
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        this.logger.log(`Attempt ${attempt}/${maxRetries}: Checking file access for ${filePath}`);
        
        // Check if file exists
        if (!fs.existsSync(filePath)) {
          throw new Error(`File does not exist: ${filePath}`);
        }

        // Check file permissions
        await this.checkFilePermissions(filePath);
        
        // Try to open file for read-write access
        await this.testFileAccess(filePath);
        
        this.logger.log(`✅ File access successful for ${filePath}`);
        return true;
        
      } catch (error) {
        this.logger.warn(`❌ Attempt ${attempt} failed for ${filePath}: ${error.message}`);
        
        if (attempt === maxRetries) {
          this.logger.error(`❌ All attempts failed for ${filePath}`);
          return false;
        }
        
        // Wait before retry with exponential backoff
        const waitTime = Math.pow(2, attempt) * 1000;
        this.logger.log(`⏳ Waiting ${waitTime}ms before retry...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      }
    }
    
    return false;
  }

  /**
   * Check file permissions and ownership
   */
  private async checkFilePermissions(filePath: string): Promise<void> {
    try {
      // Get file stats
      const stats = fs.statSync(filePath);
      this.logger.log(`File stats: size=${stats.size}, modified=${stats.mtime}`);
      
      // Check if file is readable
      fs.accessSync(filePath, fs.constants.R_OK);
      this.logger.log(`✅ File is readable`);
      
      // Check if file is writable
      fs.accessSync(filePath, fs.constants.W_OK);
      this.logger.log(`✅ File is writable`);
      
    } catch (error) {
      throw new Error(`Permission check failed: ${error.message}`);
    }
  }

  /**
   * Test file access by attempting to open and close it
   */
  private async testFileAccess(filePath: string): Promise<void> {
    try {
      // Try to open file for reading
      const fd = fs.openSync(filePath, 'r+');
      fs.closeSync(fd);
      this.logger.log(`✅ File access test successful`);
    } catch (error) {
      throw new Error(`File access test failed: ${error.message}`);
    }
  }

  /**
   * Enhanced Excel process cleanup with comprehensive error handling
   */
  async forceCleanupExcelProcesses(): Promise<void> {
    this.logger.log('🧹 Starting comprehensive Excel process cleanup...');
    
    try {
      // Step 1: Kill Excel processes gracefully
      await this.killExcelProcessesGracefully();
      
      // Step 2: Force kill any remaining processes
      await this.forceKillExcelProcesses();
      
      // Step 3: Clean up COM objects and temporary files
      await this.cleanupComObjects();
      
      // Step 4: Wait for system to stabilize
      await new Promise(resolve => setTimeout(resolve, 3000));
      
      this.logger.log('✅ Excel process cleanup completed');
      
    } catch (error) {
      this.logger.error(`❌ Error during Excel cleanup: ${error.message}`);
      throw error;
    }
  }

  /**
   * Kill Excel processes gracefully
   */
  private async killExcelProcessesGracefully(): Promise<void> {
    try {
      this.logger.log('🔄 Attempting graceful Excel process termination...');
      
      const command = `
        Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | 
        ForEach-Object { 
          try { 
            $_.CloseMainWindow() 
            Start-Sleep -Milliseconds 500
            if (!$_.HasExited) { 
              $_.Kill() 
            }
          } catch { 
            Write-Host "Process $($_.Id) already terminated" 
          }
        }
      `;
      
      await execAsync(`powershell.exe -Command "${command}"`);
      this.logger.log('✅ Graceful Excel termination completed');
      
    } catch (error) {
      this.logger.warn(`⚠️ Graceful termination failed: ${error.message}`);
    }
  }

  /**
   * Force kill Excel processes
   */
  private async forceKillExcelProcesses(): Promise<void> {
    try {
      this.logger.log('💥 Force killing remaining Excel processes...');
      
      const command = `
        Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | 
        Stop-Process -Force -ErrorAction SilentlyContinue
      `;
      
      await execAsync(`powershell.exe -Command "${command}"`);
      this.logger.log('✅ Force kill completed');
      
    } catch (error) {
      this.logger.warn(`⚠️ Force kill failed: ${error.message}`);
    }
  }

  /**
   * Clean up COM objects and temporary files
   */
  private async cleanupComObjects(): Promise<void> {
    try {
      this.logger.log('🧽 Cleaning up COM objects and temporary files...');
      
      const command = `
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        # Clean up temporary Excel files
        $tempPath = $env:TEMP
        Get-ChildItem -Path $tempPath -Filter "~$*.xls*" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $tempPath -Filter "*.tmp" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        
        Write-Host "COM cleanup completed"
      `;
      
      await execAsync(`powershell.exe -Command "${command}"`);
      this.logger.log('✅ COM cleanup completed');
      
    } catch (error) {
      this.logger.warn(`⚠️ COM cleanup failed: ${error.message}`);
    }
  }

  /**
   * Create a safe copy of Excel file with proper error handling
   */
  async createSafeFileCopy(sourcePath: string, targetPath: string): Promise<boolean> {
    this.logger.log(`📋 Creating safe copy: ${sourcePath} -> ${targetPath}`);
    
    try {
      // Ensure source file is accessible
      if (!await this.ensureFileAccess(sourcePath)) {
        throw new Error(`Source file not accessible: ${sourcePath}`);
      }
      
      // Ensure target directory exists
      const targetDir = path.dirname(targetPath);
      if (!fs.existsSync(targetDir)) {
        fs.mkdirSync(targetDir, { recursive: true });
        this.logger.log(`📁 Created target directory: ${targetDir}`);
      }
      
      // Remove target file if it exists
      if (fs.existsSync(targetPath)) {
        fs.unlinkSync(targetPath);
        this.logger.log(`🗑️ Removed existing target file`);
      }
      
      // Copy file with retry mechanism
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          fs.copyFileSync(sourcePath, targetPath);
          this.logger.log(`✅ File copy successful on attempt ${attempt}`);
          
          // Verify copy
          if (fs.existsSync(targetPath)) {
            const sourceSize = fs.statSync(sourcePath).size;
            const targetSize = fs.statSync(targetPath).size;
            
            if (sourceSize === targetSize) {
              this.logger.log(`✅ Copy verification successful (${targetSize} bytes)`);
              return true;
            } else {
              throw new Error(`Size mismatch: source=${sourceSize}, target=${targetSize}`);
            }
          } else {
            throw new Error('Target file was not created');
          }
          
        } catch (error) {
          this.logger.warn(`❌ Copy attempt ${attempt} failed: ${error.message}`);
          
          if (attempt === 3) {
            throw error;
          }
          
          // Wait before retry
          await new Promise(resolve => setTimeout(resolve, 1000 * attempt));
        }
      }
      
      return false;
      
    } catch (error) {
      this.logger.error(`❌ Safe file copy failed: ${error.message}`);
      return false;
    }
  }

  /**
   * Enhanced Excel file opening with comprehensive error handling
   */
  async openExcelFileSafely(filePath: string, password?: string): Promise<{ success: boolean; error?: string; retry?: boolean }> {
    this.logger.log(`📂 Opening Excel file safely: ${filePath}`);
    
    try {
      // Pre-flight checks
      if (!await this.ensureFileAccess(filePath)) {
        return { success: false, error: 'File access check failed', retry: true };
      }
      
      // Clean up any existing Excel processes
      await this.forceCleanupExcelProcesses();
      
      // Create PowerShell script for safe Excel opening
      const script = this.createSafeExcelOpenScript(filePath, password);
      const scriptPath = path.join(process.cwd(), 'temp', `safe_excel_open_${Date.now()}.ps1`);
      
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
        
        this.logger.log(`✅ Excel file opened successfully`);
        return { success: true };
        
      } finally {
        // Clean up script file
        if (fs.existsSync(scriptPath)) {
          fs.unlinkSync(scriptPath);
        }
      }
      
    } catch (error) {
      this.logger.error(`❌ Safe Excel opening failed: ${error.message}`);
      return { success: false, error: error.message, retry: true };
    }
  }

  /**
   * Create PowerShell script for safe Excel opening
   */
  private createSafeExcelOpenScript(filePath: string, password?: string): string {
    const escapedPath = filePath.replace(/\\/g, '\\\\');
    const passwordParam = password ? `"${password}"` : '$null';
    
    return `
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${escapedPath}"
$password = ${passwordParam}

Write-Host "Starting safe Excel file opening..." -ForegroundColor Green
Write-Host "File: $filePath" -ForegroundColor Yellow

try {
    # Final cleanup before opening
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open workbook with error handling
    try {
        if ($password) {
            $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
            Write-Host "Workbook opened with password" -ForegroundColor Green
        } else {
            $workbook = $excel.Workbooks.Open($filePath)
            Write-Host "Workbook opened without password" -ForegroundColor Green
        }
    } catch {
        Write-Host "Failed to open workbook: $_" -ForegroundColor Red
        throw "Could not open workbook: $_"
    }
    
    # Verify workbook is accessible
    if ($workbook) {
        Write-Host "Workbook verification successful" -ForegroundColor Green
        Write-Host "Worksheets count: $($workbook.Worksheets.Count)" -ForegroundColor Cyan
    } else {
        throw "Workbook object is null"
    }
    
    # Keep workbook open for operations
    Write-Host "Excel file ready for operations" -ForegroundColor Green
    
} catch {
    Write-Host "Error opening Excel file: $_" -ForegroundColor Red
    throw $_
} finally {
    # Note: We don't close Excel here as it needs to stay open for operations
    # The calling code should handle cleanup
}
`;
  }

  /**
   * Get detailed file information for debugging
   */
  async getFileInfo(filePath: string): Promise<any> {
    try {
      if (!fs.existsSync(filePath)) {
        return { exists: false, error: 'File does not exist' };
      }
      
      const stats = fs.statSync(filePath);
      const permissions = {
        readable: false,
        writable: false,
        executable: false
      };
      
      try {
        fs.accessSync(filePath, fs.constants.R_OK);
        permissions.readable = true;
      } catch {}
      
      try {
        fs.accessSync(filePath, fs.constants.W_OK);
        permissions.writable = true;
      } catch {}
      
      try {
        fs.accessSync(filePath, fs.constants.X_OK);
        permissions.executable = true;
      } catch {}
      
      return {
        exists: true,
        path: filePath,
        size: stats.size,
        created: stats.birthtime,
        modified: stats.mtime,
        accessed: stats.atime,
        permissions,
        isFile: stats.isFile(),
        isDirectory: stats.isDirectory()
      };
      
    } catch (error) {
      return { exists: false, error: error.message };
    }
  }
}
