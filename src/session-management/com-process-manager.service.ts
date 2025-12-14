import { Injectable, Logger } from '@nestjs/common';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as path from 'path';
import * as fs from 'fs';

const execAsync = promisify(exec);

export interface ComProcessInfo {
  processId: number;
  userId: string;
  sessionId: string;
  application: 'excel' | 'powerpoint';
  startTime: Date;
  workingDirectory: string;
}

@Injectable()
export class ComProcessManagerService {
  private readonly logger = new Logger(ComProcessManagerService.name);
  private readonly activeProcesses = new Map<number, ComProcessInfo>();

  /**
   * Create isolated Excel COM process for user
   */
  async createExcelProcess(userId: string, sessionId: string, workingDirectory: string): Promise<number> {
    try {
      // Kill any existing Excel processes for this user first
      await this.killUserExcelProcesses(userId);

      // Create user-specific Excel script
      const scriptPath = path.join(workingDirectory, 'excel', `excel_${Date.now()}.ps1`);
      const scriptContent = this.generateIsolatedExcelScript(workingDirectory);
      
      fs.writeFileSync(scriptPath, scriptContent);

      // Start Excel process with user isolation
      const command = `powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`;
      const { stdout, stderr } = await execAsync(command);

      if (stderr) {
        this.logger.warn(`Excel process stderr for user ${userId}:`, stderr);
      }

      // Extract process ID from output
      const processId = this.extractProcessId(stdout);
      
      if (processId) {
        const processInfo: ComProcessInfo = {
          processId,
          userId,
          sessionId,
          application: 'excel',
          startTime: new Date(),
          workingDirectory
        };

        this.activeProcesses.set(processId, processInfo);
        this.logger.log(`Created Excel process ${processId} for user ${userId}`);
        return processId;
      } else {
        throw new Error('Failed to get Excel process ID');
      }
    } catch (error) {
      this.logger.error(`Failed to create Excel process for user ${userId}:`, error);
      throw error;
    }
  }

  /**
   * Create isolated PowerPoint COM process for user
   */
  async createPowerPointProcess(userId: string, sessionId: string, workingDirectory: string): Promise<number> {
    try {
      // Kill any existing PowerPoint processes for this user first
      await this.killUserPowerPointProcesses(userId);

      // Create user-specific PowerPoint script
      const scriptPath = path.join(workingDirectory, 'powerpoint', `powerpoint_${Date.now()}.ps1`);
      const scriptContent = this.generateIsolatedPowerPointScript(workingDirectory);
      
      fs.writeFileSync(scriptPath, scriptContent);

      // Start PowerPoint process with user isolation
      const command = `powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`;
      const { stdout, stderr } = await execAsync(command);

      if (stderr) {
        this.logger.warn(`PowerPoint process stderr for user ${userId}:`, stderr);
      }

      // Extract process ID from output
      const processId = this.extractProcessId(stdout);
      
      if (processId) {
        const processInfo: ComProcessInfo = {
          processId,
          userId,
          sessionId,
          application: 'powerpoint',
          startTime: new Date(),
          workingDirectory
        };

        this.activeProcesses.set(processId, processInfo);
        this.logger.log(`Created PowerPoint process ${processId} for user ${userId}`);
        return processId;
      } else {
        throw new Error('Failed to get PowerPoint process ID');
      }
    } catch (error) {
      this.logger.error(`Failed to create PowerPoint process for user ${userId}:`, error);
      throw error;
    }
  }

  /**
   * Kill specific COM process
   */
  async killProcess(processId: number): Promise<void> {
    try {
      const processInfo = this.activeProcesses.get(processId);
      if (processInfo) {
        await execAsync(`taskkill /PID ${processId} /F`);
        this.activeProcesses.delete(processId);
        this.logger.log(`Killed ${processInfo.application} process ${processId} for user ${processInfo.userId}`);
      }
    } catch (error) {
      this.logger.warn(`Failed to kill process ${processId}:`, error);
    }
  }

  /**
   * Kill all processes for a specific user
   */
  async killUserProcesses(userId: string): Promise<void> {
    const userProcesses = Array.from(this.activeProcesses.values())
      .filter(process => process.userId === userId);

    for (const process of userProcesses) {
      await this.killProcess(process.processId);
    }
  }

  /**
   * Kill all Excel processes for a specific user
   */
  async killUserExcelProcesses(userId: string): Promise<void> {
    const userExcelProcesses = Array.from(this.activeProcesses.values())
      .filter(process => process.userId === userId && process.application === 'excel');

    for (const process of userExcelProcesses) {
      await this.killProcess(process.processId);
    }
  }

  /**
   * Kill all PowerPoint processes for a specific user
   */
  async killUserPowerPointProcesses(userId: string): Promise<void> {
    const userPowerPointProcesses = Array.from(this.activeProcesses.values())
      .filter(process => process.userId === userId && process.application === 'powerpoint');

    for (const process of userPowerPointProcesses) {
      await this.killProcess(process.processId);
    }
  }

  /**
   * Get process information
   */
  getProcessInfo(processId: number): ComProcessInfo | null {
    return this.activeProcesses.get(processId) || null;
  }

  /**
   * Get all processes for a user
   */
  getUserProcesses(userId: string): ComProcessInfo[] {
    return Array.from(this.activeProcesses.values())
      .filter(process => process.userId === userId);
  }

  /**
   * Get all active processes
   */
  getAllProcesses(): ComProcessInfo[] {
    return Array.from(this.activeProcesses.values());
  }

  /**
   * Generate isolated Excel script
   */
  private generateIsolatedExcelScript(workingDirectory: string): string {
    return `# Isolated Excel COM Process
$ErrorActionPreference = "Stop"

# Configuration
$workingDirectory = "${workingDirectory.replace(/\\/g, '\\\\')}"
$excelProcessId = $PID

Write-Host "Starting isolated Excel process..." -ForegroundColor Green
Write-Host "Process ID: $excelProcessId" -ForegroundColor Yellow
Write-Host "Working Directory: $workingDirectory" -ForegroundColor Yellow

# Create Excel application with user isolation
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1  # Enable all macros
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    Write-Host "Process ID: $excelProcessId" -ForegroundColor Green
    
    # Keep the process alive and wait for commands
    Write-Host "Excel process ready for operations..." -ForegroundColor Green
    
    # Wait indefinitely for external commands
    while ($true) {
        Start-Sleep -Seconds 1
        
        # Check for stop signal
        $stopFile = Join-Path $workingDirectory "stop_excel.txt"
        if (Test-Path $stopFile) {
            Write-Host "Stop signal received, shutting down Excel..." -ForegroundColor Yellow
            Remove-Item $stopFile -Force
            break
        }
    }
    
} catch {
    Write-Host "Error in Excel process: $_" -ForegroundColor Red
    throw $_
} finally {
    Write-Host "Cleaning up Excel process..." -ForegroundColor Yellow
    
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
    
    Write-Host "Excel process cleanup completed" -ForegroundColor Green
}

Write-Host "Excel process terminated" -ForegroundColor Green
exit 0`;
  }

  /**
   * Generate isolated PowerPoint script
   */
  private generateIsolatedPowerPointScript(workingDirectory: string): string {
    return `# Isolated PowerPoint COM Process
$ErrorActionPreference = "Stop"

# Configuration
$workingDirectory = "${workingDirectory.replace(/\\/g, '\\\\')}"
$powerpointProcessId = $PID

Write-Host "Starting isolated PowerPoint process..." -ForegroundColor Green
Write-Host "Process ID: $powerpointProcessId" -ForegroundColor Yellow
Write-Host "Working Directory: $workingDirectory" -ForegroundColor Yellow

# Create PowerPoint application with user isolation
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $ppt.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
    
    # Minimize the PowerPoint window
    try {
        $ppt.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized
    } catch {
        Write-Host "Warning: Could not minimize PowerPoint window" -ForegroundColor Yellow
    }
    
    Write-Host "PowerPoint application created successfully" -ForegroundColor Green
    Write-Host "Process ID: $powerpointProcessId" -ForegroundColor Green
    
    # Keep the process alive and wait for commands
    Write-Host "PowerPoint process ready for operations..." -ForegroundColor Green
    
    # Wait indefinitely for external commands
    while ($true) {
        Start-Sleep -Seconds 1
        
        # Check for stop signal
        $stopFile = Join-Path $workingDirectory "stop_powerpoint.txt"
        if (Test-Path $stopFile) {
            Write-Host "Stop signal received, shutting down PowerPoint..." -ForegroundColor Yellow
            Remove-Item $stopFile -Force
            break
        }
    }
    
} catch {
    Write-Host "Error in PowerPoint process: $_" -ForegroundColor Red
    throw $_
} finally {
    Write-Host "Cleaning up PowerPoint process..." -ForegroundColor Yellow
    
    if ($ppt) {
        try {
            $ppt.Quit()
        } catch {
            Write-Host "Warning: Error quitting PowerPoint: $_" -ForegroundColor Yellow
        }
    }
    
    # Force release COM objects
    try {
        if ($ppt) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {
        Write-Host "Warning: Error releasing COM objects: $_" -ForegroundColor Yellow
    }
    
    Write-Host "PowerPoint process cleanup completed" -ForegroundColor Green
}

Write-Host "PowerPoint process terminated" -ForegroundColor Green
exit 0`;
  }

  /**
   * Extract process ID from PowerShell output
   */
  private extractProcessId(output: string): number | null {
    const match = output.match(/Process ID: (\d+)/);
    return match ? parseInt(match[1], 10) : null;
  }

  /**
   * Send stop signal to user's Excel process
   */
  async stopUserExcelProcess(userId: string): Promise<void> {
    const userProcesses = this.getUserProcesses(userId)
      .filter(process => process.application === 'excel');

    for (const process of userProcesses) {
      const stopFile = path.join(process.workingDirectory, 'stop_excel.txt');
      fs.writeFileSync(stopFile, 'stop');
    }
  }

  /**
   * Send stop signal to user's PowerPoint process
   */
  async stopUserPowerPointProcess(userId: string): Promise<void> {
    const userProcesses = this.getUserProcesses(userId)
      .filter(process => process.application === 'powerpoint');

    for (const process of userProcesses) {
      const stopFile = path.join(process.workingDirectory, 'stop_powerpoint.txt');
      fs.writeFileSync(stopFile, 'stop');
    }
  }
}

