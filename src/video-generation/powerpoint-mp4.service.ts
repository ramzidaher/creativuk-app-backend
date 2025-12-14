import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import { promisify } from 'util';
import { exec } from 'child_process';

const execAsync = promisify(exec);

@Injectable()
export class PowerpointMp4Service {
  private readonly logger = new Logger(PowerpointMp4Service.name);
  private readonly outputDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'output');
  private readonly publicDir = path.join(process.cwd(), 'public', 'videos');

  constructor() {
    // Ensure public directory exists
    if (!fs.existsSync(this.publicDir)) {
      fs.mkdirSync(this.publicDir, { recursive: true });
    }
  }

  /**
   * Convert PowerPoint presentation to MP4 using COM automation
   */
  async convertPptxToMp4(pptxPath: string, mp4Path: string): Promise<{ success: boolean; error?: string }> {
    try {
      this.logger.log(`Converting PowerPoint to MP4: ${pptxPath} -> ${mp4Path}`);

      // Check if PowerPoint file exists
      if (!fs.existsSync(pptxPath)) {
        throw new Error(`PowerPoint file not found: ${pptxPath}`);
      }

      // Create PowerShell script for PowerPoint COM automation
      const psScript = this.createPowerPointToMp4Script(pptxPath, mp4Path);
      const tempScriptPath = path.join(process.cwd(), `temp-pptx-to-mp4-${Date.now()}.ps1`);
      
      fs.writeFileSync(tempScriptPath, psScript);
      this.logger.log(`Created temporary PowerShell script: ${tempScriptPath}`);

      try {
        // Execute PowerShell script
        const { stdout, stderr } = await execAsync(`powershell -ExecutionPolicy Bypass -File "${tempScriptPath}"`, {
          timeout: 300000, // 5 minutes timeout
        });

        if (stderr && !stderr.includes('warning')) {
          this.logger.warn(`PowerShell stderr: ${stderr}`);
        }

        this.logger.log(`PowerShell stdout: ${stdout}`);

        // Check if MP4 was created successfully
        if (fs.existsSync(mp4Path)) {
          const stats = fs.statSync(mp4Path);
          this.logger.log(`MP4 created successfully: ${mp4Path} (${stats.size} bytes)`);
          return { success: true };
        } else {
          throw new Error(`MP4 file was not created at expected location: ${mp4Path}`);
        }

      } finally {
        // Clean up temporary script
        try {
          fs.unlinkSync(tempScriptPath);
          this.logger.log(`Cleaned up temporary PowerShell script: ${tempScriptPath}`);
        } catch (cleanupError) {
          this.logger.warn(`Failed to cleanup temporary script: ${cleanupError.message}`);
        }
      }

    } catch (error) {
      this.logger.error(`PowerPoint to MP4 conversion failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }

  /**
   * Create PowerShell script for PowerPoint COM automation
   */
  private createPowerPointToMp4Script(pptxPath: string, mp4Path: string): string {
    const pptxPathEscaped = pptxPath.replace(/\\/g, '\\\\');
    const mp4PathEscaped = mp4Path.replace(/\\/g, '\\\\');
    
    return `
# Convert PowerPoint to MP4 using COM automation
$ErrorActionPreference = "Stop"

# Configuration
$pptxFilePath = "${pptxPathEscaped}"
$mp4Path = "${mp4PathEscaped}"

Write-Host "Converting PowerPoint to MP4..." -ForegroundColor Green
Write-Host "PowerPoint file: $pptxFilePath" -ForegroundColor Yellow
Write-Host "MP4 output: $mp4Path" -ForegroundColor Yellow

# Validate paths
if (!(Test-Path $pptxFilePath)) {
    throw "PowerPoint file not found: $pptxFilePath"
}

# Clean and validate MP4 path
$mp4Path = [System.IO.Path]::GetFullPath($mp4Path)
$mp4Dir = Split-Path $mp4Path -Parent

# Ensure the MP4 directory exists and is writable
if (!(Test-Path $mp4Dir)) {
    try {
        New-Item -ItemType Directory -Path $mp4Dir -Force | Out-Null
        Write-Host "Created MP4 directory: $mp4Dir" -ForegroundColor Green
    } catch {
        throw "Failed to create MP4 directory: $mp4Dir - $_"
    }
}

# Test write permissions
try {
    $testFile = Join-Path $mp4Dir "test-write.tmp"
    "test" | Out-File -FilePath $testFile -Encoding ASCII
    Remove-Item $testFile -Force
    Write-Host "Write permissions verified for: $mp4Dir" -ForegroundColor Green
} catch {
    throw "No write permissions for MP4 directory: $mp4Dir - $_"
}

# Kill any existing PowerPoint processes to prevent conflicts
try {
    Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Host "Cleared any existing PowerPoint processes" -ForegroundColor Green
} catch {
    Write-Host "No existing PowerPoint processes to clear" -ForegroundColor Yellow
}

$ppt = $null
$presentation = $null

try {
    # Create PowerPoint application
    Write-Host "Creating PowerPoint application..." -ForegroundColor Yellow
    $ppt = New-Object -ComObject PowerPoint.Application
    
    # Configure PowerPoint settings for better video rendering
    $ppt.Visible = -1  # -1 = msoTrue
    $ppt.DisplayAlerts = "ppAlertsNone"
    
    # Set PowerPoint to normal window state for better rendering
    try {
        $ppt.WindowState = "ppWindowNormal"  # Changed from minimized to normal for better text rendering
    } catch {
        Write-Host "Warning: Could not set PowerPoint window state" -ForegroundColor Yellow
    }
    
    # Configure presentation settings for better video export
    try {
        # Set slide show settings for better video quality
        $ppt.SlideShowSettings.ShowType = "ppShowTypeSpeaker"  # Speaker mode for better rendering
        $ppt.SlideShowSettings.LoopUntilStopped = $false
        $ppt.SlideShowSettings.ShowWithNarration = $false
        $ppt.SlideShowSettings.ShowWithAnimation = $true  # Keep animations for better text rendering
    } catch {
        Write-Host "Warning: Could not configure slide show settings" -ForegroundColor Yellow
    }
    
    Write-Host "PowerPoint application created successfully" -ForegroundColor Green
    
    # Open presentation with better error handling
    Write-Host "Opening PowerPoint presentation..." -ForegroundColor Yellow
    Write-Host "File path: $pptxFilePath" -ForegroundColor Cyan
    Write-Host "File exists: $(Test-Path $pptxFilePath)" -ForegroundColor Cyan
    
    # Verify file is not corrupted before opening
    try {
        $fileInfo = Get-Item $pptxFilePath
        Write-Host "File size: $($fileInfo.Length) bytes" -ForegroundColor Cyan
        Write-Host "File last modified: $($fileInfo.LastWriteTime)" -ForegroundColor Cyan
        
        if ($fileInfo.Length -lt 1000) {
            throw "PowerPoint file appears to be corrupted or too small: $($fileInfo.Length) bytes"
        }
    } catch {
        throw "Cannot access PowerPoint file: $_"
    }
    
    try {
        # Open with different parameters for better compatibility
        $presentation = $ppt.Presentations.Open($pptxFilePath, $false, $false, $false)  # Not ReadOnly, Not Untitled, Not WithWindow
        Write-Host "Presentation opened successfully" -ForegroundColor Green
        
        # Wait for presentation to fully load and render
        Start-Sleep -Seconds 3
        
        # Configure presentation for better video export
        try {
            # Set presentation properties for better video rendering
            $presentation.SlideShowSettings.ShowType = "ppShowTypeSpeaker"
            $presentation.SlideShowSettings.LoopUntilStopped = $false
            $presentation.SlideShowSettings.ShowWithNarration = $false
            $presentation.SlideShowSettings.ShowWithAnimation = $true
            
            # Set slide timing for better video export
            $presentation.SlideShowSettings.AdvanceMode = "ppSlideShowManualAdvance"
            
            Write-Host "Presentation configured for video export" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not configure presentation settings: $_" -ForegroundColor Yellow
        }
        
        # Additional text rendering fixes for video export
        try {
            # Force a refresh of all slides to ensure proper text rendering
            Write-Host "Refreshing slide rendering for video export..." -ForegroundColor Yellow
            
            # Go through each slide and ensure proper text rendering
            for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
                $slide = $presentation.Slides.Item($i)
                
                # Force slide to render by accessing its properties
                $slide.Shapes.Count | Out-Null
                
                # Ensure text boxes are properly positioned for video export
                for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                    $shape = $slide.Shapes.Item($j)
                    if ($shape.HasTextFrame -eq -1) {  # -1 = msoTrue
                        # Force text frame to render properly
                        $shape.TextFrame.TextRange.Text | Out-Null
                    }
                }
                
                Start-Sleep -Milliseconds 200  # Give more time for rendering
            }
            Write-Host "Slide rendering refresh completed for video export" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not refresh slide rendering: $_" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "Failed to open presentation: $_" -ForegroundColor Red
        throw "Could not open PowerPoint presentation: $_"
    }
    
    # Export to MP4 using improved method for better text rendering
    Write-Host "Exporting to MP4 with enhanced text rendering..." -ForegroundColor Yellow
    
    # Wait a moment for all rendering to complete
    Start-Sleep -Seconds 2
    
    try {
        # First try ExportAsFixedFormat with MP4 format (better for text rendering)
        Write-Host "Attempting ExportAsFixedFormat method..." -ForegroundColor Cyan
        $presentation.ExportAsFixedFormat($mp4Path, 39)  # 39 = ppFixedFormatTypeMP4
        Write-Host "MP4 export initiated successfully using ExportAsFixedFormat" -ForegroundColor Green
    } catch {
        Write-Host "ExportAsFixedFormat method failed, trying SaveAs: $_" -ForegroundColor Yellow
        try {
            # Fallback to SaveAs method with MP4 format
            Write-Host "Attempting SaveAs method..." -ForegroundColor Cyan
            $presentation.SaveAs($mp4Path, 39)  # 39 = ppSaveAsMP4
            Write-Host "MP4 export initiated successfully using SaveAs method" -ForegroundColor Green
        } catch {
            Write-Host "All MP4 export methods failed: $_" -ForegroundColor Red
            throw "MP4 export failed: $_"
        }
    }
    
    # Wait for MP4 conversion to complete (PowerPoint SaveAs is asynchronous)
    Write-Host "Waiting for MP4 conversion to complete..." -ForegroundColor Yellow
    $timeoutSeconds = 300  # 5 minutes timeout
    $startTime = Get-Date
    $lastSize = 0
    $stableCount = 0
    $maxStableChecks = 3  # Require more stability checks for better reliability
    
    while ($true) {
        if ((Get-Date) -gt ($startTime).AddSeconds($timeoutSeconds)) {
            throw "MP4 conversion timed out after $timeoutSeconds seconds"
        }
        
        if (Test-Path $mp4Path) {
            $currentSize = (Get-Item $mp4Path).Length
            Write-Host "MP4 file size: $currentSize bytes" -ForegroundColor Cyan
            
            # Check if file size is stable (not changing for 15 seconds)
            if ($currentSize -eq $lastSize -and $currentSize -gt 0) {
                $stableCount++
                if ($stableCount -ge $maxStableChecks) {  # More stability checks for reliability
                    Write-Host "MP4 file size is stable, conversion appears complete" -ForegroundColor Green
                    break
                }
            } else {
                $stableCount = 0
                $lastSize = $currentSize
            }
        } else {
            Write-Host "MP4 file not yet created, waiting..." -ForegroundColor Cyan
        }
        
        Start-Sleep -Seconds 5
    }
    
    # Final verification with additional checks
    if (Test-Path $mp4Path) {
        $finalSize = (Get-Item $mp4Path).Length
        if ($finalSize -gt 0) {
            Write-Host "MP4 created successfully: $mp4Path (Size: $finalSize bytes)" -ForegroundColor Green
            
            # Additional verification - check if file is accessible and not corrupted
            try {
                $fileInfo = Get-Item $mp4Path
                if ($fileInfo.Length -gt 1000) {  # Ensure file is not just a header
                    Write-Host "MP4 file verification passed" -ForegroundColor Green
                } else {
                    Write-Host "Warning: MP4 file seems too small, may be corrupted" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "Warning: Could not verify MP4 file integrity" -ForegroundColor Yellow
            }
        } else {
            throw "MP4 file was created but is empty (0 bytes)"
        }
    } else {
        throw "MP4 file was not created at expected location: $mp4Path"
    }
    
} catch {
    Write-Host "Critical error in MP4 conversion: $_" -ForegroundColor Red
    throw $_
} finally {
    Write-Host "Starting cleanup process..." -ForegroundColor Yellow
    
    # Always try to close PowerPoint properly
    if ($presentation) {
        try {
            Write-Host "Closing presentation..." -ForegroundColor Yellow
            $presentation.Close()
        } catch {
            Write-Host "Warning: Error closing presentation: $_" -ForegroundColor Yellow
        }
    }
    
    if ($ppt) {
        try {
            Write-Host "Quitting PowerPoint..." -ForegroundColor Yellow
            $ppt.Quit()
        } catch {
            Write-Host "Warning: Error quitting PowerPoint: $_" -ForegroundColor Yellow
        }
        
        # Release COM object
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
            Write-Host "PowerPoint COM object released" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Error releasing PowerPoint COM object: $_" -ForegroundColor Yellow
        }
    }
    
    # Force garbage collection
    try {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Garbage collection completed" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error during garbage collection: $_" -ForegroundColor Yellow
    }
    
    # Final cleanup - kill any remaining PowerPoint processes
    try {
        Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-Host "Final PowerPoint process cleanup completed" -ForegroundColor Green
    } catch {
        Write-Host "No remaining PowerPoint processes to clean up" -ForegroundColor Yellow
    }
    
    Write-Host "Cleanup completed" -ForegroundColor Green
}

Write-Host "PowerPoint to MP4 conversion completed successfully!" -ForegroundColor Green
`;
  }

  /**
   * Generate MP4 from PowerPoint presentation
   */
  async generateMp4FromPptx(pptxPath: string, opportunityId: string, customerName: string): Promise<{
    success: boolean;
    videoPath?: string;
    publicUrl?: string;
    error?: string;
  }> {
    try {
      const timestamp = Date.now();
      const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const mp4Filename = `proposal_${safeCustomerName}_${opportunityId}_${timestamp}.mp4`;
      const mp4Path = path.join(this.publicDir, mp4Filename);

      // Convert PowerPoint to MP4
      const conversionResult = await this.convertPptxToMp4(pptxPath, mp4Path);
      
      if (!conversionResult.success) {
        return {
          success: false,
          error: conversionResult.error,
        };
      }

      // MP4 is already in the public directory, no need to copy

      const publicUrl = `/videos/${mp4Filename}`;
      
      this.logger.log(`MP4 generated successfully: ${mp4Path}`);
      this.logger.log(`Public URL: ${publicUrl}`);

      return {
        success: true,
        videoPath: mp4Path,
        publicUrl: publicUrl,
      };

    } catch (error) {
      this.logger.error(`MP4 generation failed: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }
}
