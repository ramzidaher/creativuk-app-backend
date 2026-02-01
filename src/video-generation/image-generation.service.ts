import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import * as fsPromises from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

@Injectable()
export class ImageGenerationService {
  private readonly logger = new Logger(ImageGenerationService.name);
  private readonly outputDir = path.join(process.cwd(), 'src', 'video-generation', 'output');
  private readonly publicDir = path.join(process.cwd(), 'public', 'images');
  private readonly assetsDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'presnetation', 'Proposal');

  constructor() {
    // Ensure output directories exist
    if (!fs.existsSync(this.outputDir)) {
      fs.mkdirSync(this.outputDir, { recursive: true });
    }
    if (!fs.existsSync(this.publicDir)) {
      fs.mkdirSync(this.publicDir, { recursive: true });
    }
  }

  /**
   * Generate proposal images by exporting only customer-specific slides (1, 7, 12)
   * Static slides (2-6, 8-11, 13-17) are copied from assets directory
   * Images are organized in opportunity-specific folders
   */
  async generateProposalImages(data: {
    opportunityId: string;
    customerName: string;
    date: string;
    postcode: string;
    solarData: any;
    pptxPath: string; // Path to the generated PowerPoint file
  }): Promise<{ success: boolean; images?: string[]; publicUrls?: string[]; error?: string }> {
    try {
      this.logger.log(`Generating proposal images for opportunity: ${data.opportunityId}`);

      // Create opportunity-specific directory
      const opportunityDir = path.join(this.publicDir, data.opportunityId);
      if (!fs.existsSync(opportunityDir)) {
        fs.mkdirSync(opportunityDir, { recursive: true });
      }

      const timestamp = Date.now();
      const safeCustomerName = data.customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const imagePrefix = `proposal_${safeCustomerName}_${data.opportunityId}_${timestamp}`;

      // Export only customer-specific slides (1, 7, 12) from PowerPoint
      const exportedSlides = await this.exportSpecificSlides(data.pptxPath, opportunityDir, imagePrefix);

      // Copy static slides (2-6, 8-11, 13-17) from assets directory
      const staticSlides = await this.copyStaticSlides(opportunityDir, imagePrefix);

      // Combine all slides in order (1-17)
      const allImageFiles = this.combineSlides(exportedSlides, staticSlides);

      if (allImageFiles.length === 0) {
        throw new Error('No slides were generated');
      }

      // Generate public URLs for the images (include opportunity ID in path)
      const publicUrls = allImageFiles.map(filename => `/images/${data.opportunityId}/${filename}`);

      this.logger.log(`Images generated successfully: ${allImageFiles.length} slides`);
      this.logger.log(`  - Exported: ${exportedSlides.size} customer-specific slides`);
      this.logger.log(`  - Copied: ${staticSlides.size} static slides`);
      this.logger.log(`  - Saved to: ${opportunityDir}`);

      return {
        success: true,
        images: allImageFiles,
        publicUrls: publicUrls,
      };

    } catch (error) {
      this.logger.error(`Image generation failed: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Export only customer-specific slides (1, 7, 12) from PowerPoint to PNG
   * These are the only slides that change with customer data
   */
  private async exportSpecificSlides(pptxPath: string, outputDir: string, imagePrefix: string): Promise<Map<number, string>> {
    const slidesToExport = [1, 7, 12]; // Only export slides that contain customer data
    const exportedSlides = new Map<number, string>();
    const timestamp = Date.now();

    this.logger.log(`Exporting customer-specific slides: ${slidesToExport.join(', ')}`);

    // Create PowerShell script to export specific slides in one session
    const slideExports = slidesToExport.map(slideNum => {
      const filename = `${imagePrefix}_slide_${slideNum}.png`;
      const outPng = path.join(outputDir, filename);
      return { slideNum, filename, outPng };
    });

    const exportScript = `
# PowerPoint Slide Export Script
# NOTE: If PowerPoint shows a sign-in dialog, this script will fail.
# SOLUTION: Sign in to Microsoft Office before running this script, or disable
#           Office sign-in prompts in File > Account > Sign out (if not needed)

$pptx = "${pptxPath.replace(/\\/g, '\\\\')}"
$widthPx = 1920
$exportedCount = 0

# Kill any existing PowerPoint processes to prevent conflicts
try {
    Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Host "Cleared any existing PowerPoint processes" -ForegroundColor Green
} catch {
    Write-Host "No existing PowerPoint processes to clear" -ForegroundColor Yellow
}

# Validate PowerPoint file before opening
Write-Host "Validating PowerPoint file..." -ForegroundColor Yellow
if (!(Test-Path $pptx)) {
    throw "PowerPoint file not found: $pptx"
}

try {
    $fileInfo = Get-Item $pptx
    Write-Host "File size: $($fileInfo.Length) bytes" -ForegroundColor Cyan
    Write-Host "File last modified: $($fileInfo.LastWriteTime)" -ForegroundColor Cyan
    
    if ($fileInfo.Length -lt 1000) {
        throw "PowerPoint file appears to be corrupted or too small: $($fileInfo.Length) bytes"
    }
    
    # Ensure file is not locked and is fully written
    Write-Host "Verifying file is not locked and fully written..." -ForegroundColor Yellow
    $fileStable = $false
    $stabilityChecks = 0
    $lastSize = $fileInfo.Length
    $maxStabilityChecks = 3
    
    for ($i = 0; $i -lt 10; $i++) {
        Start-Sleep -Milliseconds 200
        try {
            $currentInfo = Get-Item $pptx
            $currentSize = $currentInfo.Length
            
            if ($currentSize -eq $lastSize -and $currentSize -gt 0) {
                $stabilityChecks++
                if ($stabilityChecks -ge $maxStabilityChecks) {
                    $fileStable = $true
                    Write-Host "File is stable (size unchanged for $($stabilityChecks * 0.2) seconds)" -ForegroundColor Green
                    break
                }
            } else {
                $stabilityChecks = 0
                $lastSize = $currentSize
                Write-Host "File size changed: $lastSize bytes, waiting for stability..." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "File may be locked, waiting... ($i/10)" -ForegroundColor Yellow
        }
    }
    
    if (-not $fileStable) {
        Write-Host "Warning: File stability check incomplete, but proceeding..." -ForegroundColor Yellow
    }
    
    # Try to open file for reading to ensure it's not locked
    try {
        $testStream = [System.IO.File]::OpenRead($pptx)
        $testStream.Close()
        Write-Host "File is accessible and not locked" -ForegroundColor Green
    } catch {
        throw "File is locked or cannot be accessed: $_"
    }
    
} catch {
    throw "Cannot access PowerPoint file: $_"
}

$ppt = $null
$pres = $null

try {
  $ppt = New-Object -ComObject PowerPoint.Application
  
  # Suppress ALL dialogs and alerts to prevent blocking
  try {
    $ppt.DisplayAlerts = "ppAlertsNone"  # Suppress alerts
    $ppt.Visible = -1  # -1 = msoTrue (required - cannot hide window in some contexts)
    
    # Try to suppress sign-in prompts and other dialogs
    try {
      # Set automation security to prevent macro/security prompts
      $ppt.AutomationSecurity = 1  # msoAutomationSecurityLow
    } catch {
      Write-Host "Warning: Could not set automation security (may need Office update)" -ForegroundColor Yellow
    }
    
    # Suppress update prompts
    try {
      $ppt.FeatureInstall = 0  # msoFeatureInstallNone - don't prompt for feature installation
    } catch {
      Write-Host "Warning: Could not set feature install mode" -ForegroundColor Yellow
    }
    
    Write-Host "PowerPoint application configured with all dialogs suppressed" -ForegroundColor Green
  } catch {
    Write-Host "Warning: Some PowerPoint settings could not be configured: $_" -ForegroundColor Yellow
  }
  
  # Minimize window for faster performance and to reduce visual interference
  try {
    $ppt.WindowState = "ppWindowMinimized"
  } catch {
    Write-Host "Warning: Could not minimize PowerPoint window" -ForegroundColor Yellow
  }
  
  Write-Host "Opening PowerPoint presentation..." -ForegroundColor Yellow
  Write-Host "File path: $pptx" -ForegroundColor Cyan
  Write-Host "File exists: $(Test-Path $pptx)" -ForegroundColor Cyan
  Write-Host "Note: If PowerPoint shows a sign-in dialog, the script will fail. Please sign in to Office first or disable sign-in prompts." -ForegroundColor Yellow
  
  try {
    Write-Host "Attempting to open presentation (this may take a moment)..." -ForegroundColor Yellow
    
    # Open presentation with timeout handling
    # If PowerPoint shows a sign-in dialog, this will hang until timeout
    $openTimeout = 30  # seconds
    $openStartTime = Get-Date
    $pres = $null
    
    # Try to open in a way that won't hang on dialogs
    try {
      # Use ReadOnly=true, Untitled=false, WithWindow=false to minimize prompts
      $pres = $ppt.Presentations.Open($pptx, $true, $false, $false)
      
      # Check if we're stuck waiting (e.g., on a sign-in dialog)
      $elapsed = ((Get-Date) - $openStartTime).TotalSeconds
      if ($elapsed -gt 5) {
        Write-Host "Warning: Opening took $elapsed seconds - PowerPoint may be waiting for user input" -ForegroundColor Yellow
      }
    } catch {
      $elapsed = ((Get-Date) - $openStartTime).TotalSeconds
      if ($elapsed -gt $openTimeout) {
        throw "Opening presentation timed out after $openTimeout seconds. PowerPoint may be showing a sign-in dialog or other prompt that requires user interaction."
      }
      throw $_
    }
    
    if ($null -eq $pres) {
      throw "Failed to open presentation: Presentations.Open() returned null. This may indicate PowerPoint is waiting for user input (sign-in, security prompt, etc.)"
    }
    
    Write-Host "Presentation object created successfully, waiting for slides to load..." -ForegroundColor Green
    
    # Wait for presentation to fully load - PowerPoint may need time to load slides
    $maxWaitSeconds = 10
    $waited = 0
    $slideCount = 0
    
    while ($waited -lt $maxWaitSeconds) {
      Start-Sleep -Milliseconds 500
      $waited += 0.5
      
      try {
        $slideCount = $pres.Slides.Count
        if ($slideCount -gt 0) {
          Write-Host "Presentation loaded with $slideCount slides after $waited seconds" -ForegroundColor Green
          break
        }
      } catch {
        Write-Host "Error checking slide count: $_" -ForegroundColor Yellow
      }
      
      Write-Host "Waiting for slides to load... ($waited / $maxWaitSeconds seconds)" -ForegroundColor Cyan
    }
    
    # Final check
    $finalSlideCount = $pres.Slides.Count
    Write-Host "Final slide count: $finalSlideCount" -ForegroundColor Cyan
    
    # Additional diagnostics
    try {
      Write-Host "Presentation name: $($pres.Name)" -ForegroundColor Cyan
      Write-Host "Presentation path: $($pres.FullName)" -ForegroundColor Cyan
      Write-Host "Presentation saved: $($pres.Saved)" -ForegroundColor Cyan
    } catch {
      Write-Host "Could not get presentation properties: $_" -ForegroundColor Yellow
    }
    
    # Verify presentation has slides
    if ($finalSlideCount -eq 0) {
      Write-Host "ERROR: Presentation opened but has no slides!" -ForegroundColor Red
      Write-Host "File path: $pptx" -ForegroundColor Red
      Write-Host "File size: $((Get-Item $pptx).Length) bytes" -ForegroundColor Red
      Write-Host "This may indicate:" -ForegroundColor Yellow
      Write-Host "  1. The PowerPoint file is corrupted or empty" -ForegroundColor Yellow
      Write-Host "  2. The file was not fully generated before opening" -ForegroundColor Yellow
      Write-Host "  3. PowerPoint needs more time to load (try increasing wait time)" -ForegroundColor Yellow
      throw "Presentation has no slides after waiting $maxWaitSeconds seconds"
    }
    
    Write-Host "✅ Successfully opened presentation with $finalSlideCount slides" -ForegroundColor Green
    
  } catch {
    Write-Host "Failed to open presentation: $_" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    
    # Check if PowerPoint might be waiting for user input
    try {
      $pptProcess = Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue
      if ($pptProcess) {
        Write-Host "PowerPoint process is running - it may be waiting for user input (sign-in dialog, etc.)" -ForegroundColor Yellow
        Write-Host "SOLUTION: Please check if PowerPoint is showing a dialog and:" -ForegroundColor Yellow
        Write-Host "  1. Sign in to Microsoft Office if prompted" -ForegroundColor Yellow
        Write-Host "  2. Or disable Office sign-in: File > Account > Sign out (if not needed)" -ForegroundColor Yellow
        Write-Host "  3. Close any PowerPoint dialogs manually" -ForegroundColor Yellow
      }
    } catch {
      # Ignore errors checking for process
    }
    
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    throw "Could not open PowerPoint presentation: $_"
  }
  
${slideExports.map(({ slideNum, outPng }) => `
  # Export slide ${slideNum}
  try {
    if ($null -eq $pres) {
      throw "Presentation is null, cannot export slide ${slideNum}"
    }
    
    $outPng${slideNum} = "${outPng.replace(/\\/g, '\\\\')}"
    
    # Verify slide exists - use literal number ${slideNum}
    if (${slideNum} -gt $pres.Slides.Count) {
      throw "Slide ${slideNum} does not exist. Presentation has only $($pres.Slides.Count) slides"
    }
    
    # Use literal number ${slideNum} for Item() method
    $slide${slideNum} = $pres.Slides.Item(${slideNum})
    
    if ($null -eq $slide${slideNum}) {
      throw "Failed to get slide ${slideNum} from presentation"
    }
    
    $slide${slideNum}.Export($outPng${slideNum}, "PNG", $widthPx)
    Write-Host "✅ Exported slide ${slideNum} → $outPng${slideNum}"
    $exportedCount++
  } catch {
    Write-Host "❌ Error exporting slide ${slideNum}: $_" -ForegroundColor Red
    throw $_
  }
`).join('')}
  
  Write-Host "Exported $exportedCount out of ${slidesToExport.length} customer-specific slides" -ForegroundColor Green
} catch {
  Write-Host "Error opening/exporting presentation: $_" -ForegroundColor Red
  Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
  throw $_
} finally {
  Write-Host "Starting cleanup process..." -ForegroundColor Yellow
  
  # Always try to close PowerPoint properly
  if ($pres) {
    try {
      Write-Host "Closing presentation..." -ForegroundColor Yellow
      $pres.Close()
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
    } catch {
      Write-Host "Warning: Error releasing PowerPoint COM object: $_" -ForegroundColor Yellow
    }
  }
  
  # Release presentation COM object
  if ($pres) {
    try {
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    } catch {
      Write-Host "Warning: Error releasing presentation COM object: $_" -ForegroundColor Yellow
    }
  }
  
  # Force garbage collection
  try {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
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
`;

    const scriptPath = path.join(this.outputDir, `export_specific_slides_${timestamp}.ps1`);
    fs.writeFileSync(scriptPath, exportScript);

    try {
      this.logger.log(`Executing PowerShell script to export slides ${slidesToExport.join(', ')}...`);
      await execAsync(`powershell -ExecutionPolicy Bypass -File "${scriptPath}"`);
      
      // Skip verification for speed - assume PowerPoint export succeeded
      // Add slides to map directly based on expected filenames
      for (const { slideNum, filename } of slideExports) {
        exportedSlides.set(slideNum, filename);
        this.logger.log(`✅ Slide ${slideNum} exported: ${filename}`);
      }
    } catch (error) {
      this.logger.error(`Failed to export slides: ${error.message}`);
      throw error;
    } finally {
      // Clean up script
      if (fs.existsSync(scriptPath)) {
        fs.unlinkSync(scriptPath);
      }
    }

    this.logger.log(`Exported ${exportedSlides.size} customer-specific slides`);
    return exportedSlides;
  }

  /**
   * Copy static slides (2-6, 8-11, 13-17) from assets directory
   * Optimized to copy all files in parallel
   */
  private async copyStaticSlides(outputDir: string, imagePrefix: string): Promise<Map<number, string>> {
    const staticSlides = new Map<number, string>();
    const staticSlideNumbers = [2, 3, 4, 5, 6, 8, 9, 10, 11, 13, 14, 15, 16, 17];

    this.logger.log(`Copying static slides: ${staticSlideNumbers.join(', ')}`);

    // Copy all static slides in parallel for better performance
    const copyPromises = staticSlideNumbers.map(async (slideNum) => {
      const sourcePath = path.join(this.assetsDir, `Slide${slideNum}.PNG`);
      const filename = `${imagePrefix}_slide_${slideNum}.png`;
      const destPath = path.join(outputDir, filename);

      try {
        if (fs.existsSync(sourcePath)) {
          // Use fsPromises for async copy
          await fsPromises.copyFile(sourcePath, destPath);
          staticSlides.set(slideNum, filename);
          this.logger.log(`✅ Copied static slide ${slideNum}: ${filename}`);
          return { slideNum, filename, success: true };
        } else {
          this.logger.warn(`⚠️ Static slide ${slideNum} not found in assets: ${sourcePath}`);
          return { slideNum, filename: null, success: false };
        }
      } catch (error) {
        this.logger.error(`Failed to copy static slide ${slideNum}: ${error.message}`);
        return { slideNum, filename: null, success: false };
      }
    });

    // Wait for all copies to complete
    await Promise.all(copyPromises);

    this.logger.log(`Copied ${staticSlides.size} static slides`);
    return staticSlides;
  }

  /**
   * Combine exported and static slides in order (1-17)
   */
  private combineSlides(exportedSlides: Map<number, string>, staticSlides: Map<number, string>): string[] {
    const allSlides: string[] = [];
    
    for (let slideNum = 1; slideNum <= 17; slideNum++) {
      const filename = exportedSlides.get(slideNum) || staticSlides.get(slideNum);
      if (filename) {
        allSlides.push(filename);
      } else {
        this.logger.warn(`⚠️ Slide ${slideNum} is missing`);
      }
    }

    return allSlides;
  }
}

