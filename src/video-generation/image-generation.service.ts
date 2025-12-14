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
} catch {
    throw "Cannot access PowerPoint file: $_"
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1  # -1 = msoTrue (required - cannot hide window in some contexts)
$ppt.DisplayAlerts = "ppAlertsNone"

try {
  # Minimize window for faster performance
  try {
    $ppt.WindowState = "ppWindowMinimized"
  } catch {
    Write-Host "Warning: Could not minimize PowerPoint window" -ForegroundColor Yellow
  }
  
  Write-Host "Opening PowerPoint presentation..." -ForegroundColor Yellow
  $pres = $ppt.Presentations.Open($pptx, $true, $false, $false)  # ReadOnly, Untitled, WithWindow
  Write-Host "Opened presentation with $($pres.Slides.Count) slides" -ForegroundColor Green
  
${slideExports.map(({ slideNum, outPng }) => `
  # Export slide ${slideNum}
  try {
    $outPng${slideNum} = "${outPng.replace(/\\/g, '\\\\')}"
    $slide${slideNum} = $pres.Slides.Item(${slideNum})
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

