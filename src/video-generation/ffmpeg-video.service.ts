import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

@Injectable()
export class FfmpegVideoService {
  private readonly logger = new Logger(FfmpegVideoService.name);
  private readonly outputDir = path.join(process.cwd(), 'src', 'video-generation', 'output');
  private readonly publicDir = path.join(process.cwd(), 'public', 'videos');
  private readonly assetsDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'presnetation', 'Proposal');
  private readonly templatePptxPath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'presnetation', 'Proposal.pptx');

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
   * Generate proposal video using optimized FFmpeg pipeline
   */
  async generateProposalVideo(data: {
    opportunityId: string;
    customerName: string;
    date: string;
    postcode: string;
    solarData: any;
    pptxPath: string; // Path to the generated PowerPoint file
  }): Promise<{ success: boolean; videoPath?: string; publicUrl?: string; error?: string }> {
    try {
      this.logger.log(`Generating proposal video for opportunity: ${data.opportunityId}`);

      const timestamp = Date.now();
      const safeCustomerName = data.customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const videoFilename = `proposal_${safeCustomerName}_${data.opportunityId}_${timestamp}.mp4`;
      const videoPath = path.join(this.outputDir, videoFilename);
      const publicVideoPath = path.join(this.publicDir, videoFilename);

      // Use the existing PowerPoint template directly - no need to export individual slides
      // Just create a simple video from the existing template with video overlays
      await this.createOptimizedVideo(videoPath, data.pptxPath);

      // Copy to public directory for serving
      fs.copyFileSync(videoPath, publicVideoPath);

      const publicUrl = `/videos/${videoFilename}`;
      
      this.logger.log(`Video generated successfully: ${videoPath}`);
      this.logger.log(`Public URL: ${publicUrl}`);

      return {
        success: true,
        videoPath: videoPath,
        publicUrl: publicUrl,
      };

    } catch (error) {
      this.logger.error(`Video generation failed: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Create optimized video with all slides - much faster approach
   */
  private async createOptimizedVideo(outputPath: string, pptxPath: string): Promise<void> {
    this.logger.log('Creating optimized video with all slides...');
    
    // First, we need to export all slides from the PowerPoint
    await this.exportSpecificSlides(pptxPath);
    
    // Check if we have the required assets
    const mediaFile = path.join(this.assetsDir, 'Media1.mp4');
    
    if (!fs.existsSync(mediaFile)) {
      throw new Error('Media file not found. Please ensure Media1.mp4 exists in the assets directory.');
    }

    // Create a complete video with ALL 17 slides
    // Customer-specific slides (1, 7, 12) are generated fresh
    // Static slides (2-6, 8-11, 13-17) use existing files
    const inputs: string[] = [];
    const filters: string[] = [];
    
    // Add ALL slides (1-17)
    for (let i = 1; i <= 17; i++) {
      if (i === 10) {
        // Use the pre-made Slide10_with_video.mp4 file
        const slide10VideoPath = path.join(this.assetsDir, `Slide10_with_video.mp4`);
        if (fs.existsSync(slide10VideoPath)) {
          inputs.push(`-i "${slide10VideoPath}"`);
        } else {
          throw new Error('Slide10_with_video.mp4 not found. Please ensure this file exists in the assets directory.');
        }
      } else {
        // Use PNG slides (generated or existing)
        const slidePath = path.join(this.assetsDir, `Slide${i}.PNG`);
        if (fs.existsSync(slidePath)) {
          inputs.push(`-loop 1 -i "${slidePath}"`);
        } else {
          this.logger.warn(`Slide ${i} not found, skipping...`);
        }
      }
    }
    
    // Add media file for slide 12 overlay
    inputs.push(`-i "${mediaFile}"`);
    
    // Create filters for each slide
    for (let i = 1; i <= 17; i++) {
      const slidePath = path.join(this.assetsDir, `Slide${i}.PNG`);
      if (!fs.existsSync(slidePath) && i !== 10) {
        continue; // Skip missing slides
      }
      
      if (i === 10) {
        // Slide 10: Use pre-made video, trim to 3 seconds
        filters.push(`[${i-1}:v]scale=1920:1080:flags=fast_bilinear,setsar=1:1,trim=duration=3,setsar=1:1[s${i}]`);
      } else if (i === 12) {
        // Slide 12: 30 seconds with video overlay
        if (fs.existsSync(slidePath)) {
          filters.push(`[${i-1}:v]scale=1920:1080:flags=fast_bilinear,setsar=1:1,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:black[bg${i}];[${inputs.length-1}:v]scale=trunc(1920*0.38/2)*2:-2:flags=fast_bilinear,setsar=1:1[vid${i}];[bg${i}][vid${i}]overlay=W-w-10:H-h-150:shortest=1,trim=duration=30,setsar=1:1[s${i}]`);
        }
      } else {
        // Regular slides: 3 seconds each
        if (fs.existsSync(slidePath)) {
          filters.push(`[${i-1}:v]scale=1920:1080:flags=fast_bilinear,setsar=1:1,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:black,trim=duration=3,setsar=1:1[s${i}]`);
        }
      }
    }
    
    // Concatenate all available slides
    const availableSlides = Array.from({length: 17}, (_, i) => i + 1).filter(i => {
      if (i === 10) return true; // Slide10_with_video.mp4 exists
      return fs.existsSync(path.join(this.assetsDir, `Slide${i}.PNG`));
    });
    
    const concatInputs = availableSlides.map(slideNum => `[s${slideNum}]`).join('');
    filters.push(`${concatInputs}concat=n=${availableSlides.length}:v=1:a=0,setsar=1:1[out]`);
    
    const ffmpegCommand = `ffmpeg -y ${inputs.join(' ')} -filter_complex "${filters.join(';')}" -map "[out]" -c:v libx264 -preset ultrafast -crf 23 -pix_fmt yuv420p -movflags +faststart -r 30 "${outputPath}"`;

    this.logger.log(`Executing optimized FFmpeg command for ${availableSlides.length} slides (including all static slides)...`);
    await execAsync(ffmpegCommand);
    
    // Verify output was created
    if (!fs.existsSync(outputPath)) {
      throw new Error('Optimized video was not created');
    }
  }

  /**
   * Export specific slides (1, 7, 12) from PowerPoint to PNG
   * These are the only slides that change with customer data
   */
  private async exportAllSlides(pptxPath: string): Promise<void> {
    const slidesToExport = [1, 7, 12]; // Only export slides that change
    
    for (const slideNum of slidesToExport) {
      const outPng = path.join(this.assetsDir, `Slide${slideNum}.PNG`);
      
      // Always export these slides since they contain dynamic customer data
      this.logger.log(`Exporting slide ${slideNum} (contains customer data)...`);
      
      const exportScript = `
$pptx = "\${pptxPath.replace(/\\/g, '\\\\')}"
$outPng = "${outPng.replace(/\\/g, '\\\\')}"
$slideNum = ${slideNum}
$widthPx = 1920

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

try {
  $pres = $ppt.Presentations.Open($pptx, $true, $false, $false)
  $slide = $pres.Slides.Item($slideNum)
  $slide.Export($outPng, "PNG", $widthPx)
} finally {
  if ($pres) { $pres.Close() }
  if ($ppt) { $ppt.Quit() }
  foreach ($o in @($slide,$pres,$ppt)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "Exported slide $slideNum → $outPng"
`;

      const scriptPath = path.join(this.outputDir, `export_slide_${slideNum}.ps1`);
      fs.writeFileSync(scriptPath, exportScript);

      this.logger.log(`Exporting slide ${slideNum}...`);
      await execAsync(`powershell -ExecutionPolicy Bypass -File "${scriptPath}"`);
      
      // Clean up script
      fs.unlinkSync(scriptPath);
    }
  }

  /**
   * Export specific slides (1, 7, 12) from PowerPoint to PNG
   */
  private async exportSpecificSlides(pptxPath: string): Promise<void> {
    const slidesToExport = [1, 7, 12];
    
    for (const slideNum of slidesToExport) {
      const outPng = path.join(this.assetsDir, `Slide${slideNum}.PNG`);
      
      // ALWAYS export these slides - they contain customer-specific data
      // Never skip - each customer has different data in these slides
      this.logger.log(`Exporting slide ${slideNum} (contains customer data - never skip)...`);
      
      const exportScript = `
$pptx = "${pptxPath.replace(/\\/g, '\\\\')}"
$outPng = "${outPng.replace(/\\/g, '\\\\')}"
$slideNum = ${slideNum}
$widthPx = 1920

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

try {
  $pres = $ppt.Presentations.Open($pptx, $true, $false, $false)
  $slide = $pres.Slides.Item($slideNum)
  $slide.Export($outPng, "PNG", $widthPx)
} finally {
  if ($pres) { $pres.Close() }
  if ($ppt) { $ppt.Quit() }
  foreach ($o in @($slide,$pres,$ppt)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "Exported slide $slideNum → $outPng"
`;

      const scriptPath = path.join(this.outputDir, `export_slide_${slideNum}.ps1`);
      fs.writeFileSync(scriptPath, exportScript);

      this.logger.log(`Executing PowerShell script for slide ${slideNum}...`);
      await execAsync(`powershell -ExecutionPolicy Bypass -File "${scriptPath}"`);
      
      // Clean up script
      fs.unlinkSync(scriptPath);
    }
  }

  /**
   * Create slide 10 with video overlay
   */
  private async createSlide10WithVideo(): Promise<string> {
    const slide10Png = path.join(this.assetsDir, 'Slide10.PNG');
    const mediaFile = path.join(this.assetsDir, 'Media1.mp4');
    const outputPath = path.join(this.assetsDir, 'Slide10_with_video.mp4');

    // Check if required files exist
    if (!fs.existsSync(slide10Png)) {
      throw new Error(`Slide 10 PNG not found: ${slide10Png}`);
    }
    if (!fs.existsSync(mediaFile)) {
      throw new Error(`Media file not found: ${mediaFile}`);
    }

    const ffmpegCommand = `ffmpeg -loop 1 -i "${slide10Png}" -i "${mediaFile}" -filter_complex "[0:v]scale=1920:-1:flags=fast_bilinear[bg];[1:v]scale=trunc(1920*0.5/2)*2:-2:flags=fast_bilinear[vid];[bg][vid]overlay=W-w-80:H-h-120:shortest=1" -t 3 -an -c:v libx264 -preset veryfast -crf 18 -pix_fmt yuv420p -movflags +faststart "${outputPath}"`;

    this.logger.log('Creating slide 10 with video overlay...');
    await execAsync(ffmpegCommand);

    // Verify output was created
    if (!fs.existsSync(outputPath)) {
      throw new Error('Slide 10 video was not created');
    }

    return outputPath;
  }

  /**
   * Create slide 12 with video overlay
   */
  private async createSlide12WithVideo(): Promise<string> {
    const slide12Png = path.join(this.assetsDir, 'Slide12.PNG');
    const mediaFile = path.join(this.assetsDir, 'Media1.mp4');
    const outputPath = path.join(this.assetsDir, 'Slide12_with_video_full.mp4');

    // Check if required files exist
    if (!fs.existsSync(slide12Png)) {
      throw new Error(`Slide 12 PNG not found: ${slide12Png}`);
    }
    if (!fs.existsSync(mediaFile)) {
      throw new Error(`Media file not found: ${mediaFile}`);
    }

    const ffmpegCommand = `ffmpeg -loop 1 -i "${slide12Png}" -i "${mediaFile}" -filter_complex "[0:v]scale=1920:-1:flags=fast_bilinear[bg];[1:v]scale=trunc(1920*0.38/2)*2:-2:flags=fast_bilinear[vid];[bg][vid]overlay=W-w-10:H-h-150:shortest=1" -an -c:v libx264 -preset veryfast -crf 18 -pix_fmt yuv420p -movflags +faststart -shortest "${outputPath}"`;

    this.logger.log('Creating slide 12 with video overlay...');
    await execAsync(ffmpegCommand);

    // Verify output was created
    if (!fs.existsSync(outputPath)) {
      throw new Error('Slide 12 video was not created');
    }

    return outputPath;
  }

  /**
   * Create final concatenated video with all slides
   */
  private async createFinalVideo(outputPath: string, slide10Video: string, slide12Video: string): Promise<void> {
    // Build the complex filter for concatenating all slides
    const filterComplex = this.buildConcatenationFilter(slide10Video, slide12Video);
    
    const ffmpegCommand = `ffmpeg ${this.getInputFiles()} -filter_complex "${filterComplex}" -map "[out]" -c:v libx264 -preset veryfast -crf 18 -pix_fmt yuv420p -movflags +faststart -r 30 "${outputPath}"`;

    this.logger.log('Creating final concatenated video...');
    await execAsync(ffmpegCommand);
  }

  /**
   * Get all input files for the FFmpeg command
   */
  private getInputFiles(): string {
    const inputs: string[] = [];
    
    // Add all slide images (1-17)
    for (let i = 1; i <= 17; i++) {
      if (i === 10 || i === 12) {
        // Skip slides 10 and 12 as they'll be replaced with video versions
        continue;
      }
      const slidePath = path.join(this.assetsDir, `Slide${i}.PNG`);
      if (fs.existsSync(slidePath)) {
        inputs.push(`-loop 1 -i "${slidePath}"`);
      } else {
        // If slide doesn't exist, create a placeholder or skip
        this.logger.warn(`Slide ${i} not found at ${slidePath}`);
      }
    }
    
    // Add slide 10 video
    const slide10Video = path.join(this.assetsDir, 'Slide10_with_video.mp4');
    if (fs.existsSync(slide10Video)) {
      inputs.push(`-i "${slide10Video}"`);
    } else {
      this.logger.warn(`Slide 10 video not found at ${slide10Video}`);
    }
    
    // Add slide 12 video
    const slide12Video = path.join(this.assetsDir, 'Slide12_with_video_full.mp4');
    if (fs.existsSync(slide12Video)) {
      inputs.push(`-i "${slide12Video}"`);
    } else {
      this.logger.warn(`Slide 12 video not found at ${slide12Video}`);
    }
    
    return inputs.join(' ');
  }

  /**
   * Build the filter_complex string for concatenating all slides
   */
  private buildConcatenationFilter(slide10Video: string, slide12Video: string): string {
    const filters: string[] = [];
    let inputIndex = 0;
    
    // Process slides 1-9 (5 seconds each)
    for (let i = 1; i <= 9; i++) {
      filters.push(`[${inputIndex}:v]scale=1920:-1:flags=fast_bilinear,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:black,trim=duration=5[s${i}]`);
      inputIndex++;
    }
    
    // Process slide 10 (video version, 3 seconds)
    filters.push(`[${inputIndex}:v]trim=duration=3[s10]`);
    inputIndex++;
    
    // Process slide 11 (5 seconds)
    filters.push(`[${inputIndex}:v]scale=1920:-1:flags=fast_bilinear,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:black,trim=duration=5[s11]`);
    inputIndex++;
    
    // Process slide 12 (video version, full duration)
    filters.push(`[${inputIndex}:v][s12]`);
    inputIndex++;
    
    // Process slides 13-17 (5 seconds each)
    for (let i = 13; i <= 17; i++) {
      filters.push(`[${inputIndex}:v]scale=1920:-1:flags=fast_bilinear,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:black,trim=duration=5[s${i}]`);
      inputIndex++;
    }
    
    // Concatenate all segments
    const concatInputs = Array.from({length: 17}, (_, i) => `[s${i+1}]`).join('');
    filters.push(`${concatInputs}concat=n=17:v=1:a=0[out]`);
    
    return filters.join(';');
  }
}
