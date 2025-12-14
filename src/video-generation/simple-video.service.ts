import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

@Injectable()
export class SimpleVideoService {
  private readonly logger = new Logger(SimpleVideoService.name);
  private readonly outputDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'output');
  private readonly publicDir = path.join(process.cwd(), 'public', 'videos');

  constructor() {
    // Ensure public directory exists
    if (!fs.existsSync(this.publicDir)) {
      fs.mkdirSync(this.publicDir, { recursive: true });
    }
  }

  /**
   * Convert PDF to MP4 using ImageMagick and FFmpeg
   * This is much simpler than Remotion
   */
  async convertPdfToVideo(pdfPath: string, outputVideoPath: string): Promise<{ success: boolean; error?: string }> {
    try {
      this.logger.log(`Converting PDF to video: ${pdfPath} -> ${outputVideoPath}`);

      // Check if PDF exists
      if (!fs.existsSync(pdfPath)) {
        throw new Error(`PDF file not found: ${pdfPath}`);
      }

      // Create temporary directory for images
      const tempDir = path.join(process.cwd(), 'temp', 'pdf-to-video', Date.now().toString());
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }

      try {
        // Step 1: Convert PDF to images using ImageMagick
        this.logger.log('Converting PDF to images...');
        const imagePattern = path.join(tempDir, 'page_%03d.png');
        
        const magickCommand = `magick "${pdfPath}" -density 150 "${imagePattern}"`;
        await execAsync(magickCommand);
        
        // Check if images were created
        const imageFiles = fs.readdirSync(tempDir).filter(file => file.endsWith('.png'));
        if (imageFiles.length === 0) {
          throw new Error('No images were generated from PDF');
        }

        this.logger.log(`Generated ${imageFiles.length} images from PDF`);

        // Step 2: Convert images to video using FFmpeg
        this.logger.log('Converting images to video...');
        const inputPattern = path.join(tempDir, 'page_%03d.png');
        const ffmpegCommand = `ffmpeg -framerate 0.5 -i "${inputPattern}" -c:v libx264 -pix_fmt yuv420p -r 30 "${outputVideoPath}" -y`;
        
        await execAsync(ffmpegCommand);

        // Verify video was created
        if (!fs.existsSync(outputVideoPath)) {
          throw new Error('Video file was not created');
        }

        const stats = fs.statSync(outputVideoPath);
        this.logger.log(`Video created successfully: ${outputVideoPath} (${(stats.size / 1024 / 1024).toFixed(2)} MB)`);

        return { success: true };

      } finally {
        // Clean up temporary directory
        try {
          fs.rmSync(tempDir, { recursive: true, force: true });
        } catch (cleanupError) {
          this.logger.warn(`Failed to cleanup temp directory: ${cleanupError.message}`);
        }
      }

    } catch (error) {
      this.logger.error(`PDF to video conversion failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }

  /**
   * Generate video from existing PDF presentation
   */
  async generateVideoFromPdf(pdfPath: string, opportunityId: string, customerName: string): Promise<{
    success: boolean;
    videoPath?: string;
    publicUrl?: string;
    error?: string;
  }> {
    try {
      const timestamp = Date.now();
      const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const videoFilename = `proposal_${safeCustomerName}_${opportunityId}_${timestamp}.mp4`;
      const videoPath = path.join(this.outputDir, videoFilename);
      const publicVideoPath = path.join(this.publicDir, videoFilename);

      // Convert PDF to video
      const conversionResult = await this.convertPdfToVideo(pdfPath, videoPath);
      
      if (!conversionResult.success) {
        return {
          success: false,
          error: conversionResult.error,
        };
      }

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
   * Alternative: Use online conversion service (simpler but requires internet)
   */
  async convertPdfToVideoOnline(pdfPath: string, outputVideoPath: string): Promise<{ success: boolean; error?: string }> {
    // This would use an online service like the ones mentioned in the web search
    // For now, we'll stick with the local conversion method
    return this.convertPdfToVideo(pdfPath, outputVideoPath);
  }
}
