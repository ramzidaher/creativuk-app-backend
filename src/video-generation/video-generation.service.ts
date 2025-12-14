import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

@Injectable()
export class VideoGenerationService {
  private readonly logger = new Logger(VideoGenerationService.name);
  private readonly outputDir = path.join(process.cwd(), 'src', 'video-generation', 'output');
  private readonly publicDir = path.join(process.cwd(), 'public', 'videos');

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
   * Generate MP4 video from proposal data
   */
  async generateProposalVideo(data: {
    opportunityId: string;
    customerName: string;
    date: string;
    postcode: string;
    solarData: any;
  }): Promise<{ success: boolean; videoPath?: string; publicUrl?: string; error?: string }> {
    try {
      this.logger.log(`Generating proposal video for opportunity: ${data.opportunityId}`);

      const timestamp = Date.now();
      const safeCustomerName = data.customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const videoFilename = `proposal_${safeCustomerName}_${data.opportunityId}_${timestamp}.mp4`;
      const videoPath = path.join(this.outputDir, videoFilename);
      const publicVideoPath = path.join(this.publicDir, videoFilename);

      // Create the video data file
      const videoData = {
        customerName: data.customerName,
        date: data.date,
        postcode: data.postcode,
        solarData: data.solarData,
      };

      const dataFilePath = path.join(this.outputDir, `video_data_${timestamp}.json`);
      fs.writeFileSync(dataFilePath, JSON.stringify(videoData, null, 2));

      // Generate video using Remotion
      const remotionCommand = `npx remotion render src/video-components/Root.tsx ProposalVideo --props='${dataFilePath}' --out='${videoPath}' --concurrency=1`;

      this.logger.log(`Executing Remotion command: ${remotionCommand}`);
      
      const { stdout, stderr } = await execAsync(remotionCommand, {
        cwd: process.cwd(),
        timeout: 300000, // 5 minutes timeout
      });

      if (stderr && !stderr.includes('warning')) {
        this.logger.error(`Remotion stderr: ${stderr}`);
      }

      this.logger.log(`Remotion stdout: ${stdout}`);

      // Check if video was created
      if (!fs.existsSync(videoPath)) {
        throw new Error('Video file was not created');
      }

      // Copy to public directory for serving
      fs.copyFileSync(videoPath, publicVideoPath);

      // Clean up temporary data file
      if (fs.existsSync(dataFilePath)) {
        fs.unlinkSync(dataFilePath);
      }

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
   * Get video file info
   */
  async getVideoInfo(videoPath: string): Promise<{ size: number; exists: boolean }> {
    try {
      if (!fs.existsSync(videoPath)) {
        return { size: 0, exists: false };
      }

      const stats = fs.statSync(videoPath);
      return { size: stats.size, exists: true };
    } catch (error) {
      this.logger.error(`Error getting video info: ${error.message}`);
      return { size: 0, exists: false };
    }
  }

  /**
   * Delete video files
   */
  async deleteVideo(videoPath: string): Promise<boolean> {
    try {
      if (fs.existsSync(videoPath)) {
        fs.unlinkSync(videoPath);
        this.logger.log(`Deleted video: ${videoPath}`);
        return true;
      }
      return false;
    } catch (error) {
      this.logger.error(`Error deleting video: ${error.message}`);
      return false;
    }
  }

  /**
   * List available videos
   */
  async listVideos(): Promise<string[]> {
    try {
      if (!fs.existsSync(this.outputDir)) {
        return [];
      }

      const files = fs.readdirSync(this.outputDir);
      return files.filter(file => file.endsWith('.mp4'));
    } catch (error) {
      this.logger.error(`Error listing videos: ${error.message}`);
      return [];
    }
  }
}
