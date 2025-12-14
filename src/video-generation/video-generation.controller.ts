import { Controller, Post, Get, Delete, Body, Param, UseGuards, Res } from '@nestjs/common';
import { Response } from 'express';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { VideoGenerationService } from './video-generation.service';
import { FfmpegVideoService } from './ffmpeg-video.service';
import * as path from 'path';
import * as fs from 'fs';

@Controller('video')
@UseGuards(JwtAuthGuard)
export class VideoGenerationController {
  constructor(
    private readonly videoGenerationService: VideoGenerationService,
    private readonly ffmpegVideoService: FfmpegVideoService
  ) {}

  /**
   * Generate proposal video
   */
  @Post('generate-proposal')
  async generateProposalVideo(@Body() data: {
    opportunityId: string;
    customerName: string;
    date: string;
    postcode: string;
    solarData: any;
  }) {
    try {
      const result = await this.videoGenerationService.generateProposalVideo(data);
      
      if (result.success) {
        return {
          success: true,
          message: 'Video generated successfully',
          data: {
            videoPath: result.videoPath,
            publicUrl: result.publicUrl,
          },
        };
      } else {
        return {
          success: false,
          message: 'Video generation failed',
          error: result.error,
        };
      }
    } catch (error) {
      return {
        success: false,
        message: 'Video generation error',
        error: error.message,
      };
    }
  }


  /**
   * Get video file
   */
  @Get('serve/:filename')
  async serveVideo(@Param('filename') filename: string, @Res() res: Response) {
    try {
      const publicDir = path.join(process.cwd(), 'public', 'videos');
      const videoPath = path.join(publicDir, filename);

      if (!fs.existsSync(videoPath)) {
        return res.status(404).json({ error: 'Video not found' });
      }

      const stat = fs.statSync(videoPath);
      const fileSize = stat.size;
      const range = res.req.headers.range;

      if (range) {
        // Handle range requests for video streaming
        const parts = range.replace(/bytes=/, '').split('-');
        const start = parseInt(parts[0], 10);
        const end = parts[1] ? parseInt(parts[1], 10) : fileSize - 1;
        const chunksize = (end - start) + 1;
        const file = fs.createReadStream(videoPath, { start, end });
        const head = {
          'Content-Range': `bytes ${start}-${end}/${fileSize}`,
          'Accept-Ranges': 'bytes',
          'Content-Length': chunksize,
          'Content-Type': 'video/mp4',
        };
        res.writeHead(206, head);
        file.pipe(res);
      } else {
        // Serve entire file
        const head = {
          'Content-Length': fileSize,
          'Content-Type': 'video/mp4',
        };
        res.writeHead(200, head);
        fs.createReadStream(videoPath).pipe(res);
      }
    } catch (error) {
      res.status(500).json({ error: 'Error serving video' });
    }
  }

  /**
   * List available videos
   */
  @Get('list')
  async listVideos() {
    try {
      const videos = await this.videoGenerationService.listVideos();
      return {
        success: true,
        videos: videos,
      };
    } catch (error) {
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Delete video
   */
  @Delete(':filename')
  async deleteVideo(@Param('filename') filename: string) {
    try {
      const publicDir = path.join(process.cwd(), 'public', 'videos');
      const outputDir = path.join(process.cwd(), 'src', 'video-generation', 'output');
      
      const publicVideoPath = path.join(publicDir, filename);
      const outputVideoPath = path.join(outputDir, filename);

      let deleted = false;
      if (fs.existsSync(publicVideoPath)) {
        fs.unlinkSync(publicVideoPath);
        deleted = true;
      }
      if (fs.existsSync(outputVideoPath)) {
        fs.unlinkSync(outputVideoPath);
        deleted = true;
      }

      if (deleted) {
        return {
          success: true,
          message: 'Video deleted successfully',
        };
      } else {
        return {
          success: false,
          message: 'Video not found',
        };
      }
    } catch (error) {
      return {
        success: false,
        error: error.message,
      };
    }
  }
}
