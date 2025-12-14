import { Module } from '@nestjs/common';
import { VideoGenerationService } from './video-generation.service';
import { VideoGenerationController } from './video-generation.controller';
import { FfmpegVideoService } from './ffmpeg-video.service';

@Module({
  providers: [VideoGenerationService, FfmpegVideoService],
  controllers: [VideoGenerationController],
  exports: [VideoGenerationService, FfmpegVideoService],
})
export class VideoGenerationModule {}
