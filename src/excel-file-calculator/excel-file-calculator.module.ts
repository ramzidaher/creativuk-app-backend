import { Module, forwardRef } from '@nestjs/common';
import { ExcelCellDetectorService } from '../excel-automation/excel-cell-detector.service';
import { ExcelCellDetectorController } from '../excel-automation/excel-cell-detector.controller';
import { PresentationController, PublicPresentationController } from './presentation.controller';
import { PresentationService } from './presentation.service';
import { SessionManagementModule } from '../session-management/session-management.module';
import { PowerpointMp4Service } from '../video-generation/powerpoint-mp4.service';
import { FfmpegVideoService } from '../video-generation/ffmpeg-video.service';
import { ImageGenerationService } from '../video-generation/image-generation.service';
import { ExcelAutomationModule } from '../excel-automation/excel-automation.module';

@Module({
  imports: [forwardRef(() => SessionManagementModule), forwardRef(() => ExcelAutomationModule)],
  controllers: [ExcelCellDetectorController, PresentationController, PublicPresentationController],
  providers: [ExcelCellDetectorService, PresentationService, PowerpointMp4Service, FfmpegVideoService, ImageGenerationService],
  exports: [ExcelCellDetectorService, PresentationService, PowerpointMp4Service, FfmpegVideoService, ImageGenerationService],
})
export class ExcelFileCalculatorModule {}
