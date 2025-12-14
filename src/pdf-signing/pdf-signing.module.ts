import { Module } from '@nestjs/common';
import { PdfSigningService } from './pdf-signing.service';
import { PdfSigningController } from './pdf-signing.controller';

@Module({
  providers: [PdfSigningService],
  controllers: [PdfSigningController],
  exports: [PdfSigningService],
})
export class PdfSigningModule {}
