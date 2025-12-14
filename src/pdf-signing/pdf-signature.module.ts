import { Module } from '@nestjs/common';
import { PdfSignatureController } from './pdf-signature.controller';
import { PdfSignatureService } from './pdf-signature.service';

@Module({
  controllers: [PdfSignatureController],
  providers: [PdfSignatureService],
  exports: [PdfSignatureService],
})
export class PdfSignatureModule {}
