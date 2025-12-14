import { Module } from '@nestjs/common';
import { PdfSignatureService } from './pdf-signature.service';
import { DigitalSignatureService } from './digital-signature.service';
import { DigitalSignatureController } from './digital-signature.controller';

@Module({
  controllers: [DigitalSignatureController],
  providers: [PdfSignatureService, DigitalSignatureService],
  exports: [PdfSignatureService, DigitalSignatureService],
})
export class PdfSignatureModule {}
