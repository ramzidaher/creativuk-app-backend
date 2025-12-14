import { Module } from '@nestjs/common';
import { DisclaimerController } from './disclaimer.controller';
import { DisclaimerService } from './disclaimer.service';
import { PdfSignatureModule } from '../pdf-signature/pdf-signature.module';

@Module({
  imports: [PdfSignatureModule],
  controllers: [DisclaimerController],
  providers: [DisclaimerService],
  exports: [DisclaimerService],
})
export class DisclaimerModule {}
