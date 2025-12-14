import { Module } from '@nestjs/common';
import { EmailConfirmationController } from './email-confirmation.controller';
import { EmailConfirmationService } from './email-confirmation.service';
import { PdfSignatureModule } from '../pdf-signature/pdf-signature.module';

@Module({
  imports: [PdfSignatureModule],
  controllers: [EmailConfirmationController],
  providers: [EmailConfirmationService],
  exports: [EmailConfirmationService],
})
export class EmailConfirmationModule {}



