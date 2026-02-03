import { Module } from '@nestjs/common';
import { ExpressFormController } from './expressform.controller';
import { ExpressFormService } from './expressform.service';
import { PdfSignatureModule } from '../pdf-signature/pdf-signature.module';

@Module({
  imports: [PdfSignatureModule],
  controllers: [ExpressFormController],
  providers: [ExpressFormService],
  exports: [ExpressFormService],
})
export class ExpressFormModule {}










