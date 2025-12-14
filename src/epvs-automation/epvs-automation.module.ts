import { Module, forwardRef } from '@nestjs/common';
import { EPVSAutomationService } from './epvs-automation.service';
import { EPVSAutomationController } from './epvs-automation.controller';
import { PdfSignatureModule } from '../pdf-signature/pdf-signature.module';
import { SessionManagementModule } from '../session-management/session-management.module';

@Module({
  imports: [PdfSignatureModule, forwardRef(() => SessionManagementModule)],
  controllers: [EPVSAutomationController],
  providers: [EPVSAutomationService],
  exports: [EPVSAutomationService],
})
export class EPVSAutomationModule {}

