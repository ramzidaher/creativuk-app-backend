import { Module, forwardRef } from '@nestjs/common';
import { ExcelAutomationController } from './excel-automation.controller';
import { ExcelAutomationService } from './excel-automation.service';
import { PdfSignatureModule } from '../pdf-signature/pdf-signature.module';
import { SessionManagementModule } from '../session-management/session-management.module';

@Module({
  imports: [PdfSignatureModule, forwardRef(() => SessionManagementModule)],
  controllers: [ExcelAutomationController],
  providers: [ExcelAutomationService],
  exports: [ExcelAutomationService],
})
export class ExcelAutomationModule {}
