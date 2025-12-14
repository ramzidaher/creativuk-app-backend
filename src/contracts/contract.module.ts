import { Module, forwardRef } from '@nestjs/common';
import { ContractController } from './contract.controller';
import { ContractService } from './contract.service';
import { DocuSealService } from '../integrations/docuseal.service';
import { DocuSealController } from '../integrations/docuseal.controller';
import { WebhooksController } from '../integrations/webhooks.controller';
import { PdfGeneratorService } from './pdf-generator.service';
import { EPVSAutomationModule } from '../epvs-automation/epvs-automation.module';
import { ExcelAutomationModule } from '../excel-automation/excel-automation.module';
import { PdfSignatureModule } from '../pdf-signature/pdf-signature.module';
import { SessionManagementModule } from '../session-management/session-management.module';
import { OpportunitiesModule } from '../opportunities/opportunities.module';
import { PrismaModule } from '../prisma/prisma.module';
import { OneDriveModule } from '../onedrive/onedrive.module';

@Module({
  imports: [
    PrismaModule,
    forwardRef(() => EPVSAutomationModule),
    forwardRef(() => ExcelAutomationModule),
    forwardRef(() => PdfSignatureModule),
    forwardRef(() => SessionManagementModule),
    forwardRef(() => OpportunitiesModule),
    forwardRef(() => OneDriveModule),
  ],
  controllers: [ContractController, DocuSealController, WebhooksController],
  providers: [
    ContractService, 
    DocuSealService, 
    PdfGeneratorService
  ],
  exports: [ContractService, DocuSealService, PdfGeneratorService],
})
export class ContractModule {}
