import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { OpportunitiesController } from './opportunities.controller';
import { OpportunitiesService } from './opportunities.service';
import { OpportunityWorkflowController } from './opportunity-workflow.controller';
import { OpportunityWorkflowService } from './opportunity-workflow.service';
import { SurveyController } from './survey.controller';
import { SurveyService } from './survey.service';
import { AutoSaveController } from './auto-save.controller';
import { AutoSaveService } from './auto-save.service';
import { SurveyImageService } from '../survey/survey-image.service';
import { SurveyReportService } from '../survey/survey-report.service';
import { PrismaModule } from '../prisma/prisma.module';
import { UserModule } from '../user/user.module';
import { IntegrationsModule } from '../integrations/integrations.module';
import { EmailModule } from '../email/email.module';
import { OneDriveModule } from '../onedrive/onedrive.module';
import { CloudinaryModule } from '../cloudinary/cloudinary.module';
import { OpportunityOutcomesModule } from '../opportunity-outcomes/opportunity-outcomes.module';
import { DynamicSurveyorService } from './dynamic-surveyor.service';
import { ContractModule } from '../contracts/contract.module';

@Module({
  imports: [ConfigModule, PrismaModule, UserModule, IntegrationsModule, EmailModule, OneDriveModule, CloudinaryModule, OpportunityOutcomesModule, ContractModule],
  controllers: [
    OpportunitiesController, 
    OpportunityWorkflowController,
    SurveyController,
    AutoSaveController
  ],
  providers: [
    OpportunitiesService, 
    OpportunityWorkflowService,
    SurveyService,
    AutoSaveService,
    SurveyImageService,
    SurveyReportService,
    DynamicSurveyorService
  ],
  exports: [OpportunitiesService, OpportunityWorkflowService, SurveyService, AutoSaveService, SurveyImageService, SurveyReportService],
})
export class OpportunitiesModule {} 