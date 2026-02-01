import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { PrismaModule } from './prisma/prisma.module';
import { UserModule } from './user/user.module';
import { IntegrationsModule } from './integrations/integrations.module';
import { AppointmentModule } from './appointment/appointment.module';
import { OpportunitiesModule } from './opportunities/opportunities.module';
import { AuthModule } from './auth/auth.module';
import { ExcelFileCalculatorModule } from './excel-file-calculator/excel-file-calculator.module';
import { ExcelAutomationModule } from './excel-automation/excel-automation.module';
import { EPVSAutomationModule } from './epvs-automation/epvs-automation.module';
import { SignaturesModule } from './signatures/signatures.module';
import { OpenSolarModule } from './integrations/opensolar.module';
import { SessionManagementModule } from './session-management/session-management.module';
import { OneDriveModule } from './onedrive/onedrive.module';
import { CalendarModule } from './calendar/calendar.module';
import { DocuSignModule } from './docusign/docusign.module';
import { PdfSigningModule } from './pdf-signing/pdf-signing.module';
import { ContractModule } from './contracts/contract.module';
import { FreeSignaturesModule } from './signatures/free-signatures.module';
import { PdfSignatureModule } from './pdf-signing/pdf-signature.module';
import { EmailModule } from './email/email.module';
import { CalculatorProgressModule } from './calculator-progress/calculator-progress.module';
import { SystemSettingsModule } from './system-settings/system-settings.module';
import { OpportunityOutcomesModule } from './opportunity-outcomes/opportunity-outcomes.module';
import { DisclaimerModule } from './disclaimer/disclaimer.module';
import { ExpressFormModule } from './expressform/expressform.module';
import { EmailConfirmationModule } from './email_confirmation/email-confirmation.module';
import { AdminModule } from './admin/admin.module';
import { HealthModule } from './health/health.module';
import { CacheModule } from './cache/cache.module';
import { SolarProjectionModule } from './solar-projection/solar-projection.module';
import { TestController } from './test.controller';
import { AppController } from './app.controller';
import { OAuthController } from './oauth.controller';
import { AdobeSignController } from './adobe-sign.controller';
import { AppService } from './app.service';

@Module({
  imports: [
    ConfigModule.forRoot({
      isGlobal: true,
    }),
    PrismaModule, 
    UserModule, 
    IntegrationsModule, 
    AppointmentModule, 
    OpportunitiesModule, 
    AuthModule,
    ExcelFileCalculatorModule,
    ExcelAutomationModule,
    EPVSAutomationModule,
    SignaturesModule,
    OpenSolarModule,
    SessionManagementModule,
    OneDriveModule,
    CalendarModule,
    DocuSignModule,
    PdfSigningModule,
    ContractModule,
    FreeSignaturesModule,
    PdfSignatureModule,
    EmailModule,
    CalculatorProgressModule,
    SystemSettingsModule,
    OpportunityOutcomesModule,
    DisclaimerModule,
    ExpressFormModule,
    EmailConfirmationModule,
    AdminModule,
    HealthModule,
    CacheModule,
    SolarProjectionModule
  ],
  controllers: [TestController, AppController, OAuthController, AdobeSignController],
  providers: [AppService],
})
export class AppModule {}
