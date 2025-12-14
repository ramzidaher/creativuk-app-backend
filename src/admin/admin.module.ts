import { Module } from '@nestjs/common';
import { AdminAnalyticsService } from './admin-analytics.service';
import { AdminAnalyticsController } from './admin-analytics.controller';
import { AdminOpportunityDetailsService } from './admin-opportunity-details.service';
import { AdminOpportunityDetailsController } from './admin-opportunity-details.controller';
import { PrismaModule } from '../prisma/prisma.module';
import { UserModule } from '../user/user.module';
import { OpportunitiesModule } from '../opportunities/opportunities.module';
import { AuthModule } from '../auth/auth.module';
import { OneDriveModule } from '../onedrive/onedrive.module';

@Module({
  imports: [
    PrismaModule, 
    UserModule, 
    OpportunitiesModule, 
    AuthModule,
    OneDriveModule
  ],
  controllers: [AdminAnalyticsController, AdminOpportunityDetailsController],
  providers: [AdminAnalyticsService, AdminOpportunityDetailsService],
  exports: [AdminAnalyticsService, AdminOpportunityDetailsService],
})
export class AdminModule {}

