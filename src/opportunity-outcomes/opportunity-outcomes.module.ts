import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { OpportunityOutcomesController } from './opportunity-outcomes.controller';
import { OpportunityOutcomesService } from './opportunity-outcomes.service';
import { PrismaModule } from '../prisma/prisma.module';
import { IntegrationsModule } from '../integrations/integrations.module';

@Module({
  imports: [ConfigModule, PrismaModule, IntegrationsModule],
  controllers: [OpportunityOutcomesController],
  providers: [OpportunityOutcomesService],
  exports: [OpportunityOutcomesService],
})
export class OpportunityOutcomesModule {}
