import { Module } from '@nestjs/common';
import { HealthController } from './health.controller';
import { PrismaModule } from '../prisma/prisma.module';
import { SessionManagementModule } from '../session-management/session-management.module';

@Module({
  imports: [PrismaModule, SessionManagementModule],
  controllers: [HealthController],
})
export class HealthModule {}





