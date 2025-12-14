import { Module } from '@nestjs/common';
import { SessionManagementService } from './session-management.service';
import { SessionManagementController } from './session-management.controller';
import { ComProcessManagerService } from './com-process-manager.service';
import { PrismaModule } from '../prisma/prisma.module';

@Module({
  imports: [PrismaModule],
  controllers: [SessionManagementController],
  providers: [SessionManagementService, ComProcessManagerService],
  exports: [SessionManagementService, ComProcessManagerService],
})
export class SessionManagementModule {}
