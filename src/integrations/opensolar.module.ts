import { Module } from '@nestjs/common';
import { OpenSolarController } from './opensolar.controller';
import { OpenSolarService } from './opensolar.service';
import { OpenSolarPublicController } from './opensolar-public.controller';
import { OpenSolarPublicService } from './opensolar-public.service';
import { PrismaModule } from '../prisma/prisma.module';
import { ExcelAutomationModule } from '../excel-automation/excel-automation.module';
import { EPVSAutomationModule } from '../epvs-automation/epvs-automation.module';

@Module({
  imports: [PrismaModule, ExcelAutomationModule, EPVSAutomationModule],
  controllers: [OpenSolarController, OpenSolarPublicController],
  providers: [OpenSolarService, OpenSolarPublicService],
  exports: [OpenSolarService, OpenSolarPublicService],
})
export class OpenSolarModule {}
