import { Module } from '@nestjs/common';
import { CalculatorProgressController } from './calculator-progress.controller';
import { CalculatorProgressService } from './calculator-progress.service';
import { PrismaModule } from '../prisma/prisma.module';
import { ExcelAutomationModule } from '../excel-automation/excel-automation.module';
import { EPVSAutomationModule } from '../epvs-automation/epvs-automation.module';

@Module({
  imports: [PrismaModule, ExcelAutomationModule, EPVSAutomationModule],
  controllers: [CalculatorProgressController],
  providers: [CalculatorProgressService],
  exports: [CalculatorProgressService],
})
export class CalculatorProgressModule {}

