import { Module } from '@nestjs/common';
import { SolarProjectionController } from './solar-projection.controller';
import { SolarProjectionService } from './solar-projection.service';
import { EPVSAutomationModule } from '../epvs-automation/epvs-automation.module';
import { ExcelAutomationModule } from '../excel-automation/excel-automation.module';
import { ExcelFileCalculatorModule } from '../excel-file-calculator/excel-file-calculator.module';

@Module({
  imports: [EPVSAutomationModule, ExcelAutomationModule, ExcelFileCalculatorModule],
  controllers: [SolarProjectionController],
  providers: [SolarProjectionService],
  exports: [SolarProjectionService]
})
export class SolarProjectionModule {}


