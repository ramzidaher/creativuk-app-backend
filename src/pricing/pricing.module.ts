import { Module, forwardRef } from '@nestjs/common';
import { PricingController } from './pricing.controller';
import { PricingService } from './pricing.service';
import { SessionManagementModule } from '../session-management/session-management.module';

@Module({
  imports: [forwardRef(() => SessionManagementModule)],
  controllers: [PricingController],
  providers: [PricingService],
  exports: [PricingService],
})
export class PricingModule {}
