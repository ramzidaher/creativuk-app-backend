import { Module } from '@nestjs/common';
import { OpenSolarService } from './opensolar.service';
import { GoHighLevelService } from './gohighlevel.service';
import { GhlAuthModule } from './ghl-auth/ghl-auth.module';

@Module({
  imports: [GhlAuthModule],
  providers: [OpenSolarService, GoHighLevelService],
  exports: [OpenSolarService, GoHighLevelService, GhlAuthModule],
})
export class IntegrationsModule {}
