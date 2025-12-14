import { Module } from '@nestjs/common';
import { AppointmentService } from './appointment.service';
import { AppointmentController } from './appointment.controller';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { UserModule } from '../user/user.module';

@Module({
  imports: [UserModule],
  providers: [AppointmentService, GoHighLevelService],
  controllers: [AppointmentController],
  exports: [AppointmentService],
})
export class AppointmentModule {} 