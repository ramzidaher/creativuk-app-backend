import { Controller, Get, Post, Body, Patch, Param, Delete, UseGuards, Request } from '@nestjs/common';
import { AppointmentService } from './appointment.service';
import { CreateAppointmentDto, UpdateAppointmentDto } from './dto';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';

@Controller('appointments')
@UseGuards(JwtAuthGuard)
export class AppointmentController {
  constructor(private readonly appointmentService: AppointmentService) {}

  @Post()
  create(@Body() createAppointmentDto: CreateAppointmentDto, @Request() req) {
    // Add the logged-in user's ID to the appointment
    return this.appointmentService.create({
      ...createAppointmentDto,
      userId: req.user.sub
    });
  }

  @Get()
  findAll(@Request() req) {
    // Filter appointments by the logged-in user
    return this.appointmentService.findAllByUser(req.user.sub);
  }

  @Get('ghl')
  async getGhlAppointments(@Request() req) {
    // Get appointments from GoHighLevel for the logged-in user
    return this.appointmentService.getGhlAppointmentsForUser(req.user.sub);
  }

  @Get(':id')
  findOne(@Param('id') id: string, @Request() req) {
    // Ensure the appointment belongs to the logged-in user
    return this.appointmentService.findOneByUser(id, req.user.sub);
  }

  @Patch(':id')
  update(@Param('id') id: string, @Body() updateAppointmentDto: UpdateAppointmentDto, @Request() req) {
    // Ensure the appointment belongs to the logged-in user
    return this.appointmentService.updateByUser(id, updateAppointmentDto, req.user.sub);
  }

  @Delete(':id')
  remove(@Param('id') id: string, @Request() req) {
    // Ensure the appointment belongs to the logged-in user
    return this.appointmentService.removeByUser(id, req.user.sub);
  }
} 