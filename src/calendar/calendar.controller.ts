import { Controller, Get, Post, Param, Query, Body, HttpException, HttpStatus, UseGuards, Request } from '@nestjs/common';
import { CalendarService } from './calendar.service';
import { BookAppointmentDto } from './dto/book-appointment.dto';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';

@Controller('calendar')
@UseGuards(JwtAuthGuard)
export class CalendarController {
  constructor(private readonly calendarService: CalendarService) {}

  @Get('current/events')
  async getCurrentUserCalendarEvents(
    @Request() req: any,
    @Query('startDate') startDate?: string,
    @Query('endDate') endDate?: string,
  ) {
    try {
      return await this.calendarService.getCurrentUserCalendarEvents(startDate, endDate, req.user);
    } catch (error) {
      throw new HttpException(
        {
          message: 'Failed to get current user calendar events',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  @Get(':calendarName/events')
  async getCalendarEvents(
    @Param('calendarName') calendarName: string,
    @Query('startDate') startDate?: string,
    @Query('endDate') endDate?: string,
  ) {
    try {
      return await this.calendarService.getCalendarEvents(calendarName, startDate, endDate);
    } catch (error) {
      throw new HttpException(
        {
          message: 'Failed to get calendar events',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  @Get('user/:username/calendars')
  async getUserCalendars(@Param('username') username: string) {
    try {
      return await this.calendarService.getUserCalendars(username);
    } catch (error) {
      throw new HttpException(
        {
          message: 'Failed to get user calendars',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  @Post('book-appointment')
  async bookAppointment(@Body() bookAppointmentDto: BookAppointmentDto) {
    try {
      return await this.calendarService.bookAppointment(bookAppointmentDto);
    } catch (error) {
      throw new HttpException(
        {
          message: 'Failed to book appointment',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  @Get(':calendarName/availability/:date')
  async getAvailability(
    @Param('calendarName') calendarName: string,
    @Param('date') date: string,
  ) {
    try {
      return await this.calendarService.getAvailability(calendarName, date);
    } catch (error) {
      throw new HttpException(
        {
          message: 'Failed to get availability',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

}
