import { AppointmentStatus, SourceChannel } from '@prisma/client';
import { IsString, IsOptional, IsEnum, IsDateString } from 'class-validator';

export class CreateAppointmentDto {
  @IsDateString()
  scheduledAt: string;

  @IsString()
  customerName: string;

  @IsString()
  customerPhone: string;

  @IsOptional()
  @IsString()
  customerEmail?: string;

  @IsString()
  address: string;

  @IsEnum(AppointmentStatus)
  status: AppointmentStatus;

  @IsOptional()
  @IsString()
  ghlAppointmentId?: string;

  @IsString()
  userId: string;

  @IsOptional()
  @IsString()
  notes?: string;

  @IsEnum(SourceChannel)
  sourceChannel: SourceChannel;
} 