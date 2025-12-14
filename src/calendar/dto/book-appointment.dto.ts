import { IsString, IsOptional, IsNotEmpty } from 'class-validator';

export class BookAppointmentDto {
  @IsString()
  @IsNotEmpty()
  opportunityId: string;

  @IsString()
  @IsOptional()
  customerName?: string;

  @IsString()
  @IsOptional()
  customerAddress?: string;

  @IsString()
  @IsNotEmpty()
  calendar: string;

  @IsString()
  @IsNotEmpty()
  date: string;

  @IsString()
  @IsNotEmpty()
  timeSlot: string;

  @IsString()
  @IsOptional()
  installer?: string;

  @IsString()
  @IsOptional()
  surveyor?: string;

  @IsString()
  @IsOptional()
  surveyorEmail?: string;
}
