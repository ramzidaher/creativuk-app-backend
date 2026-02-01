import {
  IsString,
  IsOptional,
  IsEnum,
  IsObject,
  IsArray,
  ValidateNested,
} from 'class-validator';
import { Type } from 'class-transformer';

export enum FormSubmissionStatus {
  DRAFT = 'DRAFT',
  SUBMITTED = 'SUBMITTED',
  PENDING = 'PENDING',
  APPROVED = 'APPROVED',
  REJECTED = 'REJECTED',
}

export class FormSubmissionFieldDto {
  @IsString()
  fieldId: string;

  @IsObject()
  value: any; // Can be string, number, boolean, array, etc.
}

export class SubmitFormDto {
  @IsString()
  formId: string;

  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => FormSubmissionFieldDto)
  fields: FormSubmissionFieldDto[];

  @IsOptional()
  @IsEnum(FormSubmissionStatus)
  status?: FormSubmissionStatus; // DRAFT or SUBMITTED
}

export class UpdateSubmissionDto {
  @IsOptional()
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => FormSubmissionFieldDto)
  fields?: FormSubmissionFieldDto[];

  @IsOptional()
  @IsEnum(FormSubmissionStatus)
  status?: FormSubmissionStatus;
}

export class ApproveSubmissionDto {
  @IsOptional()
  @IsString()
  @IsOptional()
  comments?: string;
}

export class RejectSubmissionDto {
  @IsString()
  reason: string;

  @IsOptional()
  @IsString()
  comments?: string;
}

