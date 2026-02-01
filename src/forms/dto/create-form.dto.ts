import {
  IsString,
  IsOptional,
  IsEnum,
  IsBoolean,
  IsDateString,
  IsArray,
  ValidateNested,
  IsObject,
  MinLength,
  MaxLength,
} from 'class-validator';
import { Type } from 'class-transformer';

export enum FormCategory {
  HR = 'HR',
  INSTALLER = 'INSTALLER',
  OTHER = 'OTHER',
}

export enum FormFieldType {
  TEXT = 'TEXT',
  TEXTAREA = 'TEXTAREA',
  NUMBER = 'NUMBER',
  EMAIL = 'EMAIL',
  DATE = 'DATE',
  DATETIME = 'DATETIME',
  DROPDOWN = 'DROPDOWN',
  RADIO = 'RADIO',
  CHECKBOX = 'CHECKBOX',
  MULTI_SELECT = 'MULTI_SELECT',
  IMAGE_UPLOAD = 'IMAGE_UPLOAD',
  SIGNATURE = 'SIGNATURE',
  FILE_UPLOAD = 'FILE_UPLOAD',
}

export class ValidationRulesDto {
  @IsOptional()
  @IsObject()
  min?: number;

  @IsOptional()
  @IsObject()
  max?: number;

  @IsOptional()
  @IsString()
  pattern?: string;

  @IsOptional()
  @IsBoolean()
  email?: boolean;

  @IsOptional()
  @IsBoolean()
  url?: boolean;

  @IsOptional()
  @IsString()
  minLength?: number;

  @IsOptional()
  @IsString()
  maxLength?: number;
}

export class ConditionalLogicDto {
  @IsString()
  field: string;

  @IsString()
  operator: 'equals' | 'notEquals' | 'contains' | 'greaterThan' | 'lessThan';

  @IsOptional()
  value?: any;
}

export class FormFieldDto {
  @IsEnum(FormFieldType)
  fieldType: FormFieldType;

  @IsString()
  @MinLength(1)
  @MaxLength(200)
  label: string;

  @IsString()
  @MinLength(1)
  @MaxLength(100)
  name: string;

  @IsOptional()
  @IsString()
  @MaxLength(500)
  placeholder?: string;

  @IsOptional()
  @IsString()
  @MaxLength(1000)
  helpText?: string;

  @IsOptional()
  @IsBoolean()
  isRequired?: boolean;

  @IsOptional()
  @ValidateNested()
  @Type(() => ValidationRulesDto)
  validationRules?: ValidationRulesDto;

  @IsOptional()
  @IsArray()
  options?: Array<{ label: string; value: string }>;

  @IsOptional()
  @ValidateNested()
  @Type(() => ConditionalLogicDto)
  conditionalLogic?: ConditionalLogicDto;

  @IsOptional()
  order?: number;
}

export class CreateFormDto {
  @IsString()
  @MinLength(1)
  @MaxLength(200)
  title: string;

  @IsOptional()
  @IsString()
  @MaxLength(1000)
  description?: string;

  @IsEnum(FormCategory)
  category: FormCategory;

  @IsOptional()
  @IsBoolean()
  isActive?: boolean;

  @IsOptional()
  @IsBoolean()
  isPublished?: boolean;

  @IsOptional()
  @IsBoolean()
  requiresApproval?: boolean;

  @IsOptional()
  @IsDateString()
  deadline?: string;

  @IsOptional()
  @IsString()
  @MaxLength(2000)
  instructions?: string;

  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => FormFieldDto)
  fields: FormFieldDto[];
}

export class UpdateFormDto {
  @IsOptional()
  @IsString()
  @MinLength(1)
  @MaxLength(200)
  title?: string;

  @IsOptional()
  @IsString()
  @MaxLength(1000)
  description?: string;

  @IsOptional()
  @IsEnum(FormCategory)
  category?: FormCategory;

  @IsOptional()
  @IsBoolean()
  isActive?: boolean;

  @IsOptional()
  @IsBoolean()
  isPublished?: boolean;

  @IsOptional()
  @IsBoolean()
  requiresApproval?: boolean;

  @IsOptional()
  @IsDateString()
  deadline?: string;

  @IsOptional()
  @IsString()
  @MaxLength(2000)
  instructions?: string;

  @IsOptional()
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => FormFieldDto)
  fields?: FormFieldDto[];
}

