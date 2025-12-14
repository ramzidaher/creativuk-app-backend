import { IsString, IsNumber, IsOptional, IsEnum, IsObject, IsDateString } from 'class-validator';
import { StepType } from '@prisma/client';

export { StepType };

export enum StepStatus {
  PENDING = 'PENDING',
  IN_PROGRESS = 'IN_PROGRESS',
  COMPLETED = 'COMPLETED',
  SKIPPED = 'SKIPPED',
}

export enum OpportunityStatus {
  IN_PROGRESS = 'IN_PROGRESS',
  COMPLETED = 'COMPLETED',
  PAUSED = 'PAUSED',
  CANCELLED = 'CANCELLED',
}

export class StartOpportunityDto {
  @IsString()
  ghlOpportunityId: string;
}

export class UpdateStepDto {
  @IsNumber()
  stepNumber: number;

  @IsEnum(StepStatus)
  status: StepStatus;

  @IsOptional()
  @IsObject()
  data?: any;
}

export class CompleteStepDto {
  @IsNumber()
  stepNumber: number;

  @IsOptional()
  @IsObject()
  data?: any;
}

export class OpportunityDetailsDto {
  customerName: string;
  address: string;
  contactEmail?: string;
  contactPhone?: string;
  monetaryValue?: number;
  stageName?: string;
}

export class OpportunityProgressDto {
  id: string;
  ghlOpportunityId: string;
  currentStep: number;
  totalSteps: number;
  status: OpportunityStatus;
  startedAt: Date;
  lastActivityAt: Date;
  completedAt?: Date;
  stepData?: any;
  steps: OpportunityStepDto[];
  opportunityDetails?: OpportunityDetailsDto;
  userInfo?: UserInfoDto; // For admin views
}

export class UserInfoDto {
  id: string;
  name: string;
  username: string;
  email: string;
  role: string;
}

export class OpportunityStepDto {
  id: string;
  stepNumber: number;
  stepType: StepType;
  status: StepStatus;
  data?: any;
  startedAt?: Date;
  completedAt?: Date;
}

export class WorkflowStepConfig {
  stepNumber: number;
  stepType: StepType;
  title: string;
  description: string;
  required: boolean;
  estimatedDuration: number; // in minutes
}

export class WorkflowStepResponseDto {
  stepNumber: number;
  stepType: StepType;
  title: string;
  description: string;
  required: boolean;
  estimatedDuration: number;
  status: StepStatus;
  startedAt?: Date;
  completedAt?: Date;
  data?: any;
}

export class WorkflowResponseDto {
  id: string;
  ghlOpportunityId: string;
  currentStep: number;
  totalSteps: number;
  status: OpportunityStatus;
  startedAt: Date;
  lastActivityAt: Date;
  completedAt?: Date;
  stepData?: any;
  steps: WorkflowStepResponseDto[];
} 