import { IsEmail, IsString, IsEnum, IsOptional, MinLength, MaxLength, IsArray, ArrayMinSize } from 'class-validator';

export enum UserRole {
  ADMIN = 'ADMIN',
  SURVEYOR = 'SURVEYOR',
}

export enum UserStatus {
  ACTIVE = 'ACTIVE',
  INACTIVE = 'INACTIVE',
  SUSPENDED = 'SUSPENDED',
}

export class LoginDto {
  @IsString()
  @MinLength(3)
  @MaxLength(50)
  username: string;

  @IsString()
  @MinLength(6)
  password: string;
}

export class RegisterDto {
  @IsString()
  @MinLength(3)
  @MaxLength(50)
  username: string;

  @IsEmail()
  email: string;

  @IsString()
  @MinLength(6)
  password: string;

  @IsString()
  @MinLength(2)
  @MaxLength(100)
  name: string;

  @IsEnum(UserRole)
  role: UserRole;
}

export class CreateUserDto {
  @IsString()
  @MinLength(3)
  @MaxLength(50)
  username: string;

  @IsEmail()
  email: string;

  @IsString()
  @MinLength(6)
  password: string;

  @IsString()
  @MinLength(2)
  @MaxLength(100)
  name: string;

  @IsEnum(UserRole)
  role: UserRole;

  // Surveyor area configuration (optional)
  @IsOptional()
  @IsArray()
  @IsString({ each: true })
  surveyorAreas?: string[];

  @IsOptional()
  @IsString()
  @MaxLength(20)
  surveyorLocation?: string;

  @IsOptional()
  @IsString()
  @MaxLength(50)
  maxTravelTime?: string;
}

export class UpdateUserDto {
  @IsOptional()
  @IsString()
  @MinLength(2)
  @MaxLength(100)
  name?: string;

  @IsOptional()
  @IsEmail()
  email?: string;

  @IsOptional()
  @IsEnum(UserRole)
  role?: UserRole;

  @IsOptional()
  @IsEnum(UserStatus)
  status?: UserStatus;

  // Surveyor area configuration (optional)
  @IsOptional()
  @IsArray()
  @IsString({ each: true })
  surveyorAreas?: string[];

  @IsOptional()
  @IsString()
  @MaxLength(20)
  surveyorLocation?: string;

  @IsOptional()
  @IsString()
  @MaxLength(50)
  maxTravelTime?: string;
}

export class AdminUpdateUserDto {
  @IsOptional()
  @IsString()
  @MinLength(2)
  @MaxLength(100)
  name?: string;

  @IsOptional()
  @IsEmail()
  email?: string;

  @IsOptional()
  @IsEnum(UserRole)
  role?: UserRole;

  @IsOptional()
  @IsEnum(UserStatus)
  status?: UserStatus;

  @IsOptional()
  @IsString()
  @MinLength(3)
  @MaxLength(50)
  username?: string;

  // Surveyor area configuration (optional)
  @IsOptional()
  @IsArray()
  @IsString({ each: true })
  surveyorAreas?: string[];

  @IsOptional()
  @IsString()
  @MaxLength(20)
  surveyorLocation?: string;

  @IsOptional()
  @IsString()
  @MaxLength(50)
  maxTravelTime?: string;
}

export class AdminResetPasswordDto {
  @IsString()
  @MinLength(6)
  newPassword: string;
}

export class ChangePasswordDto {
  @IsString()
  @MinLength(6)
  currentPassword: string;

  @IsString()
  @MinLength(6)
  newPassword: string;
}

export class ResetPasswordDto {
  @IsEmail()
  email: string;
}

export class ResetPasswordConfirmDto {
  @IsString()
  token: string;

  @IsString()
  @MinLength(6)
  newPassword: string;
}

export class AuthResponseDto {
  accessToken: string;
  refreshToken: string;
  user: {
    id: string;
    username: string;
    email: string;
    name: string;
    role: UserRole;
    status: UserStatus;
  };
}

export interface GHLAssignmentResult {
  success: boolean;
  ghlUserId?: string;
  ghlUserName?: string;
  message: string;
  requiresManualAssignment?: boolean;
}

export class UserResponseDto {
  id: string;
  username: string;
  email: string;
  name: string;
  role: UserRole;
  status: UserStatus;
  isEmailVerified: boolean;
  ghlUserId?: string;
  ghlUserName?: string;
  surveyorAreas?: string[];
  surveyorLocation?: string;
  maxTravelTime?: string;
  createdAt: Date;
  updatedAt: Date;
  lastLoginAt?: Date;
  ghlAssignment?: GHLAssignmentResult;
}

export class UserListResponseDto {
  users: UserResponseDto[];
  total: number;
  page: number;
  limit: number;
} 