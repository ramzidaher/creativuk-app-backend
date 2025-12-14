import { 
  Controller, 
  Post, 
  Get, 
  Put, 
  Delete, 
  Body, 
  Param, 
  UseGuards, 
  Request,
  HttpCode,
  HttpStatus,
  Query
} from '@nestjs/common';
import { AuthService } from './auth.service';
import { 
  LoginDto, 
  RegisterDto, 
  CreateUserDto, 
  UpdateUserDto, 
  AdminUpdateUserDto,
  AdminResetPasswordDto,
  ChangePasswordDto,
  ResetPasswordDto,
  ResetPasswordConfirmDto,
  AuthResponseDto,
  UserResponseDto,
  UserListResponseDto
} from './dto/auth.dto';
import { JwtAuthGuard } from './jwt-auth.guard';
import { RolesGuard } from './roles.guard';
import { Roles } from './roles.decorator';
import { UserRole } from './dto/auth.dto';

@Controller('auth')
export class AuthController {
  constructor(private readonly authService: AuthService) {}

  @Post('login')
  @HttpCode(HttpStatus.OK)
  async login(@Body() loginDto: LoginDto): Promise<AuthResponseDto> {
    return this.authService.login(loginDto);
  }

  @Post('register')
  async register(@Body() registerDto: RegisterDto): Promise<AuthResponseDto> {
    return this.authService.register(registerDto);
  }

  @Post('reset-password')
  @HttpCode(HttpStatus.OK)
  async resetPassword(@Body() resetPasswordDto: ResetPasswordDto): Promise<{ message: string }> {
    await this.authService.resetPassword(resetPasswordDto);
    return { message: 'If the email exists, a reset link has been sent' };
  }

  @Post('reset-password/confirm')
  @HttpCode(HttpStatus.OK)
  async resetPasswordConfirm(@Body() resetPasswordConfirmDto: ResetPasswordConfirmDto): Promise<{ message: string }> {
    await this.authService.resetPasswordConfirm(resetPasswordConfirmDto);
    return { message: 'Password reset successfully' };
  }

  @UseGuards(JwtAuthGuard)
  @Get('profile')
  async getProfile(@Request() req): Promise<UserResponseDto> {
    return this.authService.getUserById(req.user.sub, req.user);
  }

  @Post('create-missing-user')
  @HttpCode(HttpStatus.OK)
  async createMissingUser(@Body() body: { ghlUserId: string; name?: string; email?: string }): Promise<{ success: boolean; message: string }> {
    try {
      const user = await this.authService.createUserFromGhl(body.ghlUserId, body.name, body.email);
      return { 
        success: true, 
        message: `User created successfully: ${user.username || user.name}` 
      };
    } catch (error) {
      return { 
        success: false, 
        message: `Failed to create user: ${error.message}` 
      };
    }
  }

  @UseGuards(JwtAuthGuard)
  @Put('change-password')
  @HttpCode(HttpStatus.OK)
  async changePassword(
    @Request() req,
    @Body() changePasswordDto: ChangePasswordDto
  ): Promise<{ message: string }> {
    await this.authService.changePassword(req.user.sub, changePasswordDto);
    return { message: 'Password changed successfully' };
  }

  // Admin-only endpoints
  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Post('users')
  async createUser(
    @Request() req,
    @Body() createUserDto: CreateUserDto
  ): Promise<UserResponseDto> {
    return this.authService.createUser(createUserDto, req.user);
  }

  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Get('users')
  async getAllUsers(
    @Request() req,
    @Query('page') page: string = '1',
    @Query('limit') limit: string = '10',
    @Query('search') search?: string,
    @Query('role') role?: string,
    @Query('status') status?: string
  ): Promise<UserListResponseDto> {
    return this.authService.getAllUsers(req.user, {
      page: parseInt(page),
      limit: parseInt(limit),
      search,
      role,
      status
    });
  }

  @UseGuards(JwtAuthGuard)
  @Get('users/:id')
  async getUserById(
    @Request() req,
    @Param('id') userId: string
  ): Promise<UserResponseDto> {
    return this.authService.getUserById(userId, req.user);
  }

  @UseGuards(JwtAuthGuard)
  @Put('users/:id')
  async updateUser(
    @Request() req,
    @Param('id') userId: string,
    @Body() updateUserDto: UpdateUserDto
  ): Promise<UserResponseDto> {
    return this.authService.updateUser(userId, updateUserDto, req.user);
  }

  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Delete('users/:id')
  async deleteUser(
    @Request() req,
    @Param('id') userId: string
  ): Promise<{ success: boolean; message: string }> {
    await this.authService.deleteUser(userId, req.user);
    return { success: true, message: 'User deleted successfully' };
  }

  // Admin user management endpoints
  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Get('admin/users/:id')
  async getAdminUserById(
    @Request() req,
    @Param('id') userId: string
  ): Promise<UserResponseDto> {
    return this.authService.getUserById(userId, req.user);
  }

  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Put('admin/users/:id')
  async updateAdminUser(
    @Request() req,
    @Param('id') userId: string,
    @Body() updateUserDto: AdminUpdateUserDto
  ): Promise<UserResponseDto> {
    return this.authService.updateAdminUser(userId, updateUserDto, req.user);
  }

  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Post('admin/users/:id/reset-password')
  @HttpCode(HttpStatus.OK)
  async adminResetPassword(
    @Request() req,
    @Param('id') userId: string,
    @Body() resetPasswordDto: AdminResetPasswordDto
  ): Promise<{ message: string }> {
    await this.authService.adminResetPassword(userId, resetPasswordDto, req.user);
    return { message: 'Password reset successfully' };
  }

  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Post('admin/users/:id/activate')
  @HttpCode(HttpStatus.OK)
  async activateUser(
    @Request() req,
    @Param('id') userId: string
  ): Promise<{ message: string }> {
    await this.authService.activateUser(userId, req.user);
    return { message: 'User activated successfully' };
  }

  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Post('admin/users/:id/deactivate')
  @HttpCode(HttpStatus.OK)
  async deactivateUser(
    @Request() req,
    @Param('id') userId: string
  ): Promise<{ message: string }> {
    await this.authService.deactivateUser(userId, req.user);
    return { message: 'User deactivated successfully' };
  }

  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  @Post('admin/users/:id/suspend')
  @HttpCode(HttpStatus.OK)
  async suspendUser(
    @Request() req,
    @Param('id') userId: string
  ): Promise<{ message: string }> {
    await this.authService.suspendUser(userId, req.user);
    return { message: 'User suspended successfully' };
  }
} 