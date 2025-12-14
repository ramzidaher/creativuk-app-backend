import { Injectable, UnauthorizedException, BadRequestException, NotFoundException, Logger } from '@nestjs/common';
import { JwtService } from '@nestjs/jwt';
import { PrismaService } from '../prisma/prisma.service';
import { UserService } from '../user/user.service';
import { GHLUserLookupService } from '../integrations/ghl-user-lookup.service';
import * as bcrypt from 'bcrypt';
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
  UserListResponseDto,
  UserRole,
  UserStatus,
  GHLAssignmentResult
} from './dto/auth.dto';
import { User } from '@prisma/client';
import { ConfigService } from '@nestjs/config';
import * as crypto from 'crypto';

@Injectable()
export class AuthService {
  private readonly logger = new Logger(AuthService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly userService: UserService,
    private readonly jwtService: JwtService,
    private readonly configService: ConfigService,
    private readonly ghlUserLookupService: GHLUserLookupService,
  ) {}

  async login(loginDto: LoginDto): Promise<AuthResponseDto> {
    const user = await this.prisma.user.findUnique({
      where: { username: loginDto.username }
    });

    if (!user || user.status !== UserStatus.ACTIVE) {
      throw new UnauthorizedException('Invalid credentials');
    }

    const isPasswordValid = await bcrypt.compare(loginDto.password, user.password);
    if (!isPasswordValid) {
      throw new UnauthorizedException('Invalid credentials');
    }

    // Update last login
    await this.prisma.user.update({
      where: { id: user.id },
      data: { lastLoginAt: new Date() }
    });

    const tokens = await this.generateTokens(user);
    return {
      ...tokens,
      user: this.mapToUserResponse(user)
    };
  }

  async register(registerDto: RegisterDto): Promise<AuthResponseDto> {
    // Check if username or email already exists
    const existingUser = await this.prisma.user.findFirst({
      where: {
        OR: [
          { username: registerDto.username },
          { email: registerDto.email }
        ]
      }
    });

    if (existingUser) {
      throw new BadRequestException('Username or email already exists');
    }

    // Hash password
    const hashedPassword = await bcrypt.hash(registerDto.password, 10);

    // Create user
    const user = await this.prisma.user.create({
      data: {
        username: registerDto.username,
        email: registerDto.email,
        password: hashedPassword,
        name: registerDto.name,
        role: registerDto.role,
        status: UserStatus.ACTIVE,
        isEmailVerified: false,
      }
    });

    const tokens = await this.generateTokens(user);
    return {
      ...tokens,
      user: this.mapToUserResponse(user)
    };
  }

  async createUser(createUserDto: CreateUserDto, adminUser: User): Promise<UserResponseDto> {
    // Only admins can create users
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can create users');
    }

    // Additional security: Prevent admin from creating another admin (optional security measure)
    if (createUserDto.role === UserRole.ADMIN) {
      this.logger.warn(`Admin ${adminUser.username} attempting to create another admin user`);
      // Uncomment the line below if you want to prevent admin creation via API
      // throw new BadRequestException('Cannot create admin users via API for security reasons');
    }

    // Enhanced validation
    if (!createUserDto.username || !createUserDto.email || !createUserDto.password || !createUserDto.name) {
      throw new BadRequestException('All fields are required');
    }

    // Username validation
    if (createUserDto.username.length < 3 || createUserDto.username.length > 50) {
      throw new BadRequestException('Username must be between 3 and 50 characters');
    }

    // Email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(createUserDto.email)) {
      throw new BadRequestException('Invalid email format');
    }

    // Password validation
    if (createUserDto.password.length < 8) {
      throw new BadRequestException('Password must be at least 8 characters long');
    }

    // Name validation
    if (createUserDto.name.length < 2 || createUserDto.name.length > 100) {
      throw new BadRequestException('Name must be between 2 and 100 characters');
    }

    // Check if username or email already exists
    const existingUser = await this.prisma.user.findFirst({
      where: {
        OR: [
          { username: createUserDto.username.toLowerCase() },
          { email: createUserDto.email.toLowerCase() }
        ]
      }
    });

    if (existingUser) {
      throw new BadRequestException('Username or email already exists');
    }

    // Hash password with higher salt rounds for better security
    const hashedPassword = await bcrypt.hash(createUserDto.password, 12);

    // Create user with sanitized data
    const user = await this.prisma.user.create({
      data: {
        username: createUserDto.username.toLowerCase().trim(),
        email: createUserDto.email.toLowerCase().trim(),
        password: hashedPassword,
        name: createUserDto.name.trim(),
        role: createUserDto.role,
        status: UserStatus.ACTIVE,
        isEmailVerified: false,
        // Surveyor area configuration
        surveyorAreas: createUserDto.surveyorAreas || [],
        surveyorLocation: createUserDto.surveyorLocation,
        maxTravelTime: createUserDto.maxTravelTime,
      }
    });

    // Try to assign GHL user ID automatically
    let ghlAssignmentResult: GHLAssignmentResult | undefined = undefined;
    try {
      this.logger.log(`Attempting to assign GHL user ID for new user: ${user.username} (${user.name})`);
      
      const ghlLookupResult = await this.ghlUserLookupService.findGHLUser(user.name || '', user.email);
      
      if (ghlLookupResult.found && ghlLookupResult.ghlUser) {
        // Update user with GHL user ID
        const updatedUser = await this.prisma.user.update({
          where: { id: user.id },
          data: { ghlUserId: ghlLookupResult.ghlUser.id }
        });
        
        ghlAssignmentResult = {
          success: true,
          ghlUserId: ghlLookupResult.ghlUser.id,
          ghlUserName: `${ghlLookupResult.ghlUser.firstName} ${ghlLookupResult.ghlUser.lastName}`,
          message: ghlLookupResult.message
        };
        
        this.logger.log(`✅ Successfully assigned GHL user ID ${ghlLookupResult.ghlUser.id} to user ${user.username}`);
      } else {
        ghlAssignmentResult = {
          success: false,
          message: ghlLookupResult.message,
          requiresManualAssignment: true
        };
        
        this.logger.warn(`⚠️  Could not automatically assign GHL user ID to ${user.username}: ${ghlLookupResult.message}`);
        this.logger.warn(`   Manual assignment required. Use POST /user/assign-ghl-id/${user.id} to assign manually.`);
      }
    } catch (error) {
      ghlAssignmentResult = {
        success: false,
        message: `GHL lookup failed: ${error.message}`,
        requiresManualAssignment: true
      };
      
      this.logger.error(`❌ Error during GHL user lookup for ${user.username}: ${error.message}`);
    }

    this.logger.log(`Admin ${adminUser.username} created new user: ${user.username} (${user.role})`);
    
    // Return user response with GHL assignment info
    const userResponse = this.mapToUserResponse(user);
    return {
      ...userResponse,
      ghlAssignment: ghlAssignmentResult
    };
  }

  async updateUser(userId: string, updateUserDto: UpdateUserDto, adminUser: User): Promise<UserResponseDto> {
    // Only admins can update users, or users can update themselves
    if (adminUser.role !== UserRole.ADMIN && adminUser.id !== userId) {
      throw new UnauthorizedException('Unauthorized to update this user');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    // Update user
    const updatedUser = await this.prisma.user.update({
      where: { id: userId },
      data: {
        ...(updateUserDto.name !== undefined && { name: updateUserDto.name }),
        ...(updateUserDto.email !== undefined && { email: updateUserDto.email }),
        ...(updateUserDto.role !== undefined && { role: updateUserDto.role }),
        ...(updateUserDto.status !== undefined && { status: updateUserDto.status }),
        // Surveyor area configuration
        ...(updateUserDto.surveyorAreas !== undefined && { surveyorAreas: updateUserDto.surveyorAreas }),
        ...(updateUserDto.surveyorLocation !== undefined && { surveyorLocation: updateUserDto.surveyorLocation }),
        ...(updateUserDto.maxTravelTime !== undefined && { maxTravelTime: updateUserDto.maxTravelTime }),
      }
    });

    return this.mapToUserResponse(updatedUser);
  }

  async changePassword(userId: string, changePasswordDto: ChangePasswordDto): Promise<void> {
    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    // Verify current password
    const isCurrentPasswordValid = await bcrypt.compare(changePasswordDto.currentPassword, user.password);
    if (!isCurrentPasswordValid) {
      throw new BadRequestException('Current password is incorrect');
    }

    // Hash new password
    const hashedNewPassword = await bcrypt.hash(changePasswordDto.newPassword, 10);

    // Update password
    await this.prisma.user.update({
      where: { id: userId },
      data: { password: hashedNewPassword }
    });
  }

  async resetPassword(resetPasswordDto: ResetPasswordDto): Promise<void> {
    const user = await this.prisma.user.findUnique({
      where: { email: resetPasswordDto.email }
    });

    if (!user) {
      // Don't reveal if user exists or not
      return;
    }

    // Generate reset token
    const resetToken = crypto.randomBytes(32).toString('hex');
    const resetTokenExpiry = new Date(Date.now() + 24 * 60 * 60 * 1000); // 24 hours

    await this.prisma.user.update({
      where: { id: user.id },
      data: {
        passwordResetToken: resetToken,
        passwordResetExpires: resetTokenExpiry
      }
    });

    // TODO: Send email with reset token
    console.log(`Password reset token for ${user.email}: ${resetToken}`);
  }

  async resetPasswordConfirm(resetPasswordConfirmDto: ResetPasswordConfirmDto): Promise<void> {
    const user = await this.prisma.user.findFirst({
      where: {
        passwordResetToken: resetPasswordConfirmDto.token,
        passwordResetExpires: {
          gt: new Date()
        }
      }
    });

    if (!user) {
      throw new BadRequestException('Invalid or expired reset token');
    }

    // Hash new password
    const hashedPassword = await bcrypt.hash(resetPasswordConfirmDto.newPassword, 10);

    // Update password and clear reset token
    await this.prisma.user.update({
      where: { id: user.id },
      data: {
        password: hashedPassword,
        passwordResetToken: null,
        passwordResetExpires: null
      }
    });
  }

  async getAllUsers(adminUser: User, options?: {
    page?: number;
    limit?: number;
    search?: string;
    role?: string;
    status?: string;
  }): Promise<UserListResponseDto> {
    // Only admins can view all users
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can view all users');
    }

    const page = options?.page || 1;
    const limit = options?.limit || 10;
    const skip = (page - 1) * limit;

    // Build where clause
    const where: any = {};
    
    if (options?.search) {
      where.OR = [
        { name: { contains: options.search, mode: 'insensitive' } },
        { email: { contains: options.search, mode: 'insensitive' } },
        { username: { contains: options.search, mode: 'insensitive' } }
      ];
    }

    if (options?.role) {
      where.role = options.role;
    }

    if (options?.status) {
      where.status = options.status;
    }

    const [users, total] = await Promise.all([
      this.prisma.user.findMany({
        where,
        skip,
        take: limit,
      orderBy: { createdAt: 'desc' }
      }),
      this.prisma.user.count({ where })
    ]);

    return {
      users: users.map(user => this.mapToUserResponse(user)),
      total,
      page,
      limit
    };
  }

  async getUserById(userId: string, requestingUser: User): Promise<UserResponseDto> {
    // Users can view their own profile, admins can view any profile
    if (requestingUser.role !== UserRole.ADMIN && requestingUser.id !== userId) {
      throw new UnauthorizedException('Unauthorized to view this user');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    return this.mapToUserResponse(user);
  }

  async deleteUser(userId: string, adminUser: User): Promise<void> {
    // Only admins can delete users
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can delete users');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    // Prevent admin from deleting themselves
    if (adminUser.id === userId) {
      throw new BadRequestException('Cannot delete your own account');
    }

    // Use a transaction to ensure all related data is cleaned up properly
    // Increase timeout to 60 seconds to handle large deletions
    await this.prisma.$transaction(async (tx) => {
      // 1. Delete OpportunityProgress records (this will cascade to OpportunityStep)
      await tx.opportunityProgress.deleteMany({
        where: { userId: userId }
      });

      // 2. Delete Appointment records
      await tx.appointment.deleteMany({
        where: { userId: userId }
      });

      // 3. Handle Survey records that reference this user
      // Since createdBy is required, we need to either delete surveys or transfer ownership
      // For now, we'll delete surveys created by this user to avoid orphaned records
      await tx.survey.deleteMany({
        where: { createdBy: userId }
      });

      // Set updatedBy to null for surveys updated by this user (this field is optional)
      await tx.survey.updateMany({
        where: { updatedBy: userId },
        data: { updatedBy: null }
      });

      // 4. Delete AutoSave records (these have CASCADE, but being explicit)
      await tx.autoSave.deleteMany({
        where: { userId: userId }
      });

      // 5. Delete CalculatorProgress records (these have CASCADE, but being explicit)
      await tx.calculatorProgress.deleteMany({
        where: { userId: userId }
      });

      // 6. Delete OpportunityOutcome records
      await tx.opportunityOutcome.deleteMany({
        where: { userId: userId }
      });

      // 7. Finally, delete the user
      await tx.user.delete({
        where: { id: userId }
      });
    }, {
      maxWait: 60000, // Maximum time to wait for a transaction slot (60 seconds)
      timeout: 60000, // Maximum time the transaction can run (60 seconds)
    });
  }

  private async generateTokens(user: User): Promise<{ accessToken: string; refreshToken: string }> {
    const payload = {
      sub: user.id,
      username: user.username,
      email: user.email,
      name: user.name,
      role: user.role,
      ghlUserId: user.ghlUserId,
    };

    const [accessToken, refreshToken] = await Promise.all([
      this.jwtService.signAsync(payload, {
        secret: this.configService.get<string>('JWT_SECRET'),
        expiresIn: '1d',
      }),
      this.jwtService.signAsync(payload, {
        secret: this.configService.get<string>('JWT_REFRESH_SECRET'),
        expiresIn: '7d',
      }),
    ]);

    return { accessToken, refreshToken };
  }

  private mapToUserResponse(user: User): UserResponseDto {
    return {
      id: user.id,
      username: user.username,
      email: user.email,
      name: user.name || '',
      role: user.role as UserRole,
      status: user.status as UserStatus,
      isEmailVerified: user.isEmailVerified,
      ghlUserId: user.ghlUserId || undefined,
      ghlUserName: (user as any).ghlUserName || undefined,
      surveyorAreas: (user as any).surveyorAreas || undefined,
      surveyorLocation: (user as any).surveyorLocation || undefined,
      maxTravelTime: (user as any).maxTravelTime || undefined,
      createdAt: user.createdAt,
      updatedAt: user.updatedAt,
      lastLoginAt: user.lastLoginAt || undefined,
    };
  }

  // Admin user management methods
  async updateAdminUser(userId: string, updateUserDto: AdminUpdateUserDto, adminUser: User): Promise<UserResponseDto> {
    // Only admins can update users
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can update users');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    // Check for unique constraints
    if (updateUserDto.username || updateUserDto.email) {
      const existingUser = await this.prisma.user.findFirst({
        where: {
          OR: [
            ...(updateUserDto.username ? [{ username: updateUserDto.username }] : []),
            ...(updateUserDto.email ? [{ email: updateUserDto.email }] : [])
          ],
          NOT: { id: userId }
        }
      });

      if (existingUser) {
        throw new BadRequestException('Username or email already exists');
      }
    }

    // Determine if role is being changed from SURVEYOR to ADMIN
    const isChangingToAdmin = updateUserDto.role === UserRole.ADMIN && user.role === UserRole.SURVEYOR;
    
    // Prepare update data
    const updateData: any = {
      ...(updateUserDto.name && { name: updateUserDto.name }),
      ...(updateUserDto.email && { email: updateUserDto.email }),
      ...(updateUserDto.username && { username: updateUserDto.username }),
      ...(updateUserDto.role !== undefined && { role: updateUserDto.role }),
      ...(updateUserDto.status && { status: updateUserDto.status }),
      updatedAt: new Date()
    };

    // Handle surveyor area configuration
    if (updateUserDto.surveyorAreas !== undefined) {
      updateData.surveyorAreas = updateUserDto.surveyorAreas;
    }
    if (updateUserDto.surveyorLocation !== undefined) {
      updateData.surveyorLocation = updateUserDto.surveyorLocation;
    }
    if (updateUserDto.maxTravelTime !== undefined) {
      updateData.maxTravelTime = updateUserDto.maxTravelTime;
    }

    // When changing from SURVEYOR to ADMIN, clear surveyor-specific fields if not explicitly provided
    if (isChangingToAdmin) {
      if (updateUserDto.surveyorAreas === undefined) {
        updateData.surveyorAreas = [];
      }
      if (updateUserDto.surveyorLocation === undefined) {
        updateData.surveyorLocation = null;
      }
    }

    const updatedUser = await this.prisma.user.update({
      where: { id: userId },
      data: updateData
    });

    return this.mapToUserResponse(updatedUser);
  }

  async adminResetPassword(userId: string, resetPasswordDto: AdminResetPasswordDto, adminUser: User): Promise<void> {
    // Only admins can reset passwords
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can reset passwords');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    // Hash new password
    const hashedPassword = await bcrypt.hash(resetPasswordDto.newPassword, 10);

    await this.prisma.user.update({
      where: { id: userId },
      data: {
        password: hashedPassword,
        updatedAt: new Date()
      }
    });
  }

  async activateUser(userId: string, adminUser: User): Promise<void> {
    // Only admins can activate users
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can activate users');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    await this.prisma.user.update({
      where: { id: userId },
      data: {
        status: UserStatus.ACTIVE,
        updatedAt: new Date()
      }
    });
  }

  async deactivateUser(userId: string, adminUser: User): Promise<void> {
    // Only admins can deactivate users
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can deactivate users');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    await this.prisma.user.update({
      where: { id: userId },
      data: {
        status: UserStatus.INACTIVE,
        updatedAt: new Date()
      }
    });
  }

  async suspendUser(userId: string, adminUser: User): Promise<void> {
    // Only admins can suspend users
    if (adminUser.role !== UserRole.ADMIN) {
      throw new UnauthorizedException('Only admins can suspend users');
    }

    const user = await this.prisma.user.findUnique({
      where: { id: userId }
    });

    if (!user) {
      throw new NotFoundException('User not found');
    }

    await this.prisma.user.update({
      where: { id: userId },
      data: {
        status: UserStatus.SUSPENDED,
        updatedAt: new Date()
      }
    });
  }

  async createUserFromGhl(ghlUserId: string, name?: string, email?: string): Promise<User> {
    this.logger.log(`Creating user from GHL data: ${ghlUserId}`);
    
    return this.userService.upsertByGhlUserId({
      ghlUserId,
      name: name || `User ${ghlUserId}`,
      email: email || `user-${ghlUserId}@temp.com`,
    });
  }
} 