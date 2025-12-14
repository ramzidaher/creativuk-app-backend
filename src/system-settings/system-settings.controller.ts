import { Controller, Get, Post, Put, Delete, Body, Param, UseGuards, Query } from '@nestjs/common';
import { SystemSettingsService, SystemSettingDto, CreateSystemSettingDto, UpdateSystemSettingDto } from './system-settings.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { RolesGuard } from '../auth/roles.guard';
import { Roles } from '../auth/roles.decorator';
import { UserRole } from '../auth/dto/auth.dto';

@Controller('system-settings')
@UseGuards(JwtAuthGuard)
export class SystemSettingsController {
  constructor(private readonly systemSettingsService: SystemSettingsService) {}

  /**
   * Get all system settings (admin only)
   */
  @Get()
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getAllSettings(): Promise<{ success: boolean; data: SystemSettingDto[]; error?: string }> {
    try {
      const settings = await this.systemSettingsService.getAllSettings();
      return {
        success: true,
        data: settings
      };
    } catch (error) {
      return {
        success: false,
        data: [],
        error: error.message
      };
    }
  }

  /**
   * Get settings by category (admin only)
   */
  @Get('category/:category')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getSettingsByCategory(@Param('category') category: string): Promise<{ success: boolean; data: SystemSettingDto[]; error?: string }> {
    try {
      const settings = await this.systemSettingsService.getSettingsByCategory(category);
      return {
        success: true,
        data: settings
      };
    } catch (error) {
      return {
        success: false,
        data: [],
        error: error.message
      };
    }
  }

  /**
   * Get public settings (accessible to all authenticated users)
   */
  @Get('public')
  async getPublicSettings(): Promise<{ success: boolean; data: SystemSettingDto[]; error?: string }> {
    try {
      const settings = await this.systemSettingsService.getPublicSettings();
      return {
        success: true,
        data: settings
      };
    } catch (error) {
      return {
        success: false,
        data: [],
        error: error.message
      };
    }
  }

  /**
   * Get a specific setting by key (admin only)
   */
  @Get(':key')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getSettingByKey(@Param('key') key: string): Promise<{ success: boolean; data: SystemSettingDto | null; error?: string }> {
    try {
      const setting = await this.systemSettingsService.getSettingByKey(key);
      return {
        success: true,
        data: setting
      };
    } catch (error) {
      return {
        success: false,
        data: null,
        error: error.message
      };
    }
  }

  /**
   * Get a setting value by key (public endpoint for non-admin users to access public settings)
   */
  @Get('value/:key')
  async getSettingValue(@Param('key') key: string): Promise<{ success: boolean; data: string | null; error?: string }> {
    try {
      // First check if the setting is public
      const setting = await this.systemSettingsService.getSettingByKey(key);
      if (!setting) {
        return {
          success: false,
          data: null,
          error: 'Setting not found'
        };
      }

      // If setting is not public, only allow admin access
      if (!setting.isPublic) {
        return {
          success: false,
          data: null,
          error: 'Access denied'
        };
      }

      return {
        success: true,
        data: setting.value
      };
    } catch (error) {
      return {
        success: false,
        data: null,
        error: error.message
      };
    }
  }

  /**
   * Create a new system setting (admin only)
   */
  @Post()
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async createSetting(@Body() dto: CreateSystemSettingDto): Promise<{ success: boolean; data: SystemSettingDto | null; error?: string }> {
    try {
      const setting = await this.systemSettingsService.createSetting(dto);
      return {
        success: true,
        data: setting
      };
    } catch (error) {
      return {
        success: false,
        data: null,
        error: error.message
      };
    }
  }

  /**
   * Update a system setting (admin only)
   */
  @Put(':key')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async updateSetting(
    @Param('key') key: string,
    @Body() dto: UpdateSystemSettingDto
  ): Promise<{ success: boolean; data: SystemSettingDto | null; error?: string }> {
    try {
      const setting = await this.systemSettingsService.updateSetting(key, dto);
      return {
        success: true,
        data: setting
      };
    } catch (error) {
      return {
        success: false,
        data: null,
        error: error.message
      };
    }
  }

  /**
   * Update or create a system setting (admin only)
   */
  @Post('upsert')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async upsertSetting(@Body() dto: CreateSystemSettingDto): Promise<{ success: boolean; data: SystemSettingDto | null; error?: string }> {
    try {
      const setting = await this.systemSettingsService.upsertSetting(dto);
      return {
        success: true,
        data: setting
      };
    } catch (error) {
      return {
        success: false,
        data: null,
        error: error.message
      };
    }
  }

  /**
   * Delete a system setting (admin only)
   */
  @Delete(':key')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async deleteSetting(@Param('key') key: string): Promise<{ success: boolean; error?: string }> {
    try {
      await this.systemSettingsService.deleteSetting(key);
      return {
        success: true
      };
    } catch (error) {
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Initialize default settings (admin only)
   */
  @Post('initialize-defaults')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async initializeDefaultSettings(): Promise<{ success: boolean; error?: string }> {
    try {
      await this.systemSettingsService.initializeDefaultSettings();
      return {
        success: true
      };
    } catch (error) {
      return {
        success: false,
        error: error.message
      };
    }
  }
}
