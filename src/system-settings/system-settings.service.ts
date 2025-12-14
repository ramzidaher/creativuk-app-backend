import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';

export interface SystemSettingDto {
  id: string;
  key: string;
  value: string;
  description?: string;
  category: string;
  isPublic: boolean;
  createdAt: Date;
  updatedAt: Date;
}

export interface CreateSystemSettingDto {
  key: string;
  value: string;
  description?: string;
  category?: string;
  isPublic?: boolean;
}

export interface UpdateSystemSettingDto {
  value?: string;
  description?: string;
  category?: string;
  isPublic?: boolean;
}

@Injectable()
export class SystemSettingsService {
  private readonly logger = new Logger(SystemSettingsService.name);

  constructor(private readonly prisma: PrismaService) {}

  /**
   * Get all system settings
   */
  async getAllSettings(): Promise<SystemSettingDto[]> {
    this.logger.log('Getting all system settings');
    
    const settings = await this.prisma.systemSettings.findMany({
      orderBy: { key: 'asc' }
    });

    return settings.map(this.mapToDto);
  }

  /**
   * Get settings by category
   */
  async getSettingsByCategory(category: string): Promise<SystemSettingDto[]> {
    this.logger.log(`Getting system settings for category: ${category}`);
    
    const settings = await this.prisma.systemSettings.findMany({
      where: { category },
      orderBy: { key: 'asc' }
    });

    return settings.map(this.mapToDto);
  }

  /**
   * Get public settings (accessible to non-admin users)
   */
  async getPublicSettings(): Promise<SystemSettingDto[]> {
    this.logger.log('Getting public system settings');
    
    const settings = await this.prisma.systemSettings.findMany({
      where: { isPublic: true },
      orderBy: { key: 'asc' }
    });

    return settings.map(this.mapToDto);
  }

  /**
   * Get a specific setting by key
   */
  async getSettingByKey(key: string): Promise<SystemSettingDto | null> {
    this.logger.log(`Getting system setting: ${key}`);
    
    const setting = await this.prisma.systemSettings.findUnique({
      where: { key }
    });

    return setting ? this.mapToDto(setting) : null;
  }

  /**
   * Get a setting value by key (returns just the value)
   */
  async getSettingValue(key: string): Promise<string | null> {
    this.logger.log(`Getting system setting value: ${key}`);
    
    const setting = await this.prisma.systemSettings.findUnique({
      where: { key },
      select: { value: true }
    });

    const value = setting?.value || null;
    this.logger.log(`Retrieved setting value for ${key}: ${value}`);
    return value;
  }

  /**
   * Get a boolean setting value by key
   */
  async getBooleanSetting(key: string, defaultValue: boolean = false): Promise<boolean> {
    const value = await this.getSettingValue(key);
    if (value === null) {
      return defaultValue;
    }
    
    try {
      return JSON.parse(value);
    } catch {
      return defaultValue;
    }
  }

  /**
   * Create a new system setting
   */
  async createSetting(dto: CreateSystemSettingDto): Promise<SystemSettingDto> {
    this.logger.log(`Creating system setting: ${dto.key}`);
    
    const setting = await this.prisma.systemSettings.create({
      data: {
        key: dto.key,
        value: dto.value,
        description: dto.description,
        category: dto.category || 'general',
        isPublic: dto.isPublic || false
      }
    });

    return this.mapToDto(setting);
  }

  /**
   * Update a system setting
   */
  async updateSetting(key: string, dto: UpdateSystemSettingDto): Promise<SystemSettingDto> {
    this.logger.log(`Updating system setting: ${key}`);
    
    const setting = await this.prisma.systemSettings.update({
      where: { key },
      data: {
        ...(dto.value !== undefined && { value: dto.value }),
        ...(dto.description !== undefined && { description: dto.description }),
        ...(dto.category !== undefined && { category: dto.category }),
        ...(dto.isPublic !== undefined && { isPublic: dto.isPublic })
      }
    });

    return this.mapToDto(setting);
  }

  /**
   * Update or create a system setting (upsert)
   */
  async upsertSetting(dto: CreateSystemSettingDto): Promise<SystemSettingDto> {
    this.logger.log(`Upserting system setting: ${dto.key} with value: ${dto.value}`);
    
    const setting = await this.prisma.systemSettings.upsert({
      where: { key: dto.key },
      update: {
        value: dto.value,
        description: dto.description,
        category: dto.category || 'general',
        isPublic: dto.isPublic || false
      },
      create: {
        key: dto.key,
        value: dto.value,
        description: dto.description,
        category: dto.category || 'general',
        isPublic: dto.isPublic || false
      }
    });

    this.logger.log(`Successfully upserted setting: ${dto.key}, new value: ${setting.value}`);
    return this.mapToDto(setting);
  }

  /**
   * Delete a system setting
   */
  async deleteSetting(key: string): Promise<void> {
    this.logger.log(`Deleting system setting: ${key}`);
    
    await this.prisma.systemSettings.delete({
      where: { key }
    });
  }

  /**
   * Initialize default system settings
   */
  async initializeDefaultSettings(): Promise<void> {
    this.logger.log('Initializing default system settings');
    
    const defaultSettings = [
      {
        key: 'step_navigation_enabled',
        value: JSON.stringify(true),
        description: 'Allow users to navigate to any workflow step regardless of completion status',
        category: 'workflow',
        isPublic: true
      },
      {
        key: 'survey_autosave_enabled',
        value: JSON.stringify(true),
        description: 'Enable automatic saving of survey progress',
        category: 'survey',
        isPublic: true
      },
      {
        key: 'presentation_generation_enabled',
        value: JSON.stringify(true),
        description: 'Enable automatic presentation generation',
        category: 'presentation',
        isPublic: true
      }
    ];

    for (const setting of defaultSettings) {
      await this.upsertSetting(setting);
    }

    this.logger.log('Default system settings initialized');
  }

  /**
   * Map Prisma model to DTO
   */
  private mapToDto(setting: any): SystemSettingDto {
    return {
      id: setting.id,
      key: setting.key,
      value: setting.value,
      description: setting.description,
      category: setting.category,
      isPublic: setting.isPublic,
      createdAt: setting.createdAt,
      updatedAt: setting.updatedAt
    };
  }
}
