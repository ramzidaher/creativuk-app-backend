import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { UserService } from '../user/user.service';
import { CloudinaryService } from '../cloudinary/cloudinary.service';
import { AutoSaveDto, AutoSaveImageDto, GetAutoSaveDataDto } from './dto/auto-save.dto';

@Injectable()
export class AutoSaveService {
  private readonly logger = new Logger(AutoSaveService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly userService: UserService,
    private readonly cloudinaryService: CloudinaryService,
  ) {}

  /**
   * Auto-save a single field value
   */
  async autoSaveField(userId: string, autoSaveDto: AutoSaveDto): Promise<{ success: boolean; message: string }> {
    try {
      // this.logger.log(`Auto-saving field ${autoSaveDto.fieldName} for opportunity ${autoSaveDto.opportunityId}`);

      // Get user
      const user = await this.getUser(userId);
      
      // Get or create auto-save record
      const autoSaveRecord = await this.getOrCreateAutoSaveRecord(user.id, autoSaveDto.opportunityId);

      // Update the specific field in the JSON data
      const currentData = autoSaveRecord.data as any || {};
      
      // If pageData is provided, merge it into the current data
      if (autoSaveDto.pageData) {
        Object.assign(currentData, autoSaveDto.pageData);
        // Update lastPage when saving page data
        if (autoSaveDto.pageName && !autoSaveDto.skipLastPageUpdate) {
          currentData.lastPage = autoSaveDto.pageName;
        }
      } else if (autoSaveDto.fieldName && autoSaveDto.fieldValue !== undefined) {
        // Handle special case for lastPage field
        if (autoSaveDto.fieldName === 'lastPage') {
          currentData.lastPage = autoSaveDto.fieldValue;
          // this.logger.log(`Updated lastPage to: ${autoSaveDto.fieldValue}`);
        } else {
          // Update specific field
          const fieldParts = autoSaveDto.fieldName.split('.');
          let target = currentData;
          
          // Navigate to the nested field
          for (let i = 0; i < fieldParts.length - 1; i++) {
            if (!target[fieldParts[i]]) {
              target[fieldParts[i]] = {};
            }
            target = target[fieldParts[i]];
          }
          
          // Set the final field value
          target[fieldParts[fieldParts.length - 1]] = autoSaveDto.fieldValue;
          
          // Update lastPage when saving to a specific page (first part of field name)
          // Skip if this is an auto-fill operation
          if (fieldParts.length > 1 && !autoSaveDto.skipLastPageUpdate) {
            currentData.lastPage = fieldParts[0];
            this.logger.log(`📄 Updated lastPage to: ${fieldParts[0]} for field: ${autoSaveDto.fieldName}`);
          } else if (autoSaveDto.skipLastPageUpdate) {
            this.logger.log(`📄 Skipping lastPage update for auto-fill field: ${autoSaveDto.fieldName}`);
          }
        }
      }

      // Update the auto-save record
      this.logger.log(`📄 Updating auto-save record with data: ${JSON.stringify(this.truncateBase64ForLogging(currentData), null, 2)}`);
      const updatedRecord = await this.prisma.autoSave.update({
        where: { id: autoSaveRecord.id },
        data: {
          data: currentData,
          lastSavedAt: new Date(),
          updatedAt: new Date()
        }
      });
      this.logger.log(`📄 Auto-save record updated successfully. New data: ${JSON.stringify(this.truncateBase64ForLogging(updatedRecord.data), null, 2)}`);

      this.logger.log(`Successfully auto-saved field ${autoSaveDto.fieldName}`);
      return { success: true, message: 'Field auto-saved successfully' };
    } catch (error) {
      this.logger.error(`Failed to auto-save field: ${error.message}`);
      return { success: false, message: `Failed to auto-save: ${error.message}` };
    }
  }

  /**
   * Auto-save an image
   */
  async autoSaveImage(userId: string, autoSaveImageDto: AutoSaveImageDto): Promise<{ success: boolean; message: string; imageUrl?: string }> {
    try {
      this.logger.log(`Auto-saving image for field ${autoSaveImageDto.fieldName} in opportunity ${autoSaveImageDto.opportunityId}`);

      // Get user
      const user = await this.getUser(userId);
      
      // Get or create auto-save record
      const autoSaveRecord = await this.getOrCreateAutoSaveRecord(user.id, autoSaveImageDto.opportunityId);

      // Upload image to Cloudinary
      const publicId = `auto-save/${autoSaveImageDto.opportunityId}/${autoSaveImageDto.fieldName}-${Date.now()}`;
      
      const uploadResult = await this.cloudinaryService.uploadBase64Image(
        autoSaveImageDto.base64Data,
        `auto-save/${autoSaveImageDto.opportunityId}`,
        publicId,
        {
          resource_type: 'image',
          quality: 'auto',
          fetch_format: 'auto'
        }
      );

      // Update the auto-save record with image URL
      const currentData = autoSaveRecord.data as any || {};
      
      // Initialize images object if it doesn't exist
      if (!currentData.images) {
        currentData.images = {};
      }
      
      // Initialize field images array if it doesn't exist
      if (!currentData.images[autoSaveImageDto.fieldName]) {
        currentData.images[autoSaveImageDto.fieldName] = [];
      }
      
      // Add the new image URL
      currentData.images[autoSaveImageDto.fieldName].push({
        url: uploadResult.secure_url,
        publicId: publicId,
        fileName: autoSaveImageDto.fileName || `${autoSaveImageDto.fieldName}-${Date.now()}`,
        mimeType: autoSaveImageDto.mimeType || 'image/jpeg',
        fileSize: autoSaveImageDto.fileSize || 0,
        uploadedAt: new Date().toISOString()
      });

      // Update lastPage when saving an image (extract page from field name)
      if (autoSaveImageDto.fieldName.includes('.')) {
        const pageName = autoSaveImageDto.fieldName.split('.')[0];
        currentData.lastPage = pageName;
        this.logger.log(`📄 Updated lastPage to: ${pageName} for image field: ${autoSaveImageDto.fieldName}`);
      }

      // Update the auto-save record
      await this.prisma.autoSave.update({
        where: { id: autoSaveRecord.id },
        data: {
          data: currentData,
          lastSavedAt: new Date(),
          updatedAt: new Date()
        }
      });

      this.logger.log(`Successfully auto-saved image for field ${autoSaveImageDto.fieldName}: ${uploadResult.secure_url}`);
      return { 
        success: true, 
        message: 'Image auto-saved successfully',
        imageUrl: uploadResult.secure_url
      };
    } catch (error) {
      this.logger.error(`Failed to auto-save image: ${error.message}`);
      return { success: false, message: `Failed to auto-save image: ${error.message}` };
    }
  }

  /**
   * Get auto-saved data for an opportunity
   */
  async getAutoSaveData(userId: string, getAutoSaveDataDto: GetAutoSaveDataDto): Promise<{ success: boolean; data?: any; message?: string }> {
    try {
      this.logger.log(`Getting auto-saved data for opportunity ${getAutoSaveDataDto.opportunityId}`);

      // Get user
      const user = await this.getUser(userId);
      
      // Get auto-save record
      const autoSaveRecord = await this.prisma.autoSave.findFirst({
        where: {
          userId: user.id,
          opportunityId: getAutoSaveDataDto.opportunityId,
          isDeleted: false
        }
      });

      if (!autoSaveRecord) {
        this.logger.log(`No auto-saved data found for opportunity ${getAutoSaveDataDto.opportunityId}`);
        return { 
          success: true, 
          data: null,
          message: 'No auto-saved data found'
        };
      }

      let data = autoSaveRecord.data as any || {};

      // If pageName is specified, return only that page's data
      if (getAutoSaveDataDto.pageName && data[getAutoSaveDataDto.pageName]) {
        data = { [getAutoSaveDataDto.pageName]: data[getAutoSaveDataDto.pageName] };
      }

      // this.logger.log(`Retrieved auto-saved data for opportunity ${getAutoSaveDataDto.opportunityId}`);
      // this.logger.log(`Data structure: ${JSON.stringify(this.truncateBase64ForLogging(data), null, 2)}`);
      // this.logger.log(`LastPage value: ${data.lastPage || 'undefined'}`);  
      return { 
        success: true, 
        data: data,
        message: 'Auto-saved data retrieved successfully'
      };
    } catch (error) {
      this.logger.error(`Failed to get auto-saved data: ${error.message}`);
      return { success: false, message: `Failed to get auto-saved data: ${error.message}` };
    }
  }

  /**
   * Clear auto-saved data for an opportunity
   */
  async clearAutoSaveData(userId: string, opportunityId: string): Promise<{ success: boolean; message: string }> {
    try {
      this.logger.log(`Clearing auto-saved data for opportunity ${opportunityId}`);

      // Get user
      const user = await this.getUser(userId);
      
      // Soft delete auto-save record
      await this.prisma.autoSave.updateMany({
        where: {
          userId: user.id,
          opportunityId: opportunityId,
          isDeleted: false
        },
        data: {
          isDeleted: true,
          deletedAt: new Date(),
          updatedAt: new Date()
        }
      });

      this.logger.log(`Successfully cleared auto-saved data for opportunity ${opportunityId}`);
      return { success: true, message: 'Auto-saved data cleared successfully' };
    } catch (error) {
      this.logger.error(`Failed to clear auto-saved data: ${error.message}`);
      return { success: false, message: `Failed to clear auto-saved data: ${error.message}` };
    }
  }

  /**
   * Transfer auto-saved data to survey
   */
  async transferToSurvey(userId: string, opportunityId: string): Promise<{ success: boolean; message: string; surveyData?: any }> {
    try {
      this.logger.log(`Transferring auto-saved data to survey for opportunity ${opportunityId}`);

      // Get auto-saved data
      const autoSaveResult = await this.getAutoSaveData(userId, { opportunityId });
      
      if (!autoSaveResult.success || !autoSaveResult.data) {
        return { success: false, message: 'No auto-saved data found to transfer' };
      }

      // Get user
      const user = await this.getUser(userId);
      
      // Get or create survey
      let survey = await this.prisma.survey.findFirst({
        where: {
          ghlOpportunityId: opportunityId,
          ghlUserId: user.ghlUserId,
          isDeleted: false
        }
      });

      if (!survey) {
        // Create new survey - use upsert to handle potential race conditions
        survey = await this.prisma.survey.upsert({
          where: {
            ghlOpportunityId: opportunityId
          },
          update: {
            ghlUserId: user.ghlUserId,
            status: 'DRAFT',
            updatedBy: user.id,
            updatedAt: new Date()
          },
          create: {
            ghlOpportunityId: opportunityId,
            ghlUserId: user.ghlUserId,
            status: 'DRAFT',
            createdBy: user.id
          }
        });
      }

      // Update survey with auto-saved data
      const surveyData = autoSaveResult.data;
      const updateData: any = {
        updatedAt: new Date()
      };

      // Map auto-saved data to survey pages
      if (surveyData.page1) updateData.page1 = surveyData.page1;
      if (surveyData.page2) updateData.page2 = surveyData.page2;
      if (surveyData.page3) updateData.page3 = surveyData.page3;
      if (surveyData.page4) updateData.page4 = surveyData.page4;
      if (surveyData.page5) updateData.page5 = surveyData.page5;
      if (surveyData.page6) updateData.page6 = surveyData.page6;
      if (surveyData.page7) updateData.page7 = surveyData.page7;
      if (surveyData.page8) updateData.page8 = surveyData.page8;

      await this.prisma.survey.update({
        where: { id: survey.id },
        data: updateData
      });

      // Process images if they exist
      if (surveyData.images) {
        for (const [fieldName, images] of Object.entries(surveyData.images)) {
          if (Array.isArray(images)) {
            for (const image of images as any[]) {
              await this.prisma.surveyImage.create({
                data: {
                  surveyId: survey.id,
                  fieldName: fieldName,
                  fileName: image.fileName,
                  originalName: image.fileName,
                  mimeType: image.mimeType,
                  fileSize: image.fileSize,
                  filePath: image.url, // Cloudinary URL
                  base64Data: null
                }
              });
            }
          }
        }
      }

      // Clear auto-saved data after successful transfer
      await this.clearAutoSaveData(userId, opportunityId);

      this.logger.log(`Successfully transferred auto-saved data to survey ${survey.id}`);
      return { 
        success: true, 
        message: 'Auto-saved data transferred to survey successfully',
        surveyData: surveyData
      };
    } catch (error) {
      this.logger.error(`Failed to transfer auto-saved data: ${error.message}`);
      return { success: false, message: `Failed to transfer auto-saved data: ${error.message}` };
    }
  }

  /**
   * Helper method to get user by internal user ID
   */
  private async getUser(userId: string) {
    this.logger.debug(`Looking up user with internal ID: ${userId}`);
    
    const user = await this.userService.findById(userId);
    
    if (!user) {
      this.logger.error(`User not found with internal ID: ${userId}`);
      throw new NotFoundException(`User not found with internal ID: ${userId}`);
    }

    this.logger.debug(`Found user: ${user.username} (${user.id}) with ghlUserId: ${user.ghlUserId}`);
    return user;
  }

  /**
   * Helper method to truncate base64 data for logging
   */
  private truncateBase64ForLogging(data: any): any {
    if (typeof data === 'string' && data.length > 100 && data.match(/^[A-Za-z0-9+/=]+$/)) {
      // Likely base64 data, truncate it
      return `${data.substring(0, 50)}... (truncated, ${data.length} chars)`;
    }
    
    if (Array.isArray(data)) {
      return data.map(item => this.truncateBase64ForLogging(item));
    }
    
    if (data && typeof data === 'object') {
      const truncated: any = {};
      for (const [key, value] of Object.entries(data)) {
        if (key === 'base64' || key === 'base64Data') {
          truncated[key] = typeof value === 'string' && value.length > 100 
            ? `${value.substring(0, 50)}... (truncated, ${value.length} chars)`
            : value;
        } else {
          truncated[key] = this.truncateBase64ForLogging(value);
        }
      }
      return truncated;
    }
    
    return data;
  }

  /**
   * Helper method to get or create auto-save record
   */
  private async getOrCreateAutoSaveRecord(userId: string, opportunityId: string) {
    let autoSaveRecord = await this.prisma.autoSave.findFirst({
      where: {
        userId: userId,
        opportunityId: opportunityId,
        isDeleted: false
      }
    });

    if (!autoSaveRecord) {
      autoSaveRecord = await this.prisma.autoSave.create({
        data: {
          userId: userId,
          opportunityId: opportunityId,
          data: {},
          lastSavedAt: new Date()
        }
      });
    }

    return autoSaveRecord;
  }
}
