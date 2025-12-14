import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import * as fs from 'fs';
import * as path from 'path';
import { promisify } from 'util';

const writeFile = promisify(fs.writeFile);
const mkdir = promisify(fs.mkdir);

@Injectable()
export class SurveyImageService {
  private readonly logger = new Logger(SurveyImageService.name);
  private readonly uploadsDir = path.join(process.cwd(), 'uploads', 'survey-images');

  constructor(private readonly prisma: PrismaService) {
    this.ensureUploadsDirectory();
  }

  private async ensureUploadsDirectory() {
    try {
      if (!fs.existsSync(this.uploadsDir)) {
        await mkdir(this.uploadsDir, { recursive: true });
        this.logger.log(`Created uploads directory: ${this.uploadsDir}`);
      }
    } catch (error) {
      this.logger.error(`Failed to create uploads directory: ${error.message}`);
    }
  }

  async saveSurveyImages(surveyId: string, uploadedFiles: any, opportunityId: string): Promise<{ [key: string]: string[] }> {
    try {
      this.logger.log(`Saving survey images for survey: ${surveyId}, opportunity: ${opportunityId}`);
      
      const savedImages: { [key: string]: string[] } = {};
      
      for (const [fieldName, files] of Object.entries(uploadedFiles)) {
        if (!Array.isArray(files) || files.length === 0) continue;
        
        const fieldImages: string[] = [];
        
        for (const file of files) {
          try {
            const imageRecord = await this.saveSingleImage(surveyId, fieldName, file, opportunityId);
            fieldImages.push(imageRecord.id);
          } catch (error) {
            this.logger.error(`Failed to save image for field ${fieldName}: ${error.message}`);
          }
        }
        
        if (fieldImages.length > 0) {
          savedImages[fieldName] = fieldImages;
        }
      }
      
      this.logger.log(`Successfully saved ${Object.keys(savedImages).length} image fields for survey ${surveyId}`);
      return savedImages;
    } catch (error) {
      this.logger.error(`Failed to save survey images: ${error.message}`);
      throw error;
    }
  }

  private async saveSingleImage(surveyId: string, fieldName: string, file: any, opportunityId: string): Promise<any> {
    try {
      // Generate unique filename with better naming
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const fileExtension = this.getFileExtension(file.mimeType || 'image/jpeg');
      const fileName = `${fieldName}_${timestamp}${fileExtension}`;
      
      // Create organized directory structure: surveys/opportunityId/fieldName/
      const surveyDir = path.join(this.uploadsDir, opportunityId);
      const fieldDir = path.join(surveyDir, fieldName);
      if (!fs.existsSync(fieldDir)) {
        await mkdir(fieldDir, { recursive: true });
      }
      
      const filePath = path.join(fieldDir, fileName);
      
      // Save file to disk
      if (file.base64) {
        // Handle base64 data
        const base64Data = file.base64.replace(/^data:image\/[a-z]+;base64,/, '');
        const buffer = Buffer.from(base64Data, 'base64');
        await writeFile(filePath, buffer);
      } else if (file.uri) {
        // Handle file URI (for mobile)
        const buffer = fs.readFileSync(file.uri);
        await writeFile(filePath, buffer);
      } else {
        throw new Error('No valid file data provided');
      }
      
      // Save image record to database
      const imageRecord = await this.prisma.surveyImage.create({
        data: {
          surveyId,
          fieldName,
          fileName,
          originalName: file.name || fileName,
          mimeType: file.mimeType || 'image/jpeg',
          fileSize: file.size || 0,
          filePath: filePath,
          base64Data: file.base64 || null, // Store base64 for quick access
        }
      });
      
      this.logger.log(`Saved image: ${fileName} for field: ${fieldName} in opportunity: ${opportunityId}`);
      return imageRecord;
    } catch (error) {
      this.logger.error(`Failed to save single image: ${error.message}`);
      throw error;
    }
  }

  private getFileExtension(mimeType: string): string {
    const extensions: { [key: string]: string } = {
      'image/jpeg': '.jpg',
      'image/jpg': '.jpg',
      'image/png': '.png',
      'image/gif': '.gif',
      'image/webp': '.webp',
      'application/pdf': '.pdf',
    };
    
    return extensions[mimeType] || '.jpg';
  }

  async getSurveyImages(surveyId: string): Promise<any[]> {
    try {
      const images = await this.prisma.surveyImage.findMany({
        where: { surveyId },
        orderBy: { createdAt: 'asc' }
      });
      
      return images;
    } catch (error) {
      this.logger.error(`Failed to get survey images: ${error.message}`);
      throw error;
    }
  }

  async getSurveyImagesByField(surveyId: string, fieldName: string): Promise<any[]> {
    try {
      const images = await this.prisma.surveyImage.findMany({
        where: { 
          surveyId,
          fieldName 
        },
        orderBy: { createdAt: 'asc' }
      });
      
      return images;
    } catch (error) {
      this.logger.error(`Failed to get survey images by field: ${error.message}`);
      throw error;
    }
  }

  async deleteSurveyImage(imageId: string): Promise<void> {
    try {
      const image = await this.prisma.surveyImage.findUnique({
        where: { id: imageId }
      });
      
      if (!image) {
        throw new Error('Image not found');
      }
      
      // Delete file from disk
      if (fs.existsSync(image.filePath)) {
        fs.unlinkSync(image.filePath);
      }
      
      // Delete record from database
      await this.prisma.surveyImage.delete({
        where: { id: imageId }
      });
      
      this.logger.log(`Deleted image: ${image.fileName}`);
    } catch (error) {
      this.logger.error(`Failed to delete survey image: ${error.message}`);
      throw error;
    }
  }

  async deleteAllSurveyImages(surveyId: string): Promise<void> {
    try {
      const images = await this.prisma.surveyImage.findMany({
        where: { surveyId }
      });
      
      // Delete all files from disk
      for (const image of images) {
        if (fs.existsSync(image.filePath)) {
          fs.unlinkSync(image.filePath);
        }
      }
      
      // Delete all records from database
      await this.prisma.surveyImage.deleteMany({
        where: { surveyId }
      });
      
      this.logger.log(`Deleted ${images.length} images for survey: ${surveyId}`);
    } catch (error) {
      this.logger.error(`Failed to delete all survey images: ${error.message}`);
      throw error;
    }
  }

  async getImageAsBase64(imageId: string): Promise<string | null> {
    try {
      const image = await this.prisma.surveyImage.findUnique({
        where: { id: imageId }
      });
      
      if (!image) {
        return null;
      }
      
      // Return stored base64 data if available
      if (image.base64Data) {
        return image.base64Data;
      }
      
      // Otherwise, read from file and convert to base64
      if (fs.existsSync(image.filePath)) {
        const buffer = fs.readFileSync(image.filePath);
        return `data:${image.mimeType};base64,${buffer.toString('base64')}`;
      }
      
      return null;
    } catch (error) {
      this.logger.error(`Failed to get image as base64: ${error.message}`);
      return null;
    }
  }


  async getSurveyImagesByOpportunity(opportunityId: string): Promise<any[]> {
    try {
      // Get all images for this opportunity by looking in the directory structure
      const surveyDir = path.join(this.uploadsDir, opportunityId);
      if (!fs.existsSync(surveyDir)) {
        return [];
      }

      const images = await this.prisma.surveyImage.findMany({
        where: {
          filePath: {
            contains: opportunityId
          }
        },
        orderBy: { createdAt: 'asc' }
      });
      
      return images;
    } catch (error) {
      this.logger.error(`Failed to get survey images by opportunity: ${error.message}`);
      return [];
    }
  }
}
