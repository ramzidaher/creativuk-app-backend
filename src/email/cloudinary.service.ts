import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
// @ts-ignore
import { v2 as cloudinary } from 'cloudinary';

@Injectable()
export class CloudinaryService {
  private readonly logger = new Logger(CloudinaryService.name);

  constructor(private configService: ConfigService) {
    // Configure Cloudinary
    cloudinary.config({
      cloud_name: this.configService.get('CLOUDINARY_CLOUD_NAME'),
      api_key: this.configService.get('CLOUDINARY_API_KEY'),
      api_secret: this.configService.get('CLOUDINARY_API_SECRET'),
    });
  }

  async uploadImage(base64Data: string, opportunityId: string, fileName: string): Promise<string> {
    try {
      this.logger.log(`Uploading image ${fileName} for opportunity ${opportunityId}`);
      
      // Create folder path for the opportunity
      const folderPath = `creativ-solar/surveys/${opportunityId}`;
      
      // Remove data URL prefix if present
      let imageData = base64Data;
      if (base64Data.startsWith('data:')) {
        imageData = base64Data.split(',')[1];
      }

      // Upload to Cloudinary
      const result = await cloudinary.uploader.upload(
        `data:image/jpeg;base64,${imageData}`,
        {
          folder: folderPath,
          public_id: fileName.replace(/\.[^/.]+$/, ''), // Remove file extension
          resource_type: 'image',
          overwrite: true,
        }
      );

      this.logger.log(`Image uploaded successfully: ${result.secure_url}`);
      return result.secure_url;
    } catch (error) {
      this.logger.error(`Failed to upload image ${fileName}: ${error.message}`);
      throw error;
    }
  }

  async uploadMultipleImages(files: any[], opportunityId: string): Promise<{ [key: string]: string }> {
    try {
      this.logger.log(`Uploading ${files.length} images for opportunity ${opportunityId}`);
      
      const uploadPromises = files.map(async (file, index) => {
        const fileName = file.name || `file-${index + 1}`;
        const url = await this.uploadImage(file.base64, opportunityId, fileName);
        return { [fileName]: url };
      });

      const results = await Promise.all(uploadPromises);
      
      // Combine all results into a single object
      const uploadedUrls = results.reduce((acc, result) => {
        return { ...acc, ...result };
      }, {});

      this.logger.log(`Successfully uploaded ${files.length} images`);
      return uploadedUrls;
    } catch (error) {
      this.logger.error(`Failed to upload multiple images: ${error.message}`);
      throw error;
    }
  }

  async uploadSurveyFiles(uploadedFiles: any, opportunityId: string): Promise<{ [key: string]: string[] }> {
    try {
      this.logger.log(`Processing survey files for opportunity ${opportunityId}`);
      
      const uploadedUrls: { [key: string]: string[] } = {};

      for (const [fieldName, files] of Object.entries(uploadedFiles)) {
        if (Array.isArray(files) && files.length > 0) {
          this.logger.log(`Uploading ${files.length} files for field ${fieldName}`);
          
          const fieldUrls: string[] = [];
          
          for (const file of files) {
            if (file.base64) {
              const fileName = file.name || `${fieldName}-${Date.now()}`;
              const url = await this.uploadImage(file.base64, opportunityId, fileName);
              fieldUrls.push(url);
            }
          }
          
          uploadedUrls[fieldName] = fieldUrls;
        }
      }

      this.logger.log(`Successfully processed survey files: ${Object.keys(uploadedUrls).length} fields`);
      return uploadedUrls;
    } catch (error) {
      this.logger.error(`Failed to upload survey files: ${error.message}`);
      throw error;
    }
  }
}
