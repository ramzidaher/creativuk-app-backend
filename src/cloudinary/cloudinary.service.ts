import { Injectable } from '@nestjs/common';
import { v2 as cloudinary } from 'cloudinary';
import { ConfigService } from '@nestjs/config';

@Injectable()
export class CloudinaryService {
  constructor(private configService: ConfigService) {
    // Configure Cloudinary
    cloudinary.config({
      cloud_name: this.configService.get<string>('CLOUDINARY_CLOUD_NAME'),
      api_key: this.configService.get<string>('CLOUDINARY_API_KEY'),
      api_secret: this.configService.get<string>('CLOUDINARY_API_SECRET'),
    });
  }

  /**
   * Upload a file buffer to Cloudinary
   */
  async uploadFile(
    buffer: Buffer,
    folder: string,
    publicId?: string,
    options?: any
  ): Promise<{ url: string; public_id: string; secure_url: string }> {
    return new Promise((resolve, reject) => {
      // Build upload options - use resource_type from options if provided, otherwise default to 'auto'
      const uploadOptions = {
        folder: folder,
        public_id: publicId,
        resource_type: options?.resource_type || 'auto',
        ...options, // This will override resource_type if explicitly set in options
      };
      
      const uploadStream = cloudinary.uploader.upload_stream(
        uploadOptions,
        (error, result) => {
          if (error) {
            console.error('Cloudinary upload error:', error);
            reject(error);
          } else {
            console.log('✅ Cloudinary upload successful:', result?.secure_url);
            console.log('   Resource type:', result?.resource_type, 'Format:', result?.format);
            resolve({
              url: result?.url || '',
              public_id: result?.public_id || '',
              secure_url: result?.secure_url || '',
            });
          }
        }
      );

      uploadStream.end(buffer);
    });
  }

  /**
   * Upload a base64 image or file to Cloudinary
   */
  async uploadBase64Image(
    base64Data: string,
    folder: string,
    publicId?: string,
    options?: any
  ): Promise<{ url: string; public_id: string; secure_url: string }> {
    return new Promise((resolve, reject) => {
      // Determine resource type from options or default to 'image'
      const resourceType = options?.resource_type || 'image';
      
      // Remove data URL prefix if present
      let base64String = base64Data.replace(/^data:[^;]+;base64,/, '');
      
      // Determine the data URL format based on resource type
      let dataUrl: string;
      if (resourceType === 'raw' || resourceType === 'auto') {
        // For PDFs and raw files, try to detect from the base64 data or use application/octet-stream
        // Check if the original base64Data had a mime type
        const mimeMatch = base64Data.match(/^data:([^;]+);base64,/);
        const mimeType = mimeMatch ? mimeMatch[1] : 'application/pdf';
        dataUrl = `data:${mimeType};base64,${base64String}`;
      } else {
        // For images, use image/jpeg as default
        dataUrl = `data:image/jpeg;base64,${base64String}`;
      }
      
      cloudinary.uploader.upload(
        dataUrl,
        {
          folder: folder,
          public_id: publicId,
          resource_type: resourceType,
          ...options,
        },
        (error, result) => {
          if (error) {
            console.error('Cloudinary base64 upload error:', error);
            reject(error);
          } else {
            // console.log('Cloudinary base64 upload successful:', result?.secure_url);
            resolve({
              url: result?.url || '',
              public_id: result?.public_id || '',
              secure_url: result?.secure_url || '',
            });
          }
        }
      );
    });
  }

  /**
   * Generate a signed URL for a file (useful for PDFs and raw files)
   */
  generateSignedUrl(publicId: string, resourceType: 'image' | 'raw' | 'video' | 'auto' = 'image', options?: any): string {
    try {
      const url = cloudinary.url(publicId, {
        resource_type: resourceType,
        secure: true,
        sign_url: true, // Generate signed URL
        ...options,
      });
      return url;
    } catch (error) {
      console.error('Cloudinary signed URL generation error:', error);
      throw error;
    }
  }

  /**
   * Generate a public URL for a file
   */
  generatePublicUrl(publicId: string, resourceType: 'image' | 'raw' | 'video' | 'auto' = 'image', options?: any): string {
    try {
      const url = cloudinary.url(publicId, {
        resource_type: resourceType,
        secure: true,
        sign_url: false, // Public URL
        ...options,
      });
      return url;
    } catch (error) {
      console.error('Cloudinary public URL generation error:', error);
      throw error;
    }
  }

  /**
   * Delete a file from Cloudinary
   */
  async deleteFile(publicId: string): Promise<boolean> {
    return new Promise((resolve, reject) => {
      cloudinary.uploader.destroy(publicId, (error, result) => {
        if (error) {
          console.error('Cloudinary delete error:', error);
          reject(error);
        } else {
          console.log('✅ Cloudinary delete successful:', publicId);
          resolve(result?.result === 'ok');
        }
      });
    });
  }
}
