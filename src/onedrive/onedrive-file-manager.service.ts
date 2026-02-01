import { Injectable, Logger } from '@nestjs/common';
import * as fs from 'fs';
import * as fsPromises from 'fs/promises';
import * as path from 'path';
import * as crypto from 'crypto';
import { exec } from 'child_process';
import { promisify } from 'util';
import axios from 'axios';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class OneDriveFileManagerService {
  private readonly logger = new Logger(OneDriveFileManagerService.name);
  private readonly execAsync = promisify(exec);

  constructor(private readonly prisma: PrismaService) {}

  // OneDrive folder paths
  private readonly quotationsFolder = 'C:\\Users\\\Creativuk\\OneDrive - JARMQ LTD\\Pamela Rennie\'s files - Customer Quotations';
  private readonly ordersFolder = 'C:\\Users\\\Creativuk\\OneDrive - JARMQ LTD\\Pamela Rennie\'s files - Customer Orders 2';

  /**
   * Get survey images for a specific opportunity
   */
  async getSurveyImagesForOpportunity(opportunityId: string): Promise<Array<{
    fieldName: string;
    fileName: string;
    filePath: string;
    originalName?: string;
  }>> {
    try {
      // Find survey by opportunity ID
      const survey = await this.prisma.survey.findFirst({
        where: {
          ghlOpportunityId: opportunityId,
          isDeleted: false
        }
      });

      if (!survey) {
        this.logger.log(`No survey found for opportunity: ${opportunityId}`);
        return [];
      }

      // First, try to get images from SurveyImage table
      const surveyImages = await this.prisma.surveyImage.findMany({
        where: {
          surveyId: survey.id
        },
        orderBy: { createdAt: 'asc' }
      });

      this.logger.log(`Found ${surveyImages.length} images in SurveyImage table for opportunity: ${opportunityId}`);

      if (surveyImages.length > 0) {
        const mappedImages = surveyImages.map(image => ({
          fieldName: image.fieldName,
          fileName: image.fileName,
          filePath: image.filePath,
          originalName: image.originalName
        }));
        this.logger.log(`Returning ${mappedImages.length} images from SurveyImage table`);
        return mappedImages;
      }

      // If no images in SurveyImage table, extract from survey JSON data
      this.logger.log(`No images found in SurveyImage table, extracting from survey JSON data for opportunity: ${opportunityId}`);
      const extractedImages = this.extractImagesFromSurveyData(survey);
      this.logger.log(`Extracted ${extractedImages.length} images from survey JSON data`);
      return extractedImages;

    } catch (error) {
      this.logger.error(`Failed to get survey images for opportunity ${opportunityId}: ${error.message}`);
      return [];
    }
  }

  /**
   * Extract image URLs from survey JSON data
   */
  private extractImagesFromSurveyData(survey: any): Array<{
    fieldName: string;
    fileName: string;
    filePath: string;
    originalName?: string;
  }> {
    const images: Array<{
      fieldName: string;
      fileName: string;
      filePath: string;
      originalName?: string;
    }> = [];

    try {
      // Define the pages to check for image fields
      const pages = ['page1', 'page2', 'page3', 'page4', 'page5', 'page6', 'page7', 'page8'];
      
      for (const page of pages) {
        const pageData = survey[page];
        if (!pageData || typeof pageData !== 'object') continue;

        // Check each field in the page for image arrays
        for (const [fieldName, fieldValue] of Object.entries(pageData)) {
          if (Array.isArray(fieldValue)) {
            // Check if this looks like an image URL array
            const imageUrls = fieldValue.filter(item => 
              typeof item === 'string' && 
              (item.includes('cloudinary.com') || item.includes('.jpg') || item.includes('.jpeg') || item.includes('.png'))
            );

            if (imageUrls.length > 0) {
              // Add each image URL
              imageUrls.forEach((url, index) => {
                const fileName = `${fieldName}_${index + 1}.jpg`;
                images.push({
                  fieldName: fieldName,
                  fileName: fileName,
                  filePath: url,
                  originalName: fileName
                });
              });
            }
          }
        }
      }

      this.logger.log(`Extracted ${images.length} images from survey JSON data`);
      return images;

    } catch (error) {
      this.logger.error(`Error extracting images from survey data: ${error.message}`);
      return [];
    }
  }

  /**
   * Download image from URL (Cloudinary or local file)
   */
  private async downloadImageFromUrl(imageUrl: string, destinationPath: string): Promise<{ success: boolean; filename?: string }> {
    try {
      // Check if it's a local file path
      if (fs.existsSync(imageUrl)) {
        await this.copyFile(imageUrl, destinationPath);
        return { success: true, filename: path.basename(destinationPath) };
      }

      this.logger.log(`📥 [DOWNLOAD] Starting download from URL: ${imageUrl.substring(0, 100)}...`);

      // Download from URL (Cloudinary)
      const response = await axios({
        method: 'GET',
        url: imageUrl,
        responseType: 'stream',
        timeout: 30000, // 30 second timeout
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
      });

      // Check if the response is successful
      if (response.status !== 200) {
        this.logger.error(`❌ [DOWNLOAD] HTTP ${response.status} error downloading from: ${imageUrl}`);
        return { success: false };
      }

      // Get content type to determine proper file extension
      const contentType = response.headers['content-type'] || '';
      this.logger.log(`📥 [DOWNLOAD] Content-Type: ${contentType}`);

      // Determine file extension based on content type
      let fileExtension = '.jpg'; // default
      if (contentType.includes('image/png')) {
        fileExtension = '.png';
      } else if (contentType.includes('image/jpeg') || contentType.includes('image/jpg')) {
        fileExtension = '.jpg';
      } else if (contentType.includes('image/webp')) {
        fileExtension = '.webp';
      } else if (contentType.includes('image/gif')) {
        fileExtension = '.gif';
      }

      // Update destination path with correct extension
      const pathWithoutExt = destinationPath.replace(/\.[^/.]+$/, '');
      const finalDestinationPath = pathWithoutExt + fileExtension;

      this.logger.log(`📥 [DOWNLOAD] Downloading to: ${finalDestinationPath}`);

      const writer = fs.createWriteStream(finalDestinationPath);
      response.data.pipe(writer);

      return new Promise((resolve, reject) => {
        writer.on('finish', () => {
          // Verify the file was created and has content
          if (fs.existsSync(finalDestinationPath)) {
            const stats = fs.statSync(finalDestinationPath);
            if (stats.size > 0) {
              this.logger.log(`✅ [DOWNLOAD] Successfully downloaded image: ${finalDestinationPath} (${stats.size} bytes)`);
              resolve({ success: true, filename: path.basename(finalDestinationPath) });
            } else {
              this.logger.error(`❌ [DOWNLOAD] Downloaded file is empty: ${finalDestinationPath}`);
              resolve({ success: false });
            }
          } else {
            this.logger.error(`❌ [DOWNLOAD] Downloaded file does not exist: ${finalDestinationPath}`);
            resolve({ success: false });
          }
        });
        writer.on('error', (error) => {
          this.logger.error(`❌ [DOWNLOAD] Error writing image: ${error.message}`);
          reject(error);
        });
      });
    } catch (error) {
      this.logger.error(`❌ [DOWNLOAD] Failed to download image from ${imageUrl}: ${error.message}`);
      if (error.response) {
        this.logger.error(`❌ [DOWNLOAD] Response status: ${error.response.status}`);
        this.logger.error(`❌ [DOWNLOAD] Response headers:`, error.response.headers);
      }
      return { success: false };
    }
  }

  /**
   * Copy survey images to OneDrive customer folder
   */
  async copySurveyImagesToCustomerFolder(
    customerFolderPath: string,
    surveyImages: Array<{
      fieldName: string;
      fileName: string;
      filePath: string;
      originalName?: string;
    }>
  ): Promise<{ success: boolean; message: string; copiedImages?: string[]; error?: string }> {
    try {
      if (!surveyImages || surveyImages.length === 0) {
        return {
          success: true,
          message: 'No survey images to copy',
          copiedImages: []
        };
      }

      // Create survey_images folder
      const surveyImagesFolder = path.join(customerFolderPath, 'survey_images');
      if (!fs.existsSync(surveyImagesFolder)) {
        await fsPromises.mkdir(surveyImagesFolder, { recursive: true });
        this.logger.log(`Created survey_images folder: ${surveyImagesFolder}`);
      }

      this.logger.log(`📷 [SURVEY_IMAGES] Starting to copy ${surveyImages.length} survey images to OneDrive`);

      const copiedImages: string[] = [];
      let successCount = 0;
      let skippedCount = 0;
      
      // Track processed file paths to avoid duplicates
      const processedFilePaths = new Set<string>();

      for (let index = 0; index < surveyImages.length; index++) {
        const image = surveyImages[index];
        try {
          // Skip if we've already processed this exact file path
          if (processedFilePaths.has(image.filePath)) {
            this.logger.log(`⏭️ [COPY] Skipping duplicate file path: ${image.filePath.substring(0, 100)}...`);
            skippedCount++;
            continue;
          }
          processedFilePaths.add(image.filePath);

          // Generate a unique filename based on the original filename or field name
          // Always include a unique identifier (hash of filePath) to ensure uniqueness
          const pathHash = crypto.createHash('md5').update(image.filePath).digest('hex').substring(0, 8);
          let baseFileName: string;
          
          if (image.originalName) {
            // Use original name but sanitize it
            baseFileName = image.originalName.replace(/[<>:"/\\|?*]/g, '_');
            // Remove extension if present (we'll add it based on downloaded content)
            baseFileName = baseFileName.replace(/\.[^/.]+$/, '');
            // Always append hash to ensure uniqueness, even if multiple images have same original name
            baseFileName = `${baseFileName}_${pathHash}`;
          } else if (image.fileName) {
            baseFileName = image.fileName.replace(/[<>:"/\\|?*]/g, '_').replace(/\.[^/.]+$/, '');
            // Always append hash to ensure uniqueness
            baseFileName = `${baseFileName}_${pathHash}`;
          } else {
            // Fallback: use field name with a hash of the file path for uniqueness
            baseFileName = `${image.fieldName}_${pathHash}`;
          }

          // Check for existing files with similar names (same base name, any extension)
          const existingFiles = fs.existsSync(surveyImagesFolder) 
            ? fs.readdirSync(surveyImagesFolder)
            : [];
          
          // Check if a file with the same base name already exists (with any extension)
          const existingFile = existingFiles.find(file => {
            const fileBaseName = file.replace(/\.[^/.]+$/, '');
            return fileBaseName === baseFileName;
          });

          if (existingFile) {
            // File already exists, skip copying
            this.logger.log(`⏭️ [COPY] Skipping ${image.fieldName} - file already exists: ${existingFile}`);
            copiedImages.push(existingFile);
            skippedCount++;
            continue;
          }

          // File doesn't exist, proceed with download/copy
          const destinationPath = path.join(surveyImagesFolder, baseFileName);

          // Download/copy the image (this will add the correct extension)
          const downloadResult = await this.downloadImageFromUrl(image.filePath, destinationPath);
          
          if (downloadResult.success) {
            // Get the actual filename that was created (with correct extension)
            const actualFileName = downloadResult.filename || `${baseFileName}.jpg`;
            copiedImages.push(actualFileName);
            successCount++;
            this.logger.log(`✅ [COPY] Copied survey image: ${actualFileName}`);
          } else {
            this.logger.error(`❌ [COPY] Failed to copy survey image: ${image.fieldName}`);
          }
        } catch (error) {
          this.logger.error(`Failed to copy survey image ${image.fileName}: ${error.message}`);
        }
      }

      const totalProcessed = successCount + skippedCount;
      const failedCount = surveyImages.length - totalProcessed;
      const message = skippedCount > 0
        ? `Processed ${totalProcessed}/${surveyImages.length} survey images (${successCount} copied, ${skippedCount} already existed${failedCount > 0 ? `, ${failedCount} failed` : ''})`
        : `Successfully copied ${successCount}/${surveyImages.length} survey images${failedCount > 0 ? ` (${failedCount} failed)` : ''}`;

      this.logger.log(`📷 [SURVEY_IMAGES] Copy complete: ${message}`);

      return {
        success: totalProcessed > 0,
        message,
        copiedImages,
        error: totalProcessed === 0 ? 'Failed to process any survey images' : undefined
      };

    } catch (error) {
      this.logger.error(`Error copying survey images: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy survey images',
        error: error.message
      };
    }
  }

  /**
   * Copy proposal files to OneDrive quotations folder with automatic survey image retrieval
   */
  async copyProposalToQuotationsWithSurveyImages(
    opportunityId: string,
    customerName: string,
    proposalFiles: { pptxPath?: string; pdfPath?: string }
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      // Automatically fetch survey images for this opportunity
      const surveyImages = await this.getSurveyImagesForOpportunity(opportunityId);
      this.logger.log(`Found ${surveyImages.length} survey images for opportunity: ${opportunityId}`);
      
      return await this.copyProposalToQuotations(opportunityId, customerName, proposalFiles, surveyImages);
    } catch (error) {
      this.logger.error(`Error copying proposal with survey images: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy proposal files with survey images',
        error: error.message
      };
    }
  }

  /**
   * Copy contract files to OneDrive orders folder with automatic survey image retrieval
   */
  async copyContractToOrdersWithSurveyImages(
    opportunityId: string,
    customerName: string,
    contractFiles: { pptxPath?: string; pdfPath?: string }
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      // Automatically fetch survey images for this opportunity
      const surveyImages = await this.getSurveyImagesForOpportunity(opportunityId);
      this.logger.log(`Found ${surveyImages.length} survey images for opportunity: ${opportunityId}`);
      
      return await this.copyContractToOrders(opportunityId, customerName, contractFiles, surveyImages);
    } catch (error) {
      this.logger.error(`Error copying contract with survey images: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy contract files with survey images',
        error: error.message
      };
    }
  }

  /**
   * Copy proposal files to OneDrive quotations folder
   */
  async copyProposalToQuotations(
    opportunityId: string,
    customerName: string,
    proposalFiles: { pptxPath?: string; pdfPath?: string },
    surveyImages?: Array<{
      fieldName: string;
      fileName: string;
      filePath: string;
      originalName?: string;
    }>
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      this.logger.log(`Copying proposal files to quotations folder for opportunity: ${opportunityId}`);

      // Create customer folder name
      const folderName = this.createFolderName(customerName, opportunityId);
      const customerFolderPath = path.join(this.quotationsFolder, folderName);

      // Create the customer folder
      await this.createCustomerFolder(customerFolderPath);

      // Copy proposal files
      const copiedFiles: string[] = [];
      
      if (proposalFiles.pptxPath && fs.existsSync(proposalFiles.pptxPath)) {
        const pptxFileName = path.basename(proposalFiles.pptxPath);
        const pptxDestination = path.join(customerFolderPath, pptxFileName);
        await this.copyFile(proposalFiles.pptxPath, pptxDestination);
        copiedFiles.push(pptxFileName);
      }

      if (proposalFiles.pdfPath && fs.existsSync(proposalFiles.pdfPath)) {
        const pdfFileName = path.basename(proposalFiles.pdfPath);
        const pdfDestination = path.join(customerFolderPath, pdfFileName);
        await this.copyFile(proposalFiles.pdfPath, pdfDestination);
        copiedFiles.push(pdfFileName);
      }

      if (copiedFiles.length === 0) {
        return {
          success: false,
          message: 'No proposal files found to copy',
          error: 'No valid proposal files provided'
        };
      }

      // Copy survey images if provided
      let surveyImagesResult: { success: boolean; message: string; copiedImages?: string[]; error?: string } | null = null;
      if (surveyImages && surveyImages.length > 0) {
        this.logger.log(`Copying ${surveyImages.length} survey images to customer folder`);
        surveyImagesResult = await this.copySurveyImagesToCustomerFolder(customerFolderPath, surveyImages);
      }

      this.logger.log(`Successfully copied ${copiedFiles.length} proposal files to: ${customerFolderPath}`);

      const message = surveyImagesResult 
        ? `Successfully copied ${copiedFiles.length} proposal files and ${surveyImagesResult.copiedImages?.length || 0} survey images to quotations folder`
        : `Successfully copied ${copiedFiles.length} proposal files to quotations folder`;

      return {
        success: true,
        message,
        folderPath: customerFolderPath
      };

    } catch (error) {
      this.logger.error(`Error copying proposal to quotations: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy proposal files to quotations folder',
        error: error.message
      };
    }
  }

  /**
   * Copy contract files to OneDrive orders folder
   */
  async copyContractToOrders(
    opportunityId: string,
    customerName: string,
    contractFiles: { pptxPath?: string; pdfPath?: string },
    surveyImages?: Array<{
      fieldName: string;
      fileName: string;
      filePath: string;
      originalName?: string;
    }>
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      this.logger.log(`Copying contract files to orders folder for opportunity: ${opportunityId}`);

      // Create customer folder name
      const folderName = this.createFolderName(customerName, opportunityId);
      const customerFolderPath = path.join(this.ordersFolder, folderName);

      // Create the customer folder
      await this.createCustomerFolder(customerFolderPath);

      // Copy contract files
      const copiedFiles: string[] = [];
      
      if (contractFiles.pptxPath && fs.existsSync(contractFiles.pptxPath)) {
        const pptxFileName = path.basename(contractFiles.pptxPath);
        const pptxDestination = path.join(customerFolderPath, pptxFileName);
        await this.copyFile(contractFiles.pptxPath, pptxDestination);
        copiedFiles.push(pptxFileName);
      }

      if (contractFiles.pdfPath && fs.existsSync(contractFiles.pdfPath)) {
        const pdfFileName = path.basename(contractFiles.pdfPath);
        const pdfDestination = path.join(customerFolderPath, pdfFileName);
        await this.copyFile(contractFiles.pdfPath, pdfDestination);
        copiedFiles.push(pdfFileName);
      }

      if (copiedFiles.length === 0) {
        return {
          success: false,
          message: 'No contract files found to copy',
          error: 'No valid contract files provided'
        };
      }

      // Copy survey images if provided
      let surveyImagesResult: { success: boolean; message: string; copiedImages?: string[]; error?: string } | null = null;
      if (surveyImages && surveyImages.length > 0) {
        this.logger.log(`Copying ${surveyImages.length} survey images to customer folder`);
        surveyImagesResult = await this.copySurveyImagesToCustomerFolder(customerFolderPath, surveyImages);
      }

      this.logger.log(`Successfully copied ${copiedFiles.length} contract files to: ${customerFolderPath}`);

      const message = surveyImagesResult 
        ? `Successfully copied ${copiedFiles.length} contract files and ${surveyImagesResult.copiedImages?.length || 0} survey images to orders folder`
        : `Successfully copied ${copiedFiles.length} contract files to orders folder`;

      return {
        success: true,
        message,
        folderPath: customerFolderPath
      };

    } catch (error) {
      this.logger.error(`Error copying contract to orders: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy contract files to orders folder',
        error: error.message
      };
    }
  }

  /**
   * Copy disclaimer files to OneDrive orders folder
   */
  async copyDisclaimerToOrders(
    opportunityId: string,
    customerName: string,
    disclaimerPath: string
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      this.logger.log(`Copying disclaimer file to orders folder for opportunity: ${opportunityId}`);

      // Create customer folder name
      const folderName = this.createFolderName(customerName, opportunityId);
      const customerFolderPath = path.join(this.ordersFolder, folderName);

      // Create the customer folder
      await this.createCustomerFolder(customerFolderPath);

      // Copy disclaimer file
      if (fs.existsSync(disclaimerPath)) {
        const disclaimerFileName = path.basename(disclaimerPath);
        const disclaimerDestination = path.join(customerFolderPath, disclaimerFileName);
        await this.copyFile(disclaimerPath, disclaimerDestination);

        this.logger.log(`Successfully copied disclaimer file: ${disclaimerFileName}`);
        return {
          success: true,
          message: `Successfully copied disclaimer file to orders folder`,
          folderPath: customerFolderPath
        };
      } else {
        this.logger.warn(`Disclaimer file not found: ${disclaimerPath}`);
        return {
          success: false,
          message: `Disclaimer file not found: ${disclaimerPath}`,
          error: 'File not found'
        };
      }

    } catch (error) {
      this.logger.error(`Error copying disclaimer file: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy disclaimer file to OneDrive',
        error: error.message
      };
    }
  }

  /**
   * Upload signed document and audit log to OneDrive orders folder
   * Downloads from DocuSeal and saves to OneDrive
   */
  async uploadSignedDocumentAndAuditLogToOrders(
    opportunityId: string,
    customerName: string,
    signedDocumentBuffer: Buffer,
    auditLogBuffer: Buffer | null,
    documentType: 'CONTRACT' | 'DISCLAIMER' | 'BOOKING_CONFIRMATION' | 'EXPRESS_CONSENT',
    submissionId: string
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      this.logger.log(`Uploading signed document and audit log to orders folder for opportunity: ${opportunityId}`);

      // Try to find existing folder by opportunityId first (folders are created by survey images system)
      // Trim opportunityId to handle any whitespace issues
      const trimmedOpportunityId = opportunityId.trim();
      this.logger.log(`🔍 Searching for existing folder with opportunityId: "${trimmedOpportunityId}"`);
      let customerFolderPath: string | null = this.findExistingFolderByOpportunityId(trimmedOpportunityId);
      
      if (customerFolderPath) {
        this.logger.log(`✅ Using existing folder: ${customerFolderPath}`);
      } else {
        // If no existing folder found, create one using the standard naming convention
        const folderName = this.createFolderName(customerName, opportunityId);
        customerFolderPath = path.join(this.ordersFolder, folderName);
        this.logger.warn(`⚠️ No existing folder found. Creating new folder: ${customerFolderPath}`);
        this.logger.warn(`⚠️ This may create a duplicate folder. Please check if a folder with opportunityId "${opportunityId}" already exists.`);
      }

      // Create the customer folder if it doesn't exist
      await this.createCustomerFolder(customerFolderPath);

      // Determine file names based on document type
      const documentTypeName = documentType.toLowerCase().replace('_', '-');
      const signedDocFileName = `Signed_${documentTypeName}_${opportunityId}_${submissionId}.pdf`;
      const auditLogFileName = `Audit_Log_${documentTypeName}_${opportunityId}_${submissionId}.pdf`;

      const signedDocPath = path.join(customerFolderPath, signedDocFileName);
      const auditLogPath = auditLogBuffer ? path.join(customerFolderPath, auditLogFileName) : null;

      // Check if files already exist (to avoid overwriting if webhook is called multiple times)
      if (fs.existsSync(signedDocPath)) {
        this.logger.log(`⚠️ Signed document already exists for submission ${submissionId}, skipping upload to preserve existing file`);
      } else {
        // Write signed document to OneDrive
        fs.writeFileSync(signedDocPath, signedDocumentBuffer);
        this.logger.log(`✅ Successfully uploaded signed document: ${signedDocFileName}`);
      }

      // Write audit log to OneDrive if available
      if (auditLogBuffer && auditLogPath) {
        if (fs.existsSync(auditLogPath)) {
          this.logger.log(`⚠️ Audit log already exists for submission ${submissionId}, skipping upload to preserve existing file`);
        } else {
          fs.writeFileSync(auditLogPath, auditLogBuffer);
          this.logger.log(`✅ Successfully uploaded audit log: ${auditLogFileName}`);
        }
      } else {
        this.logger.warn(`⚠️ Audit log not available for submission ${submissionId}`);
      }

      return {
        success: true,
        message: `Successfully uploaded signed document${auditLogBuffer ? ' and audit log' : ''} to orders folder`,
        folderPath: customerFolderPath
      };

    } catch (error) {
      this.logger.error(`Error uploading signed document and audit log: ${error.message}`);
      return {
        success: false,
        message: 'Failed to upload signed document and audit log to OneDrive',
        error: error.message
      };
    }
  }

  /**
   * Copy email confirmation files to OneDrive orders folder
   */
  async copyEmailConfirmationToOrders(
    opportunityId: string,
    customerName: string,
    emailConfirmationPath: string
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      this.logger.log(`Copying email confirmation file to orders folder for opportunity: ${opportunityId}`);

      // Create customer folder name
      const folderName = this.createFolderName(customerName, opportunityId);
      const customerFolderPath = path.join(this.ordersFolder, folderName);

      // Create the customer folder
      await this.createCustomerFolder(customerFolderPath);

      // Copy email confirmation file
      if (fs.existsSync(emailConfirmationPath)) {
        const emailConfirmationFileName = path.basename(emailConfirmationPath);
        const emailConfirmationDestination = path.join(customerFolderPath, emailConfirmationFileName);
        await this.copyFile(emailConfirmationPath, emailConfirmationDestination);

        this.logger.log(`Successfully copied email confirmation file: ${emailConfirmationFileName}`);
        return {
          success: true,
          message: `Successfully copied email confirmation file to orders folder`,
          folderPath: customerFolderPath
        };
      } else {
        this.logger.warn(`Email confirmation file not found: ${emailConfirmationPath}`);
        return {
          success: false,
          message: `Email confirmation file not found: ${emailConfirmationPath}`,
          error: 'File not found'
        };
      }

    } catch (error) {
      this.logger.error(`Error copying email confirmation file: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy email confirmation file to OneDrive',
        error: error.message
      };
    }
  }

  /**
   * Copy both proposal and contract files to their respective OneDrive folders
   */
  async copyFilesToOneDrive(
    opportunityId: string,
    customerName: string,
    proposalFiles: { pptxPath?: string; pdfPath?: string },
    contractFiles?: { pptxPath?: string; pdfPath?: string }
  ): Promise<{
    success: boolean;
    message: string;
    error?: string;
    results: {
      quotations: { success: boolean; message: string; folderPath?: string };
      orders?: { success: boolean; message: string; folderPath?: string };
    };
  }> {
    try {
      this.logger.log(`Starting OneDrive file copy process for opportunity: ${opportunityId}`);

      // Copy proposal files to quotations folder - with automatic survey images
      const quotationsResult = await this.copyProposalToQuotationsWithSurveyImages(opportunityId, customerName, proposalFiles);

      // Copy contract files to orders folder (if provided) - with automatic survey images
      let ordersResult;
      if (contractFiles) {
        ordersResult = await this.copyContractToOrdersWithSurveyImages(opportunityId, customerName, contractFiles);
      }

      const allSuccessful = quotationsResult.success && (!contractFiles || ordersResult?.success);

      return {
        success: allSuccessful,
        message: allSuccessful 
          ? 'Successfully copied all files to OneDrive folders'
          : 'Some files failed to copy to OneDrive folders',
        results: {
          quotations: quotationsResult,
          ...(ordersResult && { orders: ordersResult })
        }
      };

    } catch (error) {
      this.logger.error(`Error in OneDrive file copy process: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy files to OneDrive folders',
        error: error.message,
        results: {
          quotations: { success: false, message: 'Failed to copy proposal files' }
        }
      };
    }
  }

  /**
   * Copy all won opportunity documents to OneDrive (proposal, contract, disclaimer, email confirmation)
   */
  async copyWonOpportunityDocumentsToOneDrive(
    opportunityId: string,
    customerName: string,
    documents: {
      proposalFiles?: { pptxPath?: string; pdfPath?: string };
      contractFiles?: { pptxPath?: string; pdfPath?: string };
      disclaimerPath?: string;
      emailConfirmationPath?: string;
    }
  ): Promise<{
    success: boolean;
    message: string;
    error?: string;
    results: {
      quotations?: { success: boolean; message: string; folderPath?: string };
      orders?: { success: boolean; message: string; folderPath?: string };
      disclaimer?: { success: boolean; message: string; folderPath?: string };
      emailConfirmation?: { success: boolean; message: string; folderPath?: string };
    };
  }> {
    try {
      this.logger.log(`Starting OneDrive document copy process for won opportunity: ${opportunityId}`);

      const results: any = {};
      let allSuccessful = true;

      // Copy proposal files to quotations folder (if provided) - with automatic survey images
      if (documents.proposalFiles) {
        const quotationsResult = await this.copyProposalToQuotationsWithSurveyImages(opportunityId, customerName, documents.proposalFiles);
        results.quotations = quotationsResult;
        if (!quotationsResult.success) allSuccessful = false;
      }

      // Copy contract files to orders folder (if provided) - with automatic survey images
      if (documents.contractFiles) {
        const ordersResult = await this.copyContractToOrdersWithSurveyImages(opportunityId, customerName, documents.contractFiles);
        results.orders = ordersResult;
        if (!ordersResult.success) allSuccessful = false;
      }

      // Skip disclaimer file - signed disclaimers are now handled by DocuSeal webhooks
      // Do NOT copy disclaimer files as they are automatically uploaded when signed via DocuSeal
      if (documents.disclaimerPath) {
        this.logger.log(`⚠️ Skipping disclaimer file copy - signed disclaimers are handled by DocuSeal webhooks`);
      }

      // Skip email confirmation file - signed confirmations are now handled by DocuSeal webhooks
      // Do NOT copy email confirmation files as they are automatically uploaded when signed via DocuSeal
      if (documents.emailConfirmationPath) {
        this.logger.log(`⚠️ Skipping email confirmation file copy - signed confirmations are handled by DocuSeal webhooks`);
      }

      return {
        success: allSuccessful,
        message: allSuccessful 
          ? 'Successfully copied all won opportunity documents to OneDrive folders'
          : 'Some documents failed to copy to OneDrive folders',
        results
      };

    } catch (error) {
      this.logger.error(`Error in OneDrive won opportunity document copy process: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy won opportunity documents to OneDrive folders',
        error: error.message,
        results: {}
      };
    }
  }


  /**
   * Create a folder name from customer name and opportunity ID
   */
  private createFolderName(customerName: string, opportunityId: string): string {
    // Clean and validate customer name
    let cleanName = '';
    if (customerName && !this.isUnknownOrEmpty(customerName)) {
      cleanName = customerName
        .replace(/[<>:"/\\|?*]/g, '') // Remove invalid characters
        .replace(/\s+/g, ' ') // Replace multiple spaces with single space
        .trim()
        .substring(0, 50); // Limit to 50 characters
    }

    // If we have a valid customer name, use it; otherwise use opportunity ID only
    if (cleanName) {
      return `${cleanName} - ${opportunityId}`;
    } else {
      // Fallback: use opportunity ID only if customer name is missing
      return `Opportunity ${opportunityId}`;
    }
  }

  /**
   * Create customer folder if it doesn't exist
   */
  private async createCustomerFolder(folderPath: string): Promise<void> {
    try {
      if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath, { recursive: true });
        this.logger.log(`Created customer folder: ${folderPath}`);
      } else {
        this.logger.log(`Customer folder already exists: ${folderPath}`);
      }
    } catch (error) {
      this.logger.error(`Error creating customer folder: ${error.message}`);
      throw new Error(`Failed to create customer folder: ${error.message}`);
    }
  }

  /**
   * Copy a file from source to destination
   */
  private async copyFile(sourcePath: string, destinationPath: string): Promise<void> {
    try {
      // Ensure destination directory exists
      const destDir = path.dirname(destinationPath);
      if (!fs.existsSync(destDir)) {
        fs.mkdirSync(destDir, { recursive: true });
      }

      // Copy the file
      fs.copyFileSync(sourcePath, destinationPath);
      this.logger.log(`Copied file: ${sourcePath} -> ${destinationPath}`);
    } catch (error) {
      this.logger.error(`Error copying file: ${error.message}`);
      throw new Error(`Failed to copy file: ${error.message}`);
    }
  }

  /**
   * Get file paths from presentation service output
   */
  getFilePathsFromPresentationResult(
    presentationResult: any,
    outputDir: string
  ): { pptxPath?: string; pdfPath?: string } {
    const filePaths: { pptxPath?: string; pdfPath?: string } = {};

    if (presentationResult.pptxFile) {
      filePaths.pptxPath = path.join(outputDir, presentationResult.pptxFile);
    }

    if (presentationResult.pdfFile) {
      filePaths.pdfPath = path.join(outputDir, presentationResult.pdfFile);
    }

    return filePaths;
  }

  /**
   * Verify OneDrive folder paths exist
   */
  async verifyOneDrivePaths(): Promise<{ quotations: boolean; orders: boolean }> {
    const quotationsExists = fs.existsSync(this.quotationsFolder);
    const ordersExists = fs.existsSync(this.ordersFolder);

    this.logger.log(`OneDrive path verification - Quotations: ${quotationsExists}, Orders: ${ordersExists}`);

    return {
      quotations: quotationsExists,
      orders: ordersExists
    };
  }

  /**
   * Helper function to check if a value is empty or represents "unknown"
   */
  private isUnknownOrEmpty(value: string | null | undefined): boolean {
    if (!value) return true;
    const normalized = value.trim().toLowerCase();
    return normalized === '' || 
           normalized === 'unknown' || 
           normalized === 'unknown customer' ||
           normalized === 'unknownpos' ||
           normalized === 'unknown pos' ||
           normalized === 'unknownposition' ||
           normalized === 'unknown position';
  }

  /**
   * Fetch customer details from CalculatorProgress
   */
  private async fetchCustomerDetailsFromCalculator(opportunityId: string): Promise<{ customerName: string | null; postcode: string | null }> {
    try {
      // Try to find any calculator progress for this opportunity
      const calculatorProgress = await this.prisma.calculatorProgress.findFirst({
        where: {
          opportunityId: opportunityId
        },
        orderBy: {
          updatedAt: 'desc'
        }
      });

      if (calculatorProgress && calculatorProgress.data) {
        const data = calculatorProgress.data as any;
        if (data.customerDetails) {
          const customerDetails = data.customerDetails;
          return {
            customerName: customerDetails.customerName || null,
            postcode: customerDetails.postcode || null
          };
        }
      }
    } catch (error) {
      this.logger.warn(`Failed to fetch customer details from CalculatorProgress: ${error.message}`);
    }
    return { customerName: null, postcode: null };
  }

  /**
   * Fetch customer details from OpportunityProgress
   */
  private async fetchCustomerDetailsFromOpportunityProgress(opportunityId: string): Promise<{ customerName: string | null; postcode: string | null }> {
    try {
      const opportunityProgress = await this.prisma.opportunityProgress.findUnique({
        where: {
          ghlOpportunityId: opportunityId
        },
        select: {
          contactAddress: true,
          contactPostcode: true
        }
      });

      if (opportunityProgress) {
        return {
          customerName: opportunityProgress.contactAddress || null,
          postcode: opportunityProgress.contactPostcode || null
        };
      }
    } catch (error) {
      this.logger.warn(`Failed to fetch customer details from OpportunityProgress: ${error.message}`);
    }
    return { customerName: null, postcode: null };
  }

  /**
   * Organize files based on appointment outcome (won/lost)
   */
  async organizeFilesByOutcome(
    opportunityId: string,
    customerName: string,
    postcode: string,
    outcome: 'won' | 'lost',
    files: {
      surveyPath?: string;
      calculatorPath?: string;
      contractPath?: string;
      proposalPath?: string;
      disclaimerPath?: string;
      emailConfirmationPath?: string;
    },
    userId?: string
  ): Promise<{ success: boolean; message: string; error?: string; folderPath?: string }> {
    try {
      this.logger.log(`Organizing files by outcome (${outcome}) for opportunity: ${opportunityId}`);
      this.logger.log(`Files provided:`, files);
      this.logger.log(`Initial customer details - Name: "${customerName}", Postcode: "${postcode}"`);

      // Auto-fetch customer details if provided values are invalid/missing
      let finalCustomerName = customerName;
      let finalPostcode = postcode;

      if (this.isUnknownOrEmpty(customerName) || this.isUnknownOrEmpty(postcode)) {
        this.logger.log(`Customer details are missing or invalid, attempting to fetch from alternative sources...`);

        // Try to fetch from CalculatorProgress first (most reliable)
        const calcDetails = await this.fetchCustomerDetailsFromCalculator(opportunityId);
        if (calcDetails.customerName && !this.isUnknownOrEmpty(calcDetails.customerName)) {
          finalCustomerName = calcDetails.customerName;
          this.logger.log(`✅ Found customer name from CalculatorProgress: "${finalCustomerName}"`);
        }
        if (calcDetails.postcode && !this.isUnknownOrEmpty(calcDetails.postcode)) {
          finalPostcode = calcDetails.postcode;
          this.logger.log(`✅ Found postcode from CalculatorProgress: "${finalPostcode}"`);
        }

        // If still missing, try OpportunityProgress
        if (this.isUnknownOrEmpty(finalCustomerName) || this.isUnknownOrEmpty(finalPostcode)) {
          const oppDetails = await this.fetchCustomerDetailsFromOpportunityProgress(opportunityId);
          if (this.isUnknownOrEmpty(finalCustomerName) && oppDetails.customerName && !this.isUnknownOrEmpty(oppDetails.customerName)) {
            finalCustomerName = oppDetails.customerName;
            this.logger.log(`✅ Found customer name from OpportunityProgress: "${finalCustomerName}"`);
          }
          if (this.isUnknownOrEmpty(finalPostcode) && oppDetails.postcode && !this.isUnknownOrEmpty(oppDetails.postcode)) {
            finalPostcode = oppDetails.postcode;
            this.logger.log(`✅ Found postcode from OpportunityProgress: "${finalPostcode}"`);
          }
        }

        // If still missing and we have userId, try fetching from GHL API
        // Note: This would require injecting OpportunitiesService, which we can do later if needed
        if ((this.isUnknownOrEmpty(finalCustomerName) || this.isUnknownOrEmpty(finalPostcode)) && userId) {
          this.logger.log(`⚠️ Customer details still missing. GHL API fetch would require additional setup.`);
        }

        this.logger.log(`Final customer details - Name: "${finalCustomerName}", Postcode: "${finalPostcode}"`);
      }

      // Determine target folder based on outcome
      const targetFolder = outcome === 'won' ? this.ordersFolder : this.quotationsFolder;
      
      // Create customer folder name with postcode
      const folderName = this.createOutcomeFolderName(finalCustomerName, finalPostcode, opportunityId);
      const customerFolderPath = path.join(targetFolder, folderName);

      // Create the customer folder
      await this.createCustomerFolder(customerFolderPath);

      // Copy files to the folder
      const copiedFiles: string[] = [];
      
      // Copy survey files
      if (files.surveyPath) {
        const absoluteSurveyPath = this.resolveAbsolutePath(files.surveyPath);
        this.logger.log(`Checking survey path: ${absoluteSurveyPath}, exists: ${fs.existsSync(absoluteSurveyPath)}`);
        if (fs.existsSync(absoluteSurveyPath)) {
          // For survey reports, copy only files for this specific opportunity
          const surveyFiles = await this.getSurveyFilesForOpportunity(absoluteSurveyPath, opportunityId);
          for (const surveyFile of surveyFiles) {
            const fileName = path.basename(surveyFile);
            const destination = path.join(customerFolderPath, fileName);
            await this.copyFile(surveyFile, destination);
            copiedFiles.push(fileName);
          }
        }
      }

      // Copy survey images (automatically fetch and copy all survey images for this opportunity)
      try {
        this.logger.log(`🔍 [SURVEY_IMAGES] Starting survey image copy process for opportunity: ${opportunityId}`);
        const surveyImages = await this.getSurveyImagesForOpportunity(opportunityId);
        this.logger.log(`🔍 [SURVEY_IMAGES] Retrieved ${surveyImages.length} survey images for opportunity: ${opportunityId}`);
        
        if (surveyImages && surveyImages.length > 0) {
          this.logger.log(`📷 [SURVEY_IMAGES] Found ${surveyImages.length} survey images for opportunity: ${opportunityId}`);
          this.logger.log(`📷 [SURVEY_IMAGES] Image details:`, surveyImages.map((img, index) => ({
            index,
            fieldName: img.fieldName,
            fileName: img.fileName,
            filePath: img.filePath?.substring(0, 100) + '...'
          })));
          
          const surveyImagesResult = await this.copySurveyImagesToCustomerFolder(customerFolderPath, surveyImages);
          this.logger.log(`📷 [SURVEY_IMAGES] Copy result:`, surveyImagesResult);
          
          if (surveyImagesResult.success && surveyImagesResult.copiedImages) {
            copiedFiles.push(...surveyImagesResult.copiedImages.map(img => `survey_images/${img}`));
            this.logger.log(`✅ [SURVEY_IMAGES] Successfully copied ${surveyImagesResult.copiedImages.length} survey images`);
          } else {
            this.logger.error(`❌ [SURVEY_IMAGES] Failed to copy survey images: ${surveyImagesResult.error}`);
          }
        } else {
          this.logger.log(`⚠️ [SURVEY_IMAGES] No survey images found for opportunity: ${opportunityId}`);
        }
      } catch (error) {
        this.logger.error(`❌ [SURVEY_IMAGES] Error copying survey images for opportunity ${opportunityId}: ${error.message}`);
        this.logger.error(`❌ [SURVEY_IMAGES] Error stack:`, error.stack);
        // Don't fail the entire operation if survey images fail
      }

      // Copy calculator files (get latest modified only)
      if (files.calculatorPath) {
        const absoluteCalculatorPath = this.resolveAbsolutePath(files.calculatorPath);
        this.logger.log(`Checking calculator path: ${absoluteCalculatorPath}, exists: ${fs.existsSync(absoluteCalculatorPath)}`);
        
        // Check both epvs-opportunities and opportunities directories
        const calculatorPaths = [
          absoluteCalculatorPath,
          absoluteCalculatorPath.replace('epvs-opportunities', 'opportunities'),
          absoluteCalculatorPath.replace('opportunities', 'epvs-opportunities')
        ];
        
        let latestCalcFile: string | null = null;
        for (const calcPath of calculatorPaths) {
          if (fs.existsSync(calcPath)) {
            this.logger.log(`Checking calculator path: ${calcPath}, exists: ${fs.existsSync(calcPath)}`);
            // Check if the path is a file or directory
            const stat = fs.statSync(calcPath);
            if (stat.isDirectory()) {
              latestCalcFile = await this.findLatestFileByOpportunityId(calcPath, opportunityId, ['.xlsm', '.xlsx']);
            } else if (stat.isFile() && calcPath.includes(opportunityId)) {
              // If it's a file and contains the opportunity ID, use it directly
              latestCalcFile = calcPath;
            }
            if (latestCalcFile) {
              this.logger.log(`Found latest file for opportunity ${opportunityId}: ${path.basename(latestCalcFile)} (modified: ${fs.statSync(latestCalcFile).mtime.toISOString()})`);
              break;
            }
          }
        }
        
        if (latestCalcFile) {
          const fileName = path.basename(latestCalcFile);
          const destination = path.join(customerFolderPath, fileName);
          await this.copyFile(latestCalcFile, destination);
          copiedFiles.push(fileName);
        }
      }

      // Copy contract files (get latest modified only)
      if (files.contractPath) {
        const absoluteContractPath = this.resolveAbsolutePath(files.contractPath);
        this.logger.log(`Checking contract path: ${absoluteContractPath}, exists: ${fs.existsSync(absoluteContractPath)}`);
        
        // Check both epvs-opportunities/pdfs and opportunities/pdfs directories
        const contractPaths = [
          absoluteContractPath,
          absoluteContractPath.replace('epvs-opportunities/pdfs', 'opportunities/pdfs'),
          absoluteContractPath.replace('opportunities/pdfs', 'epvs-opportunities/pdfs')
        ];
        
        let latestContractFile: string | null = null;
        for (const contractPath of contractPaths) {
          if (fs.existsSync(contractPath)) {
            this.logger.log(`Checking contract path: ${contractPath}, exists: ${fs.existsSync(contractPath)}`);
            // Check if the path is a file or directory
            const stat = fs.statSync(contractPath);
            if (stat.isDirectory()) {
              latestContractFile = await this.findLatestFileByOpportunityId(contractPath, opportunityId, ['.pdf']);
            } else if (stat.isFile() && contractPath.includes(opportunityId)) {
              // If it's a file and contains the opportunity ID, use it directly
              latestContractFile = contractPath;
            }
            if (latestContractFile) {
              this.logger.log(`Found latest file for opportunity ${opportunityId}: ${path.basename(latestContractFile)} (modified: ${fs.statSync(latestContractFile).mtime.toISOString()})`);
              break;
            }
          }
        }
        
        if (latestContractFile) {
          const fileName = path.basename(latestContractFile);
          const destination = path.join(customerFolderPath, fileName);
          await this.copyFile(latestContractFile, destination);
          copiedFiles.push(fileName);
        }
      }

      // Copy proposal files (get latest modified only)
      if (files.proposalPath) {
        const absoluteProposalPath = this.resolveAbsolutePath(files.proposalPath);
        this.logger.log(`Checking proposal path: ${absoluteProposalPath}, exists: ${fs.existsSync(absoluteProposalPath)}`);
        if (fs.existsSync(absoluteProposalPath)) {
          // Check if the path is a file or directory
          const stat = fs.statSync(absoluteProposalPath);
          let latestPptxFile: string | null = null;
          let latestPdfFile: string | null = null;
          
          if (stat.isDirectory()) {
            // Search for the latest proposal files that contain the opportunity ID
            latestPptxFile = await this.findLatestFileByOpportunityId(absoluteProposalPath, opportunityId, ['.pptx', '.ppt']);
            latestPdfFile = await this.findLatestFileByOpportunityId(absoluteProposalPath, opportunityId, ['.pdf']);
          } else if (stat.isFile() && absoluteProposalPath.includes(opportunityId)) {
            // If it's a file and contains the opportunity ID, use it directly
            const ext = path.extname(absoluteProposalPath).toLowerCase();
            if (['.pptx', '.ppt'].includes(ext)) {
              latestPptxFile = absoluteProposalPath;
            } else if (ext === '.pdf') {
              latestPdfFile = absoluteProposalPath;
            }
          }
          
          // Copy latest PowerPoint file
          if (latestPptxFile) {
            const fileName = path.basename(latestPptxFile);
            const destination = path.join(customerFolderPath, fileName);
            await this.copyFile(latestPptxFile, destination);
            copiedFiles.push(fileName);
          }
          
          // Copy latest PDF file (if different from PowerPoint)
          if (latestPdfFile && latestPdfFile !== latestPptxFile) {
            const fileName = path.basename(latestPdfFile);
            const destination = path.join(customerFolderPath, fileName);
            await this.copyFile(latestPdfFile, destination);
            copiedFiles.push(fileName);
          }
        }
      }

      // Disclaimer and email confirmation files are handled by DocuSeal webhooks, not moved here

      if (copiedFiles.length === 0) {
        this.logger.warn(`No files found to copy for opportunity ${opportunityId}. Files provided:`, files);
        return {
          success: false,
          message: 'No files found to copy',
          error: 'No valid files provided or files do not exist at the specified paths'
        };
      }

      this.logger.log(`Successfully copied ${copiedFiles.length} files to: ${customerFolderPath}`);

      return {
        success: true,
        message: `Successfully organized ${copiedFiles.length} files for ${outcome} outcome`,
        folderPath: customerFolderPath
      };

    } catch (error) {
      this.logger.error(`Error organizing files by outcome: ${error.message}`);
      return {
        success: false,
        message: 'Failed to organize files by outcome',
        error: error.message
      };
    }
  }

  /**
   * Find existing folder by opportunityId in orders folder
   * Returns the full path if found, null otherwise
   * Folders are created by the survey images system and may have postcode at the beginning
   * Prefers folders with postcodes as they are more specific and correct
   */
  private findExistingFolderByOpportunityId(opportunityId: string): string | null {
    try {
      if (!fs.existsSync(this.ordersFolder)) {
        this.logger.warn(`Orders folder does not exist: ${this.ordersFolder}`);
        return null;
      }
      
      // Trim opportunityId to ensure no whitespace issues
      const trimmedOpportunityId = opportunityId.trim();
      
      // Read all folders in orders directory
      const folders = fs.readdirSync(this.ordersFolder, { withFileTypes: true })
        .filter(dirent => dirent.isDirectory())
        .map(dirent => dirent.name);

      this.logger.log(`🔍 Searching for folder with opportunityId: "${trimmedOpportunityId}" in ${folders.length} folders`);

      const opportunityIdUpper = trimmedOpportunityId.toUpperCase();
      const matchingFolders: Array<{ name: string; hasPostcode: boolean; path: string }> = [];

      // Collect all matching folders first
      for (const folderName of folders) {
        const trimmedFolderName = folderName.trim();
        const folderNameUpper = trimmedFolderName.toUpperCase();
        
        // Check if folder ends with " - opportunityId" (exact or case-insensitive)
        const exactMatch = trimmedFolderName.endsWith(` - ${trimmedOpportunityId}`);
        const caseInsensitiveMatch = folderNameUpper.endsWith(` - ${opportunityIdUpper}`);
        
        if (exactMatch || caseInsensitiveMatch) {
          const fullPath = path.join(this.ordersFolder, folderName);
          
          // Check if folder has postcode (postcodes are typically 5-8 alphanumeric characters)
          // Look for patterns like: "Tf75dj, " or "Tf75dj " at the start, or postcode-like strings
          // UK postcodes are typically 5-7 characters (e.g., "SW1A1AA", "M1 1AA", "Tf75dj")
          const hasPostcode = /^[A-Z0-9]{2,8}[,\s]/.test(trimmedFolderName) || 
                             /[A-Z0-9]{5,8}[,\s]/.test(trimmedFolderName);
          
          matchingFolders.push({
            name: folderName,
            hasPostcode,
            path: fullPath
          });
        }
      }

      // If we found matching folders, prefer the one with postcode
      if (matchingFolders.length > 0) {
        // Sort: folders with postcode first
        matchingFolders.sort((a, b) => {
          if (a.hasPostcode && !b.hasPostcode) return -1;
          if (!a.hasPostcode && b.hasPostcode) return 1;
          return 0;
        });

        const selectedFolder = matchingFolders[0];
        this.logger.log(`✅ Found ${matchingFolders.length} matching folder(s) for opportunityId "${trimmedOpportunityId}"`);
        this.logger.log(`✅ Selected folder (${selectedFolder.hasPostcode ? 'with' : 'without'} postcode): ${selectedFolder.name}`);
        
        if (matchingFolders.length > 1) {
          this.logger.warn(`⚠️ Multiple folders found for opportunityId "${trimmedOpportunityId}":`);
          matchingFolders.forEach((f, idx) => {
            this.logger.warn(`  ${idx + 1}. ${f.name} ${f.hasPostcode ? '(has postcode)' : '(no postcode)'}`);
          });
          this.logger.warn(`⚠️ Using the first one (preferring postcode folder if available)`);
        }
        
        return selectedFolder.path;
      }

      // If still no match, try to find folder that contains the opportunityId anywhere (as last resort)
      for (const folderName of folders) {
        const folderNameUpper = folderName.toUpperCase();
        
        if (folderNameUpper.includes(opportunityIdUpper)) {
          // Make sure it's at the end or after a separator
          const lastIndex = folderNameUpper.lastIndexOf(opportunityIdUpper);
          const afterMatch = folderNameUpper.substring(lastIndex + opportunityIdUpper.length);
          
          // If the opportunityId is at the end or followed by nothing/whitespace, it's a match
          if (afterMatch.trim().length === 0 || folderName.toUpperCase().endsWith(opportunityIdUpper)) {
            const fullPath = path.join(this.ordersFolder, folderName);
            this.logger.log(`✅ Found matching folder (contains opportunityId): ${folderName} (opportunityId: ${trimmedOpportunityId})`);
            return fullPath;
          }
        }
      }

      this.logger.warn(`❌ No existing folder found for opportunityId "${trimmedOpportunityId}". Searched ${folders.length} folders.`);
      if (folders.length > 0) {
        // Show folders that might be related (contain part of the opportunityId)
        const relatedFolders = folders.filter(f => f.toUpperCase().includes(trimmedOpportunityId.substring(0, 5).toUpperCase()));
        if (relatedFolders.length > 0) {
          this.logger.warn(`Related folders found: ${relatedFolders.slice(0, 5).join(', ')}`);
        }
      }
      return null;
    } catch (error) {
      this.logger.error(`Error finding existing folder: ${error.message}`);
      return null;
    }
  }

  /**
   * Create folder name for outcome-based organization
   */
  private createOutcomeFolderName(customerName: string, postcode: string, opportunityId: string): string {
    // Clean and validate customer name
    let cleanName = '';
    if (customerName && !this.isUnknownOrEmpty(customerName)) {
      cleanName = customerName
        .replace(/[<>:"/\\|?*]/g, '')
        .replace(/\s+/g, ' ')
        .trim()
        .substring(0, 30);
    }

    // Clean and validate postcode
    let cleanPostcode = '';
    if (postcode && !this.isUnknownOrEmpty(postcode)) {
      cleanPostcode = postcode
        .replace(/[<>:"/\\|?*]/g, '')
        .replace(/\s+/g, '')
        .trim()
        .substring(0, 10);
    }

    // Build folder name with consistent format
    // Format: [CustomerName] [Postcode] - [OpportunityId]
    // Only include parts that are valid
    const parts: string[] = [];
    
    if (cleanName) {
      parts.push(cleanName);
    }
    
    if (cleanPostcode) {
      parts.push(cleanPostcode);
    }
    
    // If we have at least one valid part, use it; otherwise just use opportunity ID
    if (parts.length > 0) {
      return `${parts.join(' ')} - ${opportunityId}`;
    } else {
      // Fallback: use opportunity ID only if both customer name and postcode are missing
      return `Opportunity ${opportunityId}`;
    }
  }

  /**
   * Get latest survey files from survey directory
   */
  private async getLatestSurveyFiles(surveyPath: string): Promise<string[]> {
    const surveyDir = path.dirname(surveyPath);
    const opportunityId = path.basename(surveyDir);
    
    if (!fs.existsSync(surveyDir)) {
      return [];
    }

    const files: string[] = [];
    
    // Get all files in the survey directory
    const items = fs.readdirSync(surveyDir);
    
    for (const item of items) {
      const itemPath = path.join(surveyDir, item);
      const stat = fs.statSync(itemPath);
      
      if (stat.isFile()) {
        files.push(itemPath);
      } else if (stat.isDirectory()) {
        // Recursively get files from subdirectories
        const subFiles = this.getAllFilesInDirectory(itemPath);
        files.push(...subFiles);
      }
    }

    return files;
  }

  /**
   * Get survey files for a specific opportunity (only files that contain the opportunity ID)
   */
  private async getSurveyFilesForOpportunity(surveyPath: string, opportunityId: string): Promise<string[]> {
    const surveyDir = path.dirname(surveyPath);
    
    if (!fs.existsSync(surveyDir)) {
      return [];
    }

    const files: string[] = [];
    
    // Get all files in the survey directory
    const items = fs.readdirSync(surveyDir);
    
    for (const item of items) {
      const itemPath = path.join(surveyDir, item);
      const stat = fs.statSync(itemPath);
      
      if (stat.isFile()) {
        // Only include files that contain the opportunity ID
        if (item.includes(opportunityId)) {
          files.push(itemPath);
        }
      } else if (stat.isDirectory()) {
        // Recursively get files from subdirectories, but only if the directory name contains the opportunity ID
        if (item.includes(opportunityId)) {
          const subFiles = this.getAllFilesInDirectory(itemPath);
          files.push(...subFiles);
        }
      }
    }

    return files;
  }

  /**
   * Get latest calculator files (Excel files)
   */
  private async getLatestCalculatorFiles(calculatorPath: string): Promise<string[]> {
    const calcDir = path.dirname(calculatorPath);
    
    if (!fs.existsSync(calcDir)) {
      return [];
    }

    const excelFiles: { path: string; modified: Date }[] = [];
    
    // Look for Excel files in the directory
    const items = fs.readdirSync(calcDir);
    
    for (const item of items) {
      if (item.toLowerCase().endsWith('.xlsm') || item.toLowerCase().endsWith('.xlsx')) {
        const filePath = path.join(calcDir, item);
        const stat = fs.statSync(filePath);
        excelFiles.push({
          path: filePath,
          modified: stat.mtime
        });
      }
    }

    // Sort by modification date and return the latest
    excelFiles.sort((a, b) => b.modified.getTime() - a.modified.getTime());
    
    return excelFiles.length > 0 ? [excelFiles[0].path] : [];
  }

  /**
   * Get latest contract files (PDF files)
   */
  private async getLatestContractFiles(contractPath: string): Promise<string[]> {
    const contractDir = path.dirname(contractPath);
    
    if (!fs.existsSync(contractDir)) {
      return [];
    }

    const pdfFiles: { path: string; modified: Date }[] = [];
    
    // Look for PDF files in the directory
    const items = fs.readdirSync(contractDir);
    
    for (const item of items) {
      if (item.toLowerCase().endsWith('.pdf')) {
        const filePath = path.join(contractDir, item);
        const stat = fs.statSync(filePath);
        pdfFiles.push({
          path: filePath,
          modified: stat.mtime
        });
      }
    }

    // Sort by modification date and return the latest
    pdfFiles.sort((a, b) => b.modified.getTime() - a.modified.getTime());
    
    return pdfFiles.length > 0 ? [pdfFiles[0].path] : [];
  }

  /**
   * Get latest proposal files (PowerPoint and PDF files)
   */
  private async getLatestProposalFiles(proposalPath: string): Promise<string[]> {
    const proposalDir = path.dirname(proposalPath);
    
    if (!fs.existsSync(proposalDir)) {
      return [];
    }

    const proposalFiles: { path: string; modified: Date }[] = [];
    
    // Look for PowerPoint and PDF files in the directory
    const items = fs.readdirSync(proposalDir);
    
    for (const item of items) {
      if (item.toLowerCase().endsWith('.pptx') || item.toLowerCase().endsWith('.pdf')) {
        const filePath = path.join(proposalDir, item);
        const stat = fs.statSync(filePath);
        proposalFiles.push({
          path: filePath,
          modified: stat.mtime
        });
      }
    }

    // Sort by modification date and return the latest
    proposalFiles.sort((a, b) => b.modified.getTime() - a.modified.getTime());
    
    return proposalFiles.length > 0 ? [proposalFiles[0].path] : [];
  }

  /**
   * Get all files in a directory recursively
   */
  private getAllFilesInDirectory(dirPath: string): string[] {
    const files: string[] = [];
    
    if (!fs.existsSync(dirPath)) {
      return files;
    }

    const items = fs.readdirSync(dirPath);
    
    for (const item of items) {
      const itemPath = path.join(dirPath, item);
      const stat = fs.statSync(itemPath);
      
      if (stat.isFile()) {
        files.push(itemPath);
      } else if (stat.isDirectory()) {
        const subFiles = this.getAllFilesInDirectory(itemPath);
        files.push(...subFiles);
      }
    }

    return files;
  }

  /**
   * Resolve relative path to absolute path
   */
  private resolveAbsolutePath(relativePath: string): string {
    // If path is already absolute, return as is
    if (path.isAbsolute(relativePath)) {
      return relativePath;
    }
    
    // Resolve relative path from project root
    const projectRoot = process.cwd();
    return path.resolve(projectRoot, relativePath);
  }

  /**
   * Find files by opportunity ID in a directory
   */
  private async findFilesByOpportunityId(directoryPath: string, opportunityId: string, extensions: string[]): Promise<string[]> {
    try {
      if (!fs.existsSync(directoryPath)) {
        this.logger.warn(`Directory does not exist: ${directoryPath}`);
        return [];
      }

      const files = fs.readdirSync(directoryPath);
      const matchingFiles: string[] = [];

      for (const file of files) {
        const filePath = path.join(directoryPath, file);
        const stat = fs.statSync(filePath);
        
        if (stat.isFile()) {
          const fileExt = path.extname(file).toLowerCase();
          if (extensions.includes(fileExt) && file.includes(opportunityId)) {
            matchingFiles.push(filePath);
          }
        }
      }

      // Sort by modification time (newest first) and return the latest
      matchingFiles.sort((a, b) => {
        const statA = fs.statSync(a);
        const statB = fs.statSync(b);
        return statB.mtime.getTime() - statA.mtime.getTime();
      });

      this.logger.log(`Found ${matchingFiles.length} files for opportunity ${opportunityId} in ${directoryPath}`);
      return matchingFiles;
    } catch (error) {
      this.logger.error(`Error finding files by opportunity ID: ${error.message}`);
      return [];
    }
  }

  /**
   * Find the latest file by opportunity ID in a directory (returns only the most recent file)
   */
  private async findLatestFileByOpportunityId(directoryPath: string, opportunityId: string, extensions: string[]): Promise<string | null> {
    try {
      if (!fs.existsSync(directoryPath)) {
        this.logger.warn(`Directory does not exist: ${directoryPath}`);
        return null;
      }

      const files = fs.readdirSync(directoryPath);
      const matchingFiles: { filePath: string; mtime: Date }[] = [];

      for (const file of files) {
        const filePath = path.join(directoryPath, file);
        const stat = fs.statSync(filePath);
        
        if (stat.isFile()) {
          const fileExt = path.extname(file).toLowerCase();
          if (extensions.includes(fileExt)) {
            // More flexible matching - check if opportunity ID appears anywhere in filename
            const fileName = file.toLowerCase();
            const opportunityIdLower = opportunityId.toLowerCase();
            
            // Check various patterns for opportunity ID matching
            const matches = fileName.includes(opportunityIdLower) || 
                           fileName.includes(`_${opportunityIdLower}_`) ||
                           fileName.includes(`-${opportunityIdLower}-`) ||
                           fileName.includes(`_${opportunityIdLower}-`) ||
                           fileName.includes(`-${opportunityIdLower}_`) ||
                           fileName.endsWith(`_${opportunityIdLower}.${fileExt.substring(1)}`) ||
                           fileName.endsWith(`-${opportunityIdLower}.${fileExt.substring(1)}`);
            
            if (matches) {
              matchingFiles.push({ filePath, mtime: stat.mtime });
              this.logger.log(`Found matching file: ${file} for opportunity ${opportunityId}`);
            }
          }
        }
      }

      if (matchingFiles.length === 0) {
        this.logger.log(`No files found for opportunity ${opportunityId} in ${directoryPath}`);
        return null;
      }

      // Sort by modification time (newest first) and return only the latest
      matchingFiles.sort((a, b) => b.mtime.getTime() - a.mtime.getTime());
      const latestFile = matchingFiles[0].filePath;
      
      this.logger.log(`Found latest file for opportunity ${opportunityId}: ${path.basename(latestFile)} (modified: ${matchingFiles[0].mtime.toISOString()})`);
      return latestFile;
    } catch (error) {
      this.logger.error(`Error finding latest file by opportunity ID: ${error.message}`);
      return null;
    }
  }

  /**
   * Find the most recent PDF file in multiple directories (fallback when opportunity ID not in filename)
   */
  private async findMostRecentPdfInDirectories(directories: string[], opportunityId: string): Promise<string | null> {
    try {
      const allPdfFiles: { filePath: string; mtime: Date }[] = [];
      
      for (const directory of directories) {
        if (fs.existsSync(directory)) {
          const files = fs.readdirSync(directory);
          
          for (const file of files) {
            const filePath = path.join(directory, file);
            const stat = fs.statSync(filePath);
            
            if (stat.isFile() && path.extname(file).toLowerCase() === '.pdf') {
              // Only include files that might be related to this opportunity
              // Check if file was created recently (within last 30 days) or contains opportunity-related keywords
              const fileAge = Date.now() - stat.mtime.getTime();
              const thirtyDaysInMs = 30 * 24 * 60 * 60 * 1000;
              
              const fileName = file.toLowerCase();
              const isRecent = fileAge < thirtyDaysInMs;
              const hasOpportunityKeywords = fileName.includes('disclaimer') || 
                                           fileName.includes('email') || 
                                           fileName.includes('confirmation') ||
                                           fileName.includes('signed');
              
              if (isRecent || hasOpportunityKeywords) {
                allPdfFiles.push({ filePath, mtime: stat.mtime });
                this.logger.log(`Found potential PDF file: ${file} (age: ${Math.round(fileAge / (24 * 60 * 60 * 1000))} days)`);
              }
            }
          }
        }
      }
      
      if (allPdfFiles.length === 0) {
        this.logger.log(`No recent PDF files found in directories for opportunity: ${opportunityId}`);
        return null;
      }
      
      // Sort by modification time (newest first) and return the most recent
      allPdfFiles.sort((a, b) => b.mtime.getTime() - a.mtime.getTime());
      const mostRecentFile = allPdfFiles[0].filePath;
      
      this.logger.log(`Found most recent PDF file: ${path.basename(mostRecentFile)} (modified: ${allPdfFiles[0].mtime.toISOString()})`);
      return mostRecentFile;
    } catch (error) {
      this.logger.error(`Error finding most recent PDF in directories: ${error.message}`);
      return null;
    }
  }

  /**
   * Get folder statistics for OneDrive paths
   */
  async getOneDriveStats(): Promise<{
    quotations: { exists: boolean; folderCount?: number };
    orders: { exists: boolean; folderCount?: number };
  }> {
    const quotationsExists = fs.existsSync(this.quotationsFolder);
    const ordersExists = fs.existsSync(this.ordersFolder);

    let quotationsFolderCount;
    let ordersFolderCount;

    if (quotationsExists) {
      try {
        const quotationsContents = fs.readdirSync(this.quotationsFolder);
        quotationsFolderCount = quotationsContents.filter(item => 
          fs.statSync(path.join(this.quotationsFolder, item)).isDirectory()
        ).length;
      } catch (error) {
        this.logger.warn(`Could not read quotations folder: ${error.message}`);
      }
    }

    if (ordersExists) {
      try {
        const ordersContents = fs.readdirSync(this.ordersFolder);
        ordersFolderCount = ordersContents.filter(item => 
          fs.statSync(path.join(this.ordersFolder, item)).isDirectory()
        ).length;
      } catch (error) {
        this.logger.warn(`Could not read orders folder: ${error.message}`);
      }
    }

    return {
      quotations: { exists: quotationsExists, folderCount: quotationsFolderCount },
      orders: { exists: ordersExists, folderCount: ordersFolderCount }
    };
  }
}

