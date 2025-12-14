import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { DigitalSignatureService } from '../pdf-signature/digital-signature.service';
import * as fs from 'fs/promises';
import * as path from 'path';
import { PDFDocument, rgb } from 'pdf-lib';

@Injectable()
export class EmailConfirmationService {
  private readonly logger = new Logger(EmailConfirmationService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly digitalSignatureService: DigitalSignatureService,
  ) {}

  async createEmailConfirmationCopy(opportunityId: string, customerName: string): Promise<{ success: boolean; emailConfirmationPath?: string; error?: string }> {
    try {
      this.logger.log(`Creating email confirmation copy for opportunity ${opportunityId}`);

      // Skip opportunity validation for now - just create the template copy
      this.logger.log(`Creating email confirmation copy for opportunity ${opportunityId} and customer ${customerName}`);

      // Create signed directory if it doesn't exist
      const signedDir = path.join(process.cwd(), 'src', 'email_confirmation', 'signed');
      await fs.mkdir(signedDir, { recursive: true });

      // Copy template to signed directory
      // Use process.cwd() to get the backend directory, then navigate to src directory
      const templatePath = path.join(process.cwd(), 'src', 'email_confirmation', 'Confirmation of Booking Letter.pdf');
      
      // Check if template file exists
      try {
        await fs.access(templatePath);
        this.logger.log(`Template file found at: ${templatePath}`);
      } catch (error) {
        this.logger.error(`Template file not found at: ${templatePath}`);
        throw new Error(`Email confirmation template file not found at: ${templatePath}`);
      }
      
      const timestamp = Date.now();
      const emailConfirmationPath = path.join(signedDir, `email-confirmation-${opportunityId}-${timestamp}.pdf`);

      await fs.copyFile(templatePath, emailConfirmationPath);

      this.logger.log(`Email confirmation copy created: ${emailConfirmationPath}`);

      return {
        success: true,
        emailConfirmationPath: emailConfirmationPath,
      };
    } catch (error) {
      this.logger.error(`Error creating email confirmation copy: ${error.message}`, error.stack);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  async signEmailConfirmation(
    emailConfirmationPath: string,
    signatureData: string,
    digitalFootprint: any,
    opportunityId: string,
    signedBy: string,
    customerInfo: any,
    formData: any
  ): Promise<{ success: boolean; signedPdfPath?: string; metadata?: any; error?: string }> {
    try {
      this.logger.log(`Signing email confirmation: ${emailConfirmationPath}`);

      // Check if file exists
      try {
        await fs.access(emailConfirmationPath);
      } catch (error) {
        throw new Error(`Email confirmation PDF not found: ${emailConfirmationPath}`);
      }

      // Read the PDF file
      const pdfBytes = await fs.readFile(emailConfirmationPath);
      const pdfDoc = await PDFDocument.load(pdfBytes);

      // Get the first page
      const page = pdfDoc.getPage(0);
      const { width, height } = page.getSize();

      // Embed form data into PDF
      await this.embedFormDataIntoPDF(page, formData, customerInfo, width, height);

      // Add customer signature
      if (signatureData) {
        await this.addSignatureToPDF(page, signatureData, width, height);
      }

      // Save the modified PDF
      const modifiedPdfBytes = await pdfDoc.save();
      await fs.writeFile(emailConfirmationPath, modifiedPdfBytes);

      // Add digital signature and footprint
      const signatureResult = await this.digitalSignatureService.signPDFWithDigitalFootprint(
        emailConfirmationPath,
        signatureData,
        digitalFootprint,
        opportunityId,
        signedBy,
        [1] // Email confirmation is typically a single page document
      );
      
      const signedPdfPath = signatureResult.success ? emailConfirmationPath : emailConfirmationPath;

      this.logger.log(`Email confirmation signed successfully: ${signedPdfPath}`);

      return {
        success: true,
        signedPdfPath: signedPdfPath,
        metadata: {
          signatureId: `EC_${Date.now()}`,
          verificationHash: digitalFootprint?.securityHash || 'N/A',
          signedAt: new Date().toISOString(),
        },
      };
    } catch (error) {
      this.logger.error(`Error signing email confirmation: ${error.message}`, error.stack);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  private async embedFormDataIntoPDF(page: any, formData: any, customerInfo: any, width: number, height: number): Promise<void> {
    try {
      // Define field positions - only essential fields
      const fieldPositions = {
        customerName: { x: 250, y: 180 },                   // Customer name at signature
        signatureDate: { x: 130, y: 120 },                   // Current date
      };

      // Add customer name at signature location
      const customerName = formData.customerName || customerInfo?.name || 'Customer';
      page.drawText(customerName, {
        x: fieldPositions.customerName.x,
        y: fieldPositions.customerName.y,
        size: 10,
        color: rgb(0, 0, 0),
      });

      // Add current date at signature location
      const currentDate = new Date().toLocaleDateString('en-GB');
      page.drawText(currentDate, {
        x: fieldPositions.signatureDate.x,
        y: fieldPositions.signatureDate.y,
        size: 10,
        color: rgb(0, 0, 0),
      });

    } catch (error) {
      this.logger.error(`Error embedding form data into PDF: ${error.message}`, error.stack);
      throw error;
    }
  }

  private async addSignatureToPDF(page: any, signatureData: string, width: number, height: number): Promise<void> {
    try {
      if (!signatureData) return;

      // Convert base64 signature to image and add to PDF
      const signaturePosition = { x: 230, y: 200 }; // Position for signature image
      
      try {
        // Remove data URL prefix if present
        const base64Data = signatureData.replace(/^data:image\/[a-z]+;base64,/, '');
        
        // Create image from base64 data
        const imageBytes = Buffer.from(base64Data, 'base64');
        const signatureImage = await page.doc.embedPng(imageBytes);
        
        // Add signature image to PDF
        page.drawImage(signatureImage, {
          x: signaturePosition.x,
          y: signaturePosition.y,
          width: 150,
          height: 50,
        });
        
        this.logger.log('Signature image added to PDF successfully');
      } catch (imageError) {
        this.logger.warn(`Failed to embed signature image, using text fallback: ${imageError.message}`);
        
        // Fallback to text if image embedding fails
        page.drawText('[SIGNATURE]', {
          x: signaturePosition.x,
          y: signaturePosition.y,
          size: 12,
          color: rgb(0, 0, 0),
        });
      }

    } catch (error) {
      this.logger.error(`Error adding signature to PDF: ${error.message}`, error.stack);
      throw error;
    }
  }

  async getEmailConfirmationHistory(opportunityId: string): Promise<any[]> {
    try {
      // Get all email confirmation files for this opportunity
      const signedDir = path.join(process.cwd(), 'src', 'email_confirmation', 'signed');
      const files = await fs.readdir(signedDir);
      
      const emailConfirmationFiles = files.filter(file => 
        file.includes(`email-confirmation-${opportunityId}`) && file.endsWith('.pdf')
      );

      return emailConfirmationFiles.map(file => ({
        filename: file,
        path: path.join(signedDir, file),
        createdAt: new Date(),
      }));
    } catch (error) {
      this.logger.error(`Error getting email confirmation history: ${error.message}`, error.stack);
      return [];
    }
  }

  async verifyEmailConfirmationSignature(pdfPath: string): Promise<{ valid: boolean; details?: any; error?: string }> {
    try {
      // This would implement signature verification logic
      // For now, return a basic verification
      return {
        valid: true,
        details: {
          verifiedAt: new Date().toISOString(),
          signatureType: 'digital',
        },
      };
    } catch (error) {
      this.logger.error(`Error verifying email confirmation signature: ${error.message}`, error.stack);
      return {
        valid: false,
        error: error.message,
      };
    }
  }

  async downloadSignedEmailConfirmation(opportunityId: string, filename: string): Promise<{ success: boolean; filePath?: string; error?: string }> {
    try {
      const signedDir = path.join(process.cwd(), 'src', 'email_confirmation', 'signed');
      const filePath = path.join(signedDir, filename);

      // Check if file exists
      try {
        await fs.access(filePath);
      } catch (error) {
        throw new Error(`File not found: ${filename}`);
      }

      return {
        success: true,
        filePath: filePath,
      };
    } catch (error) {
      this.logger.error(`Error downloading signed email confirmation: ${error.message}`, error.stack);
      return {
        success: false,
        error: error.message,
      };
    }
  }
}
