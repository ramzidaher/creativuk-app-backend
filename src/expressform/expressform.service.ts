import { Injectable, Logger } from '@nestjs/common';
import * as fs from 'fs/promises';
import * as path from 'path';
import { PDFDocument, PDFPage, rgb } from 'pdf-lib';
import { DigitalSignatureService, DigitalFootprint } from '../pdf-signature/digital-signature.service';

@Injectable()
export class ExpressFormService {
  private readonly logger = new Logger(ExpressFormService.name);
  private readonly templatePath = path.join(process.cwd(), 'src', 'expressform', 'Express Consent.pdf');
  private readonly signedFolderPath = path.join(process.cwd(), 'src', 'expressform', 'signed');

  constructor(private readonly digitalSignatureService: DigitalSignatureService) {}

  /**
   * Create a copy of the express consent template for signing
   */
  async createExpressConsentCopy(opportunityId: string, customerName: string): Promise<{
    success: boolean;
    expressConsentPath?: string;
    error?: string;
  }> {
    try {
      this.logger.log(`Creating express consent copy for opportunity: ${opportunityId}`);

      // Ensure signed folder exists
      await this.ensureSignedFolderExists();

      // Create filename with opportunity ID and customer name
      const sanitizedCustomerName = customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const filename = `Express_Consent_${opportunityId}_${sanitizedCustomerName}_${timestamp}.pdf`;
      const expressConsentPath = path.join(this.signedFolderPath, filename);

      // Check if template exists
      try {
        await fs.access(this.templatePath);
      } catch (error) {
        throw new Error(`Express consent template not found at: ${this.templatePath}`);
      }

      // Copy template to signed folder
      await fs.copyFile(this.templatePath, expressConsentPath);

      this.logger.log(`Express consent copy created successfully: ${expressConsentPath}`);

      return {
        success: true,
        expressConsentPath,
      };
    } catch (error) {
      this.logger.error(`Error creating express consent copy: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Sign the express consent PDF with digital signature and footprint
   */
  async signExpressConsent(
    expressConsentPath: string,
    signatureData: string,
    digitalFootprint: DigitalFootprint,
    opportunityId: string,
    signedBy: string,
    customerInfo?: {
      name?: string;
      email?: string;
      phone?: string;
      address?: string;
    }
  ): Promise<{
    success: boolean;
    message: string;
    signedExpressConsentPath?: string;
    metadata?: any;
    error?: string;
  }> {
    try {
      this.logger.log(`Signing express consent: ${expressConsentPath}`);

      // Validate express consent file exists
      try {
        await fs.access(expressConsentPath);
      } catch (error) {
        throw new Error(`Express consent file not found: ${expressConsentPath}`);
      }

      // Embed customer info into the PDF before signing
      if (customerInfo) {
        this.logger.log(`Embedding customer info into express consent PDF: ${expressConsentPath}`);
        await this.embedCustomerInfoIntoPDF(expressConsentPath, customerInfo);
        this.logger.log(`Customer info embedded successfully into express consent PDF`);
      }

      // Use the provided digital footprint as-is
      const enhancedDigitalFootprint: DigitalFootprint = {
        ...digitalFootprint,
      };

      // Sign the PDF using the digital signature service
      // For express consent, we'll sign on page 1
      const result = await this.digitalSignatureService.signPDFWithDigitalFootprint(
        expressConsentPath,
        signatureData,
        enhancedDigitalFootprint,
        opportunityId,
        signedBy,
        [1] // Sign on page 1
      );

      if (result.success) {
        this.logger.log(`Express consent signed successfully: ${expressConsentPath}`);
        return {
          success: true,
          message: 'Express consent signed successfully with digital signature and footprint',
          signedExpressConsentPath: expressConsentPath,
          metadata: result.metadata,
        };
      } else {
        throw new Error(result.error || 'Failed to sign express consent');
      }
    } catch (error) {
      this.logger.error(`Error signing express consent: ${error.message}`);
      return {
        success: false,
        message: 'Failed to sign express consent',
        error: error.message,
      };
    }
  }

  /**
   * Get express consent signing history for an opportunity
   */
  async getExpressConsentHistory(opportunityId: string): Promise<{
    success: boolean;
    expressConsents: any[];
    error?: string;
  }> {
    try {
      this.logger.log(`Getting express consent history for opportunity: ${opportunityId}`);

      // List all signed express consents for this opportunity
      const files = await fs.readdir(this.signedFolderPath);
      const expressConsentFiles = files.filter(file => 
        file.includes(opportunityId) && file.endsWith('.pdf')
      );

      const expressConsents = await Promise.all(
        expressConsentFiles.map(async (file) => {
          const filePath = path.join(this.signedFolderPath, file);
          const stats = await fs.stat(filePath);
          
          return {
            filename: file,
            path: filePath,
            signedAt: stats.mtime,
            size: stats.size,
            opportunityId,
          };
        })
      );

      // Sort by signing date (newest first)
      expressConsents.sort((a, b) => b.signedAt.getTime() - a.signedAt.getTime());

      return {
        success: true,
        expressConsents,
      };
    } catch (error) {
      this.logger.error(`Error getting express consent history: ${error.message}`);
      return {
        success: false,
        expressConsents: [],
        error: error.message,
      };
    }
  }

  /**
   * Verify a signed express consent
   */
  async verifyExpressConsent(expressConsentPath: string): Promise<{
    success: boolean;
    isValid: boolean;
    metadata?: any;
    error?: string;
  }> {
    try {
      this.logger.log(`Verifying express consent: ${expressConsentPath}`);

      // Use the digital signature service to verify
      const result = await this.digitalSignatureService.verifySignature(expressConsentPath, '');

      return {
        success: result.success,
        isValid: result.isValid,
        metadata: result.metadata,
        error: result.error,
      };
    } catch (error) {
      this.logger.error(`Error verifying express consent: ${error.message}`);
      return {
        success: false,
        isValid: false,
        error: error.message,
      };
    }
  }

  /**
   * Ensure the signed folder exists
   */
  private async ensureSignedFolderExists(): Promise<void> {
    try {
      await fs.access(this.signedFolderPath);
    } catch (error) {
      this.logger.log(`Creating signed folder: ${this.signedFolderPath}`);
      await fs.mkdir(this.signedFolderPath, { recursive: true });
    }
  }

  /**
   * Get express consent template info
   */
  async getExpressConsentTemplateInfo(): Promise<{
    success: boolean;
    templatePath?: string;
    exists?: boolean;
    size?: number;
    error?: string;
  }> {
    try {
      const stats = await fs.stat(this.templatePath);
      return {
        success: true,
        templatePath: this.templatePath,
        exists: true,
        size: stats.size,
      };
    } catch (error) {
      return {
        success: false,
        templatePath: this.templatePath,
        exists: false,
        error: error.message,
      };
    }
  }

  /**
   * Embed customer info into the express consent PDF
   */
  private async embedCustomerInfoIntoPDF(
    pdfPath: string,
    customerInfo: {
      name?: string;
      email?: string;
      phone?: string;
      address?: string;
    }
  ): Promise<void> {
    try {
      this.logger.log(`Embedding customer info into PDF: ${pdfPath}`);

      // Read the PDF file
      const pdfBytes = await fs.readFile(pdfPath);
      const pdfDoc = await PDFDocument.load(pdfBytes);

      // Get the first page
      const page = pdfDoc.getPage(0);
      const { width, height } = page.getSize();

      // Based on the Express Consent PDF structure, we need to fill in:
      // - Name(s) field
      // - Address field
      // - Date field
      // Signature will be handled by DocuSeal

      // Approximate positions (these may need adjustment based on actual PDF layout)
      const fieldPositions = {
        name: { x: 150, y: height - 300 },      // Name field position
        address: { x: 150, y: height - 350 },   // Address field position
        date: { x: 150, y: height - 400 },      // Date field position
      };

      // Add customer name
      if (customerInfo.name) {
        page.drawText(customerInfo.name, {
          x: fieldPositions.name.x,
          y: fieldPositions.name.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }

      // Add customer address
      if (customerInfo.address) {
        page.drawText(customerInfo.address, {
          x: fieldPositions.address.x,
          y: fieldPositions.address.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }

      // Add current date
      const currentDate = new Date().toLocaleDateString('en-GB');
      page.drawText(currentDate, {
        x: fieldPositions.date.x,
        y: fieldPositions.date.y,
        size: 10,
        color: rgb(0, 0, 0),
      });

      // Save the modified PDF
      const modifiedPdfBytes = await pdfDoc.save();
      await fs.writeFile(pdfPath, modifiedPdfBytes);

      this.logger.log(`Customer info embedded successfully into PDF: ${pdfPath}`);
    } catch (error) {
      this.logger.error(`Error embedding customer info into PDF: ${error.message}`);
      throw error;
    }
  }
}






