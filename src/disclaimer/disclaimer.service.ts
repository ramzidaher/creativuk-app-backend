import { Injectable, Logger } from '@nestjs/common';
import * as fs from 'fs/promises';
import * as path from 'path';
import { PDFDocument, PDFPage, rgb } from 'pdf-lib';
import { DigitalSignatureService, DigitalFootprint } from '../pdf-signature/digital-signature.service';

@Injectable()
export class DisclaimerService {
  private readonly logger = new Logger(DisclaimerService.name);
  private readonly templatePath = path.join(process.cwd(), 'src', 'disclaimer', 'EPVS_Disclaimer_Template.pdf');
  private readonly signedFolderPath = path.join(process.cwd(), 'src', 'disclaimer', 'signed');

  constructor(private readonly digitalSignatureService: DigitalSignatureService) {}

  /**
   * Create a copy of the disclaimer template for signing
   */
  async createDisclaimerCopy(opportunityId: string, customerName: string): Promise<{
    success: boolean;
    disclaimerPath?: string;
    error?: string;
  }> {
    try {
      this.logger.log(`Creating disclaimer copy for opportunity: ${opportunityId}`);

      // Ensure signed folder exists
      await this.ensureSignedFolderExists();

      // Create filename with opportunity ID and customer name
      const sanitizedCustomerName = customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const filename = `EPVS_Disclaimer_${opportunityId}_${sanitizedCustomerName}_${timestamp}.pdf`;
      const disclaimerPath = path.join(this.signedFolderPath, filename);

      // Check if template exists
      try {
        await fs.access(this.templatePath);
      } catch (error) {
        throw new Error(`Disclaimer template not found at: ${this.templatePath}`);
      }

      // Copy template to signed folder
      await fs.copyFile(this.templatePath, disclaimerPath);

      this.logger.log(`Disclaimer copy created successfully: ${disclaimerPath}`);

      return {
        success: true,
        disclaimerPath,
      };
    } catch (error) {
      this.logger.error(`Error creating disclaimer copy: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Sign the disclaimer PDF with digital signature and footprint
   */
  async signDisclaimer(
    disclaimerPath: string,
    signatureData: string,
    digitalFootprint: DigitalFootprint,
    opportunityId: string,
    signedBy: string,
    customerInfo?: {
      name?: string;
      email?: string;
      phone?: string;
      address?: string;
    },
    formData?: {
      installerName?: string;
      customerName?: string;
      unitRate?: string;
      unitRateReason?: string;
      gridConsumptionKnown?: boolean;
      annualGridConsumption?: string;
      annualElectricitySpend?: string;
      standingCharge?: string;
      gridConsumptionReason?: string;
      utilityBillReason?: string;
    }
  ): Promise<{
    success: boolean;
    message: string;
    signedDisclaimerPath?: string;
    metadata?: any;
    error?: string;
  }> {
    try {
      this.logger.log(`Signing disclaimer: ${disclaimerPath}`);

      // Validate disclaimer file exists
      try {
        await fs.access(disclaimerPath);
      } catch (error) {
        throw new Error(`Disclaimer file not found: ${disclaimerPath}`);
      }

      // Embed form data into the PDF before signing
      if (formData) {
        this.logger.log(`Embedding form data into disclaimer PDF: ${disclaimerPath}`);
        await this.embedFormDataIntoPDF(disclaimerPath, formData, customerInfo);
        this.logger.log(`Form data embedded successfully into disclaimer PDF`);
      }

      // Use the provided digital footprint as-is
      const enhancedDigitalFootprint: DigitalFootprint = {
        ...digitalFootprint,
      };

      // Sign the PDF using the digital signature service
      // For disclaimer, we'll sign on the last page (typically page 1 for single page disclaimers)
      const result = await this.digitalSignatureService.signPDFWithDigitalFootprint(
        disclaimerPath,
        signatureData,
        enhancedDigitalFootprint,
        opportunityId,
        signedBy,
        [1] // Sign on page 1 (adjust based on your disclaimer template)
      );

      if (result.success) {
        this.logger.log(`Disclaimer signed successfully: ${disclaimerPath}`);
        return {
          success: true,
          message: 'Disclaimer signed successfully with digital signature and footprint',
          signedDisclaimerPath: disclaimerPath,
          metadata: result.metadata,
        };
      } else {
        throw new Error(result.error || 'Failed to sign disclaimer');
      }
    } catch (error) {
      this.logger.error(`Error signing disclaimer: ${error.message}`);
      return {
        success: false,
        message: 'Failed to sign disclaimer',
        error: error.message,
      };
    }
  }

  /**
   * Get disclaimer signing history for an opportunity
   */
  async getDisclaimerHistory(opportunityId: string): Promise<{
    success: boolean;
    disclaimers: any[];
    error?: string;
  }> {
    try {
      this.logger.log(`Getting disclaimer history for opportunity: ${opportunityId}`);

      // List all signed disclaimers for this opportunity
      const files = await fs.readdir(this.signedFolderPath);
      const disclaimerFiles = files.filter(file => 
        file.includes(opportunityId) && file.endsWith('.pdf')
      );

      const disclaimers = await Promise.all(
        disclaimerFiles.map(async (file) => {
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
      disclaimers.sort((a, b) => b.signedAt.getTime() - a.signedAt.getTime());

      return {
        success: true,
        disclaimers,
      };
    } catch (error) {
      this.logger.error(`Error getting disclaimer history: ${error.message}`);
      return {
        success: false,
        disclaimers: [],
        error: error.message,
      };
    }
  }

  /**
   * Verify a signed disclaimer
   */
  async verifyDisclaimer(disclaimerPath: string): Promise<{
    success: boolean;
    isValid: boolean;
    metadata?: any;
    error?: string;
  }> {
    try {
      this.logger.log(`Verifying disclaimer: ${disclaimerPath}`);

      // Use the digital signature service to verify
      const result = await this.digitalSignatureService.verifySignature(disclaimerPath, '');

      return {
        success: result.success,
        isValid: result.isValid,
        metadata: result.metadata,
        error: result.error,
      };
    } catch (error) {
      this.logger.error(`Error verifying disclaimer: ${error.message}`);
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
   * Get disclaimer template info
   */
  async getDisclaimerTemplateInfo(): Promise<{
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
   * Embed form data into the disclaimer PDF
   */
  private async embedFormDataIntoPDF(
    pdfPath: string,
    formData: {
      installerName?: string;
      customerName?: string;
      unitRate?: string;
      unitRateReason?: string;
      gridConsumptionKnown?: boolean;
      annualGridConsumption?: string;
      annualElectricitySpend?: string;
      standingCharge?: string;
      gridConsumptionReason?: string;
      utilityBillReason?: string;
    },
    customerInfo?: {
      name?: string;
      email?: string;
      phone?: string;
      address?: string;
    }
  ): Promise<void> {
    try {
      this.logger.log(`Embedding form data into PDF: ${pdfPath}`);

      // Read the PDF file
      const pdfBytes = await fs.readFile(pdfPath);
      const pdfDoc = await PDFDocument.load(pdfBytes);

      // Get the first page (disclaimer is typically single page)
      const page = pdfDoc.getPage(0);
      const { width, height } = page.getSize();

      // Define positions for form fields (matching the disclaimer template)
      const fieldPositions = {
        installerName: { x: 150, y: height - 144 },           // After "Installer Name:"
        customerName: { x: 400, y: height - 144 },           // After "Customer Name:"
        unitRate: { x: 400, y: height - 213 },               // In the unit rate section
        annualGridConsumption: { x: 160, y: height - 355 },  // In grid consumption section
        annualElectricitySpend: { x: 180, y: height - 407 }, // In spend section
        standingCharge: { x: 410, y: height - 407 },         // In spend section
        utilityBillReason: { x: 70, y: height - 573 },       // Reason for not having utility bill
        signatureName: { x: 150, y: height - 690 },          // Name field at bottom
        signatureDate: { x: 150, y: height - 720 },          // Date field at bottom
        signatureLine: { x: 550, y: height - 40 }           // Signature line at bottom
      };

      // Add form data to the PDF (just the values, no labels)
      if (formData.installerName) {
        page.drawText(formData.installerName, {
          x: fieldPositions.installerName.x,
          y: fieldPositions.installerName.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }

      if (formData.customerName) {
        page.drawText(formData.customerName, {
          x: fieldPositions.customerName.x,
          y: fieldPositions.customerName.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }

      if (formData.unitRate) {
        page.drawText(formData.unitRate, {
          x: fieldPositions.unitRate.x,
          y: fieldPositions.unitRate.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }


      if (formData.gridConsumptionKnown && formData.annualGridConsumption) {
        page.drawText(formData.annualGridConsumption, {
          x: fieldPositions.annualGridConsumption.x,
          y: fieldPositions.annualGridConsumption.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      } else if (!formData.gridConsumptionKnown && formData.annualElectricitySpend) {
        page.drawText(formData.annualElectricitySpend, {
          x: fieldPositions.annualElectricitySpend.x,
          y: fieldPositions.annualElectricitySpend.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }

      if (!formData.gridConsumptionKnown && formData.standingCharge) {
        page.drawText(formData.standingCharge, {
          x: fieldPositions.standingCharge.x,
          y: fieldPositions.standingCharge.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }


      // Add utility bill reason (this field is missing from our form data)
      if (formData.utilityBillReason) {
        page.drawText(formData.utilityBillReason, {
          x: fieldPositions.utilityBillReason.x,
          y: fieldPositions.utilityBillReason.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }

      // Add signature fields at the bottom
      if (formData.customerName) {
        page.drawText(formData.customerName, {
          x: fieldPositions.signatureName.x,
          y: fieldPositions.signatureName.y,
          size: 10,
          color: rgb(0, 0, 0),
        });
      }

      // Add current date
      const currentDate = new Date().toLocaleDateString('en-GB');
      page.drawText(currentDate, {
        x: fieldPositions.signatureDate.x,
        y: fieldPositions.signatureDate.y,
        size: 10,
        color: rgb(0, 0, 0),
      });

      // Add customer info if available
      if (customerInfo) {
        let yPos = height - 400;
        if (customerInfo.email) {
          page.drawText(`Email: ${customerInfo.email}`, {
            x: 100,
            y: yPos,
            size: 10,
            color: rgb(0, 0, 0),
          });
          yPos -= 20;
        }
        if (customerInfo.phone) {
          page.drawText(`Phone: ${customerInfo.phone}`, {
            x: 100,
            y: yPos,
            size: 10,
            color: rgb(0, 0, 0),
          });
          yPos -= 20;
        }
        if (customerInfo.address) {
          page.drawText(`Address: ${customerInfo.address}`, {
            x: 100,
            y: yPos,
            size: 10,
            color: rgb(0, 0, 0),
          });
        }
      }

      // Save the modified PDF
      const modifiedPdfBytes = await pdfDoc.save();
      await fs.writeFile(pdfPath, modifiedPdfBytes);

      this.logger.log(`Form data embedded successfully into PDF: ${pdfPath}`);
    } catch (error) {
      this.logger.error(`Error embedding form data into PDF: ${error.message}`);
      throw error;
    }
  }
}
