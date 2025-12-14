import { Controller, Post, Get, Body, Param, Res, Logger } from '@nestjs/common';
import { Response } from 'express';
import { PdfSignatureService } from './pdf-signature.service';
import * as path from 'path';
import * as fs from 'fs/promises';

@Controller('pdf-signing')
export class PdfSignatureController {
  private readonly logger = new Logger(PdfSignatureController.name);

  constructor(private readonly pdfSignatureService: PdfSignatureService) {}

  /**
   * Create a signed PDF with embedded signature
   */
  @Post('create-signed-pdf')
  async createSignedPDF(@Body() body: {
    originalPdfData: string; // Base64 PDF data
    signatureData: string;
    digitalFootprint: any;
    opportunityId: string;
    customerName: string;
  }) {
    try {
      this.logger.log(`Creating signed PDF for opportunity: ${body.opportunityId}`);

      // Create output path
      const timestamp = Date.now();
      const safeCustomerName = body.customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      const outputFilename = `signed_${safeCustomerName}_${body.opportunityId}_${timestamp}.pdf`;
      const outputPath = path.join(process.cwd(), 'src', 'signatures', 'signed-pdfs', outputFilename);

      // Ensure directory exists
      await fs.mkdir(path.dirname(outputPath), { recursive: true });

      // Create signed PDF
      const result = await this.pdfSignatureService.createSignedPDF(
        body.originalPdfData,
        body.signatureData,
        body.digitalFootprint,
        outputPath
      );

      if (result.success) {
        return {
          success: true,
          message: 'Signed PDF created successfully',
          data: {
            signedPdfPath: result.signedPdfPath,
            filename: outputFilename,
            downloadUrl: `/pdf-signing/download/${outputFilename}`,
          }
        };
      } else {
        return {
          success: false,
          error: result.error,
        };
      }
    } catch (error) {
      this.logger.error(`Error creating signed PDF: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Download signed PDF
   */
  @Get('download/:filename')
  async downloadSignedPDF(
    @Param('filename') filename: string,
    @Res() res: Response
  ) {
    try {
      this.logger.log(`Downloading signed PDF: ${filename}`);

      const filePath = path.join(process.cwd(), 'src', 'signatures', 'signed-pdfs', filename);
      
      // Check if file exists
      try {
        await fs.access(filePath);
      } catch {
        return res.status(404).json({
          success: false,
          error: 'Signed PDF not found',
        });
      }

      // Set headers for PDF download
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      
      // Send the file
      const fileBuffer = await fs.readFile(filePath);
      res.send(fileBuffer);
    } catch (error) {
      this.logger.error(`Error downloading signed PDF: ${error.message}`);
      res.status(500).json({
        success: false,
        error: error.message,
      });
    }
  }

  /**
   * Extract digital footprint from signed PDF
   */
  @Get('extract-footprint/:filename')
  async extractDigitalFootprint(@Param('filename') filename: string) {
    try {
      this.logger.log(`Extracting digital footprint from: ${filename}`);

      const filePath = path.join(process.cwd(), 'src', 'signatures', 'signed-pdfs', filename);
      
      // Check if file exists
      try {
        await fs.access(filePath);
      } catch {
        return {
          success: false,
          error: 'Signed PDF not found',
        };
      }

      const result = await this.pdfSignatureService.extractDigitalFootprint(filePath);
      
      return {
        success: true,
        data: result,
      };
    } catch (error) {
      this.logger.error(`Error extracting digital footprint: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Verify PDF signature
   */
  @Post('verify-signature')
  async verifySignature(@Body() body: {
    filename: string;
    expectedHash: string;
  }) {
    try {
      this.logger.log(`Verifying signature for: ${body.filename}`);

      const filePath = path.join(process.cwd(), 'src', 'signatures', 'signed-pdfs', body.filename);
      
      // Check if file exists
      try {
        await fs.access(filePath);
      } catch {
        return {
          success: false,
          error: 'Signed PDF not found',
        };
      }

      const result = await this.pdfSignatureService.verifyPdfSignature(filePath, body.expectedHash);
      
      return {
        success: true,
        data: result,
      };
    } catch (error) {
      this.logger.error(`Error verifying signature: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Health check endpoint
   */
  @Get('health')
  async healthCheck() {
    return {
      success: true,
      message: 'PDF signing service is running',
      timestamp: new Date().toISOString(),
      features: [
        'PDF signature embedding',
        'Digital footprint integration',
        'Verification watermark',
        'Signature verification',
        'Digital footprint extraction',
      ]
    };
  }
}
