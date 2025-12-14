import { Controller, Post, Get, Body, Res, HttpException, HttpStatus } from '@nestjs/common';
import { Response } from 'express';
import { PdfSigningService } from './pdf-signing.service';

interface SignPdfRequest {
  pdfBase64: string;
  signerInfo?: {
    name?: string;
    reason?: string;
    location?: string;
    contactInfo?: string;
  };
}

interface CreateSignedDocumentRequest {
  templatePath: string;
  outputPath: string;
  signerInfo?: {
    name?: string;
    reason?: string;
    location?: string;
    contactInfo?: string;
  };
}

@Controller('pdf-signing')
export class PdfSigningController {
  constructor(private readonly pdfSigningService: PdfSigningService) {}

  /**
   * Sign a PDF document
   */
  @Post('sign')
  async signPdf(@Body() body: SignPdfRequest, @Res() res: Response) {
    try {
      const { pdfBase64, signerInfo } = body;
      
      if (!pdfBase64) {
        throw new HttpException('PDF base64 data is required', HttpStatus.BAD_REQUEST);
      }

      // Convert base64 to buffer
      const pdfBuffer = Buffer.from(pdfBase64, 'base64');
      
      // Sign the PDF
      const signedPdfBuffer = await this.pdfSigningService.signPdf(pdfBuffer, signerInfo);
      
      // Convert back to base64
      const signedPdfBase64 = signedPdfBuffer.toString('base64');
      
      res.json({
        success: true,
        message: 'PDF signed successfully',
        data: {
          signedPdfBase64,
          signatureInfo: {
            timestamp: new Date().toISOString(),
            signer: signerInfo?.name || 'Creativ Solar',
            reason: signerInfo?.reason || 'Document approval',
            location: signerInfo?.location || 'San Francisco, CA'
          }
        }
      });
      
    } catch (error) {
      throw new HttpException(
        `Failed to sign PDF: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a signed document from a template
   */
  @Post('create-signed-document')
  async createSignedDocument(@Body() body: CreateSignedDocumentRequest, @Res() res: Response) {
    try {
      const { templatePath, outputPath, signerInfo } = body;
      
      if (!templatePath || !outputPath) {
        throw new HttpException('Template path and output path are required', HttpStatus.BAD_REQUEST);
      }

      // Create the signed document
      const resultPath = await this.pdfSigningService.createSignedDocument(
        templatePath,
        outputPath,
        signerInfo
      );
      
      res.json({
        success: true,
        message: 'Signed document created successfully',
        data: {
          outputPath: resultPath,
          signatureInfo: {
            timestamp: new Date().toISOString(),
            signer: signerInfo?.name || 'Creativ Solar',
            reason: signerInfo?.reason || 'Document approval',
            location: signerInfo?.location || 'San Francisco, CA'
          }
        }
      });
      
    } catch (error) {
      throw new HttpException(
        `Failed to create signed document: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Verify a PDF signature
   */
  @Post('verify')
  async verifyPdfSignature(@Body() body: { pdfBase64: string }, @Res() res: Response) {
    try {
      const { pdfBase64 } = body;
      
      if (!pdfBase64) {
        throw new HttpException('PDF base64 data is required', HttpStatus.BAD_REQUEST);
      }

      // Convert base64 to buffer
      const pdfBuffer = Buffer.from(pdfBase64, 'base64');
      
      // Verify the signature
      const verificationResult = await this.pdfSigningService.verifyPdfSignature(pdfBuffer);
      
      res.json({
        success: true,
        message: 'PDF signature verification completed',
        data: verificationResult
      });
      
    } catch (error) {
      throw new HttpException(
        `Failed to verify PDF signature: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Generate a certificate for download
   */
  @Get('generate-certificate')
  async generateCertificate(@Res() res: Response) {
    try {
      const certificate = await this.pdfSigningService.generateCertificateForDownload();
      
      res.json({
        success: true,
        message: 'Certificate generated successfully',
        data: certificate
      });
      
    } catch (error) {
      throw new HttpException(
        `Failed to generate certificate: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Test PDF signing
   */
  @Get('test')
  async testPdfSigning(@Res() res: Response) {
    try {
      // Create a simple test PDF
      const { PDFDocument, rgb } = await import('pdf-lib');
      const pdfDoc = await PDFDocument.create();
      const page = pdfDoc.addPage([600, 400]);
      
      page.drawText('Test Document for Digital Signing', {
        x: 50,
        y: 350,
        size: 20,
        color: rgb(0, 0, 0),
      });
      
      page.drawText('This is a test document to verify PDF signing functionality.', {
        x: 50,
        y: 300,
        size: 12,
        color: rgb(0, 0, 0),
      });
      
      const pdfBytes = await pdfDoc.save();
      const pdfBuffer = Buffer.from(pdfBytes);
      
      // Sign the test PDF
      const signedPdfBuffer = await this.pdfSigningService.signPdf(pdfBuffer, {
        name: 'Test Signer',
        reason: 'Testing PDF signing functionality',
        location: 'Test Environment',
        contactInfo: 'test@creativsolar.com'
      });
      
      // Convert to base64 for response
      const signedPdfBase64 = signedPdfBuffer.toString('base64');
      
      res.json({
        success: true,
        message: 'Test PDF signing completed successfully',
        data: {
          signedPdfBase64,
          signatureInfo: {
            timestamp: new Date().toISOString(),
            signer: 'Test Signer',
            reason: 'Testing PDF signing functionality',
            location: 'Test Environment'
          }
        }
      });
      
    } catch (error) {
      throw new HttpException(
        `Test PDF signing failed: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}
