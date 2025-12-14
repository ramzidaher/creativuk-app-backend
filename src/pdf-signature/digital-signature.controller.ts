import { Controller, Post, Get, Body, Param, HttpException, HttpStatus, Res } from '@nestjs/common';
import { Response } from 'express';
import { DigitalSignatureService, DigitalFootprint } from './digital-signature.service';
import * as path from 'path';

@Controller('digital-signature')
export class DigitalSignatureController {
  constructor(private readonly digitalSignatureService: DigitalSignatureService) {}

  @Post('sign-pdf')
  async signPDF(@Body() body: {
    pdfPath: string;
    signatureData: string;
    digitalFootprint: DigitalFootprint;
    opportunityId: string;
    signedBy: string;
    pageNumbers?: number[];
  }): Promise<{
    success: boolean;
    message: string;
    metadata?: any;
    error?: string;
  }> {
    try {
      const { pdfPath, signatureData, digitalFootprint, opportunityId, signedBy, pageNumbers } = body;
      
      if (!pdfPath || !signatureData || !digitalFootprint || !opportunityId || !signedBy) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: pdfPath, signatureData, digitalFootprint, opportunityId, and signedBy are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Validate PDF path exists
      const fullPdfPath = path.resolve(pdfPath);
      const isValidCalculatorPath = fullPdfPath.includes('epvs-opportunities') || 
                                   fullPdfPath.includes('excel-file-calculator') ||
                                   fullPdfPath.includes('opportunities');
      
      if (!isValidCalculatorPath || !fullPdfPath.endsWith('.pdf')) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid PDF path: must be a valid calculator PDF (EPVS or Off Peak)',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.digitalSignatureService.signPDFWithDigitalFootprint(
        fullPdfPath,
        signatureData,
        digitalFootprint,
        opportunityId,
        signedBy,
        pageNumbers
      );

      if (result.success) {
        return {
          success: true,
          message: result.message,
          metadata: result.metadata
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during PDF signing',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('verify/:signatureId')
  async verifySignature(@Param('signatureId') signatureId: string): Promise<{
    success: boolean;
    isValid: boolean;
    metadata?: any;
    error?: string;
  }> {
    try {
      if (!signatureId) {
        throw new HttpException(
          {
            success: false,
            message: 'Signature ID is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // For verification, we need to find the PDF path from metadata
      // This is a simplified version - in production you'd want to store this in a database
      const result = await this.digitalSignatureService.verifySignature('', signatureId);

      return {
        success: result.success,
        isValid: result.isValid,
        metadata: result.metadata,
        error: result.error
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during signature verification',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('history/:opportunityId')
  async getSignatureHistory(@Param('opportunityId') opportunityId: string): Promise<{
    success: boolean;
    signatures: any[];
    error?: string;
  }> {
    try {
      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            message: 'Opportunity ID is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const signatures = await this.digitalSignatureService.getSignatureHistory(opportunityId);

      return {
        success: true,
        signatures
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during signature history retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('test-pdf/:opportunityId')
  async getTestPDF(@Param('opportunityId') opportunityId: string, @Res() res: Response): Promise<void> {
    try {
      // Use the specific PDF path provided by the user
      const pdfPath = `C:\\Users\\\Creativuk\\creativ-solar-app\\apps\\backend\\src\\excel-file-calculator\\epvs-opportunities\\pdfs\\EPVS Calculator - 47hmE2SisQlAC8Ppd5O3.pdf`;
      
      // Check if file exists
      const fs = require('fs');
      if (!fs.existsSync(pdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Test PDF not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Set headers for PDF download
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `inline; filename="EPVS Calculator - ${opportunityId}.pdf"`);
      
      // Stream the PDF file
      const fileStream = fs.createReadStream(pdfPath);
      fileStream.pipe(res);
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during PDF retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}
