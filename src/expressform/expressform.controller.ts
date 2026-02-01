import { Controller, Post, Get, Body, Param, HttpException, HttpStatus, Res } from '@nestjs/common';
import { Response } from 'express';
import { ExpressFormService } from './expressform.service';
import { DigitalFootprint } from '../pdf-signature/digital-signature.service';
import * as fs from 'fs';
import * as path from 'path';

@Controller('expressform')
export class ExpressFormController {
  constructor(private readonly expressFormService: ExpressFormService) {}

  @Post('create-copy')
  async createExpressConsentCopy(@Body() body: {
    opportunityId: string;
    customerName: string;
  }): Promise<{
    success: boolean;
    expressConsentPath?: string;
    error?: string;
  }> {
    try {
      const { opportunityId, customerName } = body;
      
      if (!opportunityId || !customerName) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId and customerName are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.expressFormService.createExpressConsentCopy(opportunityId, customerName);

      if (result.success) {
        return {
          success: true,
          expressConsentPath: result.expressConsentPath,
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.error || 'Failed to create express consent copy',
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
          message: 'Internal server error during express consent copy creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('sign')
  async signExpressConsent(@Body() body: {
    expressConsentPath: string;
    signatureData: string;
    digitalFootprint: DigitalFootprint;
    opportunityId: string;
    signedBy: string;
    customerInfo?: {
      name?: string;
      email?: string;
      phone?: string;
      address?: string;
    };
  }): Promise<{
    success: boolean;
    message: string;
    signedExpressConsentPath?: string;
    metadata?: any;
    error?: string;
  }> {
    try {
      const { expressConsentPath, signatureData, digitalFootprint, opportunityId, signedBy, customerInfo } = body;
      
      if (!expressConsentPath || !signatureData || !digitalFootprint || !opportunityId || !signedBy) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: expressConsentPath, signatureData, digitalFootprint, opportunityId, and signedBy are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Validate express consent path exists
      if (!fs.existsSync(expressConsentPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Express consent file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      const result = await this.expressFormService.signExpressConsent(
        expressConsentPath,
        signatureData,
        digitalFootprint,
        opportunityId,
        signedBy,
        customerInfo
      );

      if (result.success) {
        return {
          success: true,
          message: result.message,
          signedExpressConsentPath: result.signedExpressConsentPath,
          metadata: result.metadata,
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
          message: 'Internal server error during express consent signing',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('history/:opportunityId')
  async getExpressConsentHistory(@Param('opportunityId') opportunityId: string): Promise<{
    success: boolean;
    expressConsents: any[];
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

      const result = await this.expressFormService.getExpressConsentHistory(opportunityId);

      return {
        success: result.success,
        expressConsents: result.expressConsents,
        error: result.error,
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during express consent history retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('verify')
  async verifyExpressConsent(@Body() body: {
    expressConsentPath: string;
  }): Promise<{
    success: boolean;
    isValid: boolean;
    metadata?: any;
    error?: string;
  }> {
    try {
      const { expressConsentPath } = body;
      
      if (!expressConsentPath) {
        throw new HttpException(
          {
            success: false,
            message: 'Express consent path is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Validate express consent path exists
      if (!fs.existsSync(expressConsentPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Express consent file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      const result = await this.expressFormService.verifyExpressConsent(expressConsentPath);

      return {
        success: result.success,
        isValid: result.isValid,
        metadata: result.metadata,
        error: result.error,
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during express consent verification',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('template-info')
  async getExpressConsentTemplateInfo(): Promise<{
    success: boolean;
    templatePath?: string;
    exists?: boolean;
    size?: number;
    error?: string;
  }> {
    try {
      const result = await this.expressFormService.getExpressConsentTemplateInfo();

      return {
        success: result.success,
        templatePath: result.templatePath,
        exists: result.exists,
        size: result.size,
        error: result.error,
      };
    } catch (error) {
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during template info retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('download/:filename')
  async downloadExpressConsent(@Param('filename') filename: string, @Res() res: Response): Promise<void> {
    try {
      const expressConsentPath = path.join(process.cwd(), 'src', 'expressform', 'signed', filename);
      
      // Check if file exists
      if (!fs.existsSync(expressConsentPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Express consent file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Set headers for PDF download
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
      
      // Stream the PDF file
      const fileStream = fs.createReadStream(expressConsentPath);
      fileStream.pipe(res);
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during express consent download',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}

