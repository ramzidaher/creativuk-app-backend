import { Controller, Post, Get, Body, Param, HttpException, HttpStatus, Res } from '@nestjs/common';
import { Response } from 'express';
import { EmailConfirmationService } from './email-confirmation.service';
import { DigitalFootprint } from '../pdf-signature/digital-signature.service';
import * as fs from 'fs';

@Controller('email-confirmation')
export class EmailConfirmationController {
  constructor(private readonly emailConfirmationService: EmailConfirmationService) {}

  @Post('create-copy')
  async createEmailConfirmationCopy(@Body() body: {
    opportunityId: string;
    customerName: string;
  }): Promise<{
    success: boolean;
    emailConfirmationPath?: string;
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

      const result = await this.emailConfirmationService.createEmailConfirmationCopy(opportunityId, customerName);

      if (result.success) {
        return {
          success: true,
          emailConfirmationPath: result.emailConfirmationPath,
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.error || 'Failed to create email confirmation copy',
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
          message: 'Internal server error during email confirmation copy creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('sign')
  async signEmailConfirmation(@Body() body: {
    emailConfirmationPath?: string;
    signatureData: string;
    digitalFootprint: DigitalFootprint;
    opportunityId: string;
    signedBy: string;
    customerName?: string;
    customerInfo?: {
      name?: string;
      email?: string;
      phone?: string;
      address?: string;
    };
    formData?: {
      customerName?: string;
      installationDate?: string;
      systemSize?: string;
      totalCost?: string;
      depositAmount?: string;
      balanceAmount?: string;
      paymentTerms?: string;
    };
  }): Promise<{
    success: boolean;
    message: string;
    signedEmailConfirmationPath?: string;
    metadata?: any;
    error?: string;
  }> {
    try {
      const { emailConfirmationPath, signatureData, digitalFootprint, opportunityId, signedBy, customerName, customerInfo, formData } = body;
      
      if (!signatureData || !digitalFootprint || !opportunityId || !signedBy) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: signatureData, digitalFootprint, opportunityId, and signedBy are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      let finalEmailConfirmationPath: string;

      // If no emailConfirmationPath provided or file doesn't exist, create a new copy
      if (!emailConfirmationPath || !fs.existsSync(emailConfirmationPath)) {
        const customerNameForCopy = customerName || customerInfo?.name || formData?.customerName || 'Customer';
        const copyResult = await this.emailConfirmationService.createEmailConfirmationCopy(opportunityId, customerNameForCopy);
        
        if (!copyResult.success || !copyResult.emailConfirmationPath) {
          throw new HttpException(
            {
              success: false,
              message: `Failed to create email confirmation copy: ${copyResult.error || 'No path returned'}`,
            },
            HttpStatus.INTERNAL_SERVER_ERROR
          );
        }
        
        finalEmailConfirmationPath = copyResult.emailConfirmationPath;
      } else {
        finalEmailConfirmationPath = emailConfirmationPath;
      }

      const result = await this.emailConfirmationService.signEmailConfirmation(
        finalEmailConfirmationPath,
        signatureData,
        digitalFootprint,
        opportunityId,
        signedBy,
        customerInfo,
        formData
      );

      if (result.success) {
        return {
          success: true,
          message: 'Email confirmation signed successfully',
          signedEmailConfirmationPath: result.signedPdfPath,
          metadata: result.metadata,
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.error || 'Failed to sign email confirmation',
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
          message: 'Internal server error during email confirmation signing',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('history/:opportunityId')
  async getEmailConfirmationHistory(@Param('opportunityId') opportunityId: string): Promise<{
    success: boolean;
    emailConfirmations: any[];
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

      const result = await this.emailConfirmationService.getEmailConfirmationHistory(opportunityId);

      return {
        success: true,
        emailConfirmations: result,
        error: undefined,
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during email confirmation history retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('verify')
  async verifyEmailConfirmation(@Body() body: {
    emailConfirmationPath: string;
  }): Promise<{
    success: boolean;
    isValid: boolean;
    metadata?: any;
    error?: string;
  }> {
    try {
      const { emailConfirmationPath } = body;
      
      if (!emailConfirmationPath) {
        throw new HttpException(
          {
            success: false,
            message: 'Email confirmation path is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Validate email confirmation path exists
      if (!fs.existsSync(emailConfirmationPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Email confirmation file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      const result = await this.emailConfirmationService.verifyEmailConfirmationSignature(emailConfirmationPath);

      return {
        success: result.valid,
        isValid: result.valid,
        metadata: result.details,
        error: result.error,
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during email confirmation verification',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('template-info')
  async getEmailConfirmationTemplateInfo(): Promise<{
    success: boolean;
    templatePath?: string;
    exists?: boolean;
    size?: number;
    error?: string;
  }> {
    try {
      // Simple template info check - use same path construction as service
      const templatePath = require('path').join(process.cwd(), 'src', 'email_confirmation', 'Confirmation of Booking Letter.pdf');
      const exists = require('fs').existsSync(templatePath);
      const stats = exists ? require('fs').statSync(templatePath) : null;
      
      const result = {
        success: true,
        templatePath: templatePath,
        exists: exists,
        size: stats ? stats.size : 0,
        error: undefined,
      };

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
  async downloadEmailConfirmation(@Param('filename') filename: string, @Res() res: Response): Promise<void> {
    try {
      const emailConfirmationPath = require('path').join(process.cwd(), 'src', 'email_confirmation', 'signed', filename);
      
      // Check if file exists
      if (!fs.existsSync(emailConfirmationPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Email confirmation file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Set headers for PDF download
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
      
      // Stream the PDF file
      const fileStream = fs.createReadStream(emailConfirmationPath);
      fileStream.pipe(res);
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during email confirmation download',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}



