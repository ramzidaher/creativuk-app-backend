import { Controller, Post, Get, Body, Param, HttpException, HttpStatus, Res } from '@nestjs/common';
import { Response } from 'express';
import { DisclaimerService } from './disclaimer.service';
import { DigitalFootprint } from '../pdf-signature/digital-signature.service';
import * as fs from 'fs';

@Controller('disclaimer')
export class DisclaimerController {
  constructor(private readonly disclaimerService: DisclaimerService) {}

  @Post('create-copy')
  async createDisclaimerCopy(@Body() body: {
    opportunityId: string;
    customerName: string;
  }): Promise<{
    success: boolean;
    disclaimerPath?: string;
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

      const result = await this.disclaimerService.createDisclaimerCopy(opportunityId, customerName);

      if (result.success) {
        return {
          success: true,
          disclaimerPath: result.disclaimerPath,
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.error || 'Failed to create disclaimer copy',
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
          message: 'Internal server error during disclaimer copy creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('sign')
  async signDisclaimer(@Body() body: {
    disclaimerPath: string;
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
    };
  }): Promise<{
    success: boolean;
    message: string;
    signedDisclaimerPath?: string;
    metadata?: any;
    error?: string;
  }> {
    try {
      const { disclaimerPath, signatureData, digitalFootprint, opportunityId, signedBy, customerInfo, formData } = body;
      
      if (!disclaimerPath || !signatureData || !digitalFootprint || !opportunityId || !signedBy) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: disclaimerPath, signatureData, digitalFootprint, opportunityId, and signedBy are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Validate disclaimer path exists
      if (!fs.existsSync(disclaimerPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Disclaimer file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      const result = await this.disclaimerService.signDisclaimer(
        disclaimerPath,
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
          message: result.message,
          signedDisclaimerPath: result.signedDisclaimerPath,
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
          message: 'Internal server error during disclaimer signing',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('history/:opportunityId')
  async getDisclaimerHistory(@Param('opportunityId') opportunityId: string): Promise<{
    success: boolean;
    disclaimers: any[];
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

      const result = await this.disclaimerService.getDisclaimerHistory(opportunityId);

      return {
        success: result.success,
        disclaimers: result.disclaimers,
        error: result.error,
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during disclaimer history retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('verify')
  async verifyDisclaimer(@Body() body: {
    disclaimerPath: string;
  }): Promise<{
    success: boolean;
    isValid: boolean;
    metadata?: any;
    error?: string;
  }> {
    try {
      const { disclaimerPath } = body;
      
      if (!disclaimerPath) {
        throw new HttpException(
          {
            success: false,
            message: 'Disclaimer path is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Validate disclaimer path exists
      if (!fs.existsSync(disclaimerPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Disclaimer file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      const result = await this.disclaimerService.verifyDisclaimer(disclaimerPath);

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
          message: 'Internal server error during disclaimer verification',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('template-info')
  async getDisclaimerTemplateInfo(): Promise<{
    success: boolean;
    templatePath?: string;
    exists?: boolean;
    size?: number;
    error?: string;
  }> {
    try {
      const result = await this.disclaimerService.getDisclaimerTemplateInfo();

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
  async downloadDisclaimer(@Param('filename') filename: string, @Res() res: Response): Promise<void> {
    try {
      const disclaimerPath = `C:\\Users\\\Creativuk\\creativ-solar-app\\apps\\backend\\src\\disclaimer\\signed\\${filename}`;
      
      // Check if file exists
      if (!fs.existsSync(disclaimerPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Disclaimer file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Set headers for PDF download
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
      
      // Stream the PDF file
      const fileStream = fs.createReadStream(disclaimerPath);
      fileStream.pipe(res);
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during disclaimer download',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}
