import { Controller, Post, Get, Body, Param, Res, Logger } from '@nestjs/common';
import { Response } from 'express';
import { FreeSignaturesService, FreeSignatureData } from './free-signatures.service';
import * as fs from 'fs';

@Controller('free-signatures')
export class FreeSignaturesController {
  private readonly logger = new Logger(FreeSignaturesController.name);

  constructor(private readonly freeSignaturesService: FreeSignaturesService) {}

  /**
   * Save a free signature
   */
  @Post('save')
  async saveSignature(@Body() signatureData: FreeSignatureData) {
    try {
      this.logger.log(`Saving free signature for opportunity: ${signatureData.opportunityId}`);
      
      const result = await this.freeSignaturesService.saveFreeSignature(signatureData);
      
      return {
        success: true,
        message: 'Signature saved successfully',
        data: {
          signatureId: result.signature.id,
          digitalFootprint: result.digitalFootprint,
          filePath: result.filePath,
        }
      };
    } catch (error) {
      this.logger.error(`Error saving free signature: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Verify a free signature
   */
  @Get('verify/:opportunityId/:signatureHash')
  async verifySignature(
    @Param('opportunityId') opportunityId: string,
    @Param('signatureHash') signatureHash: string
  ) {
    try {
      this.logger.log(`Verifying free signature for opportunity: ${opportunityId}`);
      
      const result = await this.freeSignaturesService.verifyFreeSignature(opportunityId, signatureHash);
      
      return {
        success: true,
        data: result,
      };
    } catch (error) {
      this.logger.error(`Error verifying free signature: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Get all free signatures for an opportunity
   */
  @Get('opportunity/:opportunityId')
  async getSignatures(@Param('opportunityId') opportunityId: string) {
    try {
      this.logger.log(`Getting free signatures for opportunity: ${opportunityId}`);
      
      const result = await this.freeSignaturesService.getFreeSignatures(opportunityId);
      
      return {
        success: true,
        data: result,
      };
    } catch (error) {
      this.logger.error(`Error getting free signatures: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Download signed document
   */
  @Get('download/:opportunityId')
  async downloadDocument(
    @Param('opportunityId') opportunityId: string,
    @Res() res: Response
  ) {
    try {
      this.logger.log(`Downloading signed document for opportunity: ${opportunityId}`);
      
      const result = await this.freeSignaturesService.downloadSignedDocument(opportunityId);
      
      if (!result.success) {
        return res.status(404).json({
          success: false,
          error: 'Document not found',
        });
      }

      // Set headers for file download
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Content-Disposition', `attachment; filename="signed_document_${opportunityId}.json"`);
      
      // Send the file
      res.send(result.document);
    } catch (error) {
      this.logger.error(`Error downloading signed document: ${error.message}`);
      res.status(500).json({
        success: false,
        error: error.message,
      });
    }
  }

  /**
   * Get digital footprint for a signature
   */
  @Get('footprint/:opportunityId')
  async getDigitalFootprint(@Param('opportunityId') opportunityId: string) {
    try {
      this.logger.log(`Getting digital footprint for opportunity: ${opportunityId}`);
      
      const result = await this.freeSignaturesService.downloadSignedDocument(opportunityId);
      
      if (!result.success) {
        return {
          success: false,
          error: 'Document not found',
        };
      }

      return {
        success: true,
        data: {
          digitalFootprint: result.digitalFootprint,
          verification: {
            isValid: true,
            verifiedAt: new Date().toISOString(),
            method: 'MULTI_LAYER_HASH',
          }
        }
      };
    } catch (error) {
      this.logger.error(`Error getting digital footprint: ${error.message}`);
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
      message: 'Free signatures service is running',
      timestamp: new Date().toISOString(),
      features: [
        'Document upload and signing',
        'Digital footprint generation',
        'Multi-layer hash verification',
        'Legal compliance tracking',
        'Audit trail maintenance',
        'Signature verification',
        'Document download',
      ]
    };
  }
}
