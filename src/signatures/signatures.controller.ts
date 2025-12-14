import { Controller, Post, Get, Body, Param, HttpException, HttpStatus } from '@nestjs/common';
import { SignaturesService } from './signatures.service';

@Controller('signatures')
export class SignaturesController {
  constructor(private readonly signaturesService: SignaturesService) {}

  @Post('save')
  async saveSignature(@Body() body: {
    opportunityId: string;
    signatureData: string;
    signedAt: string;
    signedBy?: string;
    ipAddress?: string;
    userAgent?: string;
  }) {
    try {
      const { opportunityId, signatureData, signedAt, signedBy, ipAddress, userAgent } = body;

      if (!opportunityId || !signatureData || !signedAt) {
        throw new HttpException(
          {
            success: false,
            message: 'Missing required fields: opportunityId, signatureData, signedAt',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const signature = await this.signaturesService.saveSignature(opportunityId, signatureData, {
        signedAt,
        signedBy,
        ipAddress,
        userAgent
      });

      return {
        success: true,
        message: 'Signature saved successfully',
        data: {
          id: signature.id,
          opportunityId: signature.opportunityId,
          signedAt: signature.signedAt
        }
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Error saving signature',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get(':opportunityId')
  async getSignature(@Param('opportunityId') opportunityId: string) {
    try {
      const signature = await this.signaturesService.getSignature(opportunityId);

      if (!signature) {
        return {
          success: true,
          data: null,
          message: 'No signature found for this opportunity'
        };
      }

      return {
        success: true,
        data: {
          id: signature.id,
          opportunityId: signature.opportunityId,
          signedAt: signature.signedAt,
          signedBy: signature.signedBy
        }
      };
    } catch (error) {
      throw new HttpException(
        {
          success: false,
          message: 'Error retrieving signature',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get(':opportunityId/all')
  async getSignaturesByOpportunity(@Param('opportunityId') opportunityId: string) {
    try {
      const signatures = await this.signaturesService.getSignaturesByOpportunity(opportunityId);

      return {
        success: true,
        data: signatures.map(sig => ({
          id: sig.id,
          opportunityId: sig.opportunityId,
          signedAt: sig.signedAt,
          signedBy: sig.signedBy
        }))
      };
    } catch (error) {
      throw new HttpException(
        {
          success: false,
          message: 'Error retrieving signatures',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}
