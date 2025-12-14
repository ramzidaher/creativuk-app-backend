import { Controller, Post, Get, Param, Body, Res, Logger } from '@nestjs/common';
import { Response } from 'express';
import { ContractService, ContractData } from './contract.service';

@Controller('contracts')
export class ContractController {
  private readonly logger = new Logger(ContractController.name);

  constructor(private readonly contractService: ContractService) {}

  /**
   * Create a new contract signing workflow
   */
  @Post('create-signing-workflow')
  async createSigningWorkflow(@Body() contractData: ContractData) {
    try {
      this.logger.log(`Creating signing workflow for opportunity: ${contractData.opportunityId}`);
      
      const result = await this.contractService.createContractSigningWorkflow(contractData);
      
      return {
        success: true,
        data: result,
        message: 'Contract signing workflow created successfully',
      };
    } catch (error) {
      this.logger.error(`Failed to create signing workflow: ${error.message}`);
      return {
        success: false,
        error: error.message,
        message: 'Failed to create contract signing workflow',
      };
    }
  }

  /**
   * Get contract signing status
   */
  @Get('status/:submissionId')
  async getSigningStatus(@Param('submissionId') submissionId: string) {
    try {
      const status = await this.contractService.getContractSigningStatus(submissionId);
      
      return {
        success: true,
        data: status,
        message: 'Contract status retrieved successfully',
      };
    } catch (error) {
      this.logger.error(`Failed to get signing status: ${error.message}`);
      return {
        success: false,
        error: error.message,
        message: 'Failed to get contract signing status',
      };
    }
  }

  /**
   * Download signed contract
   */
  @Get('download/:submissionId')
  async downloadSignedContract(
    @Param('submissionId') submissionId: string,
    @Res() res: Response
  ) {
    try {
      const signedDocument = await this.contractService.downloadSignedContract(submissionId);
      
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="signed_contract_${submissionId}.pdf"`);
      res.send(signedDocument);
    } catch (error) {
      this.logger.error(`Failed to download signed contract: ${error.message}`);
      res.status(500).json({
        success: false,
        error: error.message,
        message: 'Failed to download signed contract',
      });
    }
  }

  /**
   * Get signing URL for a specific submission
   */
  @Get('signing-url/:submissionId')
  async getSigningUrl(@Param('submissionId') submissionId: string) {
    try {
      // This would typically get the signing URL from your database
      // For now, we'll return a placeholder
      return {
        success: true,
        data: {
          signingUrl: `${process.env.DOCUSEAL_BASE_URL || 'http://localhost:3001'}/s/${submissionId}`,
          message: 'Use this URL to access the signing interface',
        },
        message: 'Signing URL retrieved successfully',
      };
    } catch (error) {
      this.logger.error(`Failed to get signing URL: ${error.message}`);
      return {
        success: false,
        error: error.message,
        message: 'Failed to get signing URL',
      };
    }
  }
}
