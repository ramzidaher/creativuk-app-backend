import { Controller, Post, Body, UseGuards, Request } from '@nestjs/common';
import { OneDriveFileManagerService } from './onedrive-file-manager.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { Logger } from '@nestjs/common';

@Controller('onedrive')
@UseGuards(JwtAuthGuard)
export class OneDriveFileManagerController {
  private readonly logger = new Logger(OneDriveFileManagerController.name);

  constructor(
    private readonly oneDriveFileManagerService: OneDriveFileManagerService
  ) {}

  @Post('copy-proposal')
  async copyProposalToOneDrive(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
      proposalFiles: { pptxPath?: string; pdfPath?: string };
    }
  ) {
    try {
      this.logger.log(`Copying proposal to OneDrive for opportunity: ${body.opportunityId}`);
      
      const result = await this.oneDriveFileManagerService.copyProposalToQuotationsWithSurveyImages(
        body.opportunityId,
        body.customerName,
        body.proposalFiles
      );

      return {
        success: result.success,
        message: result.message,
        error: result.error,
        folderPath: result.folderPath
      };
    } catch (error) {
      this.logger.error(`Error copying proposal to OneDrive: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy proposal to OneDrive',
        error: error.message
      };
    }
  }

  @Post('copy-contract')
  async copyContractToOneDrive(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
      contractFiles: { pptxPath?: string; pdfPath?: string };
    }
  ) {
    try {
      this.logger.log(`Copying contract to OneDrive for opportunity: ${body.opportunityId}`);
      
      const result = await this.oneDriveFileManagerService.copyContractToOrdersWithSurveyImages(
        body.opportunityId,
        body.customerName,
        body.contractFiles
      );

      return {
        success: result.success,
        message: result.message,
        error: result.error,
        folderPath: result.folderPath
      };
    } catch (error) {
      this.logger.error(`Error copying contract to OneDrive: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy contract to OneDrive',
        error: error.message
      };
    }
  }

  @Post('copy-files')
  async copyFilesToOneDrive(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
      proposalFiles: { pptxPath?: string; pdfPath?: string };
      contractFiles?: { pptxPath?: string; pdfPath?: string };
      contractSubmissionId?: string;
    }
  ) {
    try {
      this.logger.log(`Copying files to OneDrive for opportunity: ${body.opportunityId}`);
      
      const result = await this.oneDriveFileManagerService.copyFilesToOneDrive(
        body.opportunityId,
        body.customerName,
        body.proposalFiles,
        body.contractFiles
      );

      return result;
    } catch (error) {
      this.logger.error(`Error copying files to OneDrive: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy files to OneDrive',
        error: error.message,
        results: {
          quotations: { success: false, message: 'Failed to copy proposal files' }
        }
      };
    }
  }

  @Post('verify-paths')
  async verifyOneDrivePaths(@Request() req) {
    try {
      const result = await this.oneDriveFileManagerService.verifyOneDrivePaths();
      return {
        success: true,
        message: 'OneDrive paths verified',
        paths: result
      };
    } catch (error) {
      this.logger.error(`Error verifying OneDrive paths: ${error.message}`);
      return {
        success: false,
        message: 'Failed to verify OneDrive paths',
        error: error.message
      };
    }
  }

  @Post('get-stats')
  async getOneDriveStats(@Request() req) {
    try {
      const stats = await this.oneDriveFileManagerService.getOneDriveStats();
      return {
        success: true,
        message: 'OneDrive statistics retrieved',
        stats
      };
    } catch (error) {
      this.logger.error(`Error getting OneDrive stats: ${error.message}`);
      return {
        success: false,
        message: 'Failed to get OneDrive statistics',
        error: error.message
      };
    }
  }

  @Post('copy-disclaimer')
  async copyDisclaimerToOneDrive(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
      disclaimerPath: string;
    }
  ) {
    try {
      this.logger.log(`Copying disclaimer to OneDrive for opportunity: ${body.opportunityId}`);
      
      const result = await this.oneDriveFileManagerService.copyDisclaimerToOrders(
        body.opportunityId,
        body.customerName,
        body.disclaimerPath
      );

      return {
        success: result.success,
        message: result.message,
        error: result.error,
        folderPath: result.folderPath
      };
    } catch (error) {
      this.logger.error(`Error copying disclaimer to OneDrive: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy disclaimer to OneDrive',
        error: error.message
      };
    }
  }

  @Post('copy-email-confirmation')
  async copyEmailConfirmationToOneDrive(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
      emailConfirmationPath: string;
    }
  ) {
    try {
      this.logger.log(`Copying email confirmation to OneDrive for opportunity: ${body.opportunityId}`);
      
      const result = await this.oneDriveFileManagerService.copyEmailConfirmationToOrders(
        body.opportunityId,
        body.customerName,
        body.emailConfirmationPath
      );

      return {
        success: result.success,
        message: result.message,
        error: result.error,
        folderPath: result.folderPath
      };
    } catch (error) {
      this.logger.error(`Error copying email confirmation to OneDrive: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy email confirmation to OneDrive',
        error: error.message
      };
    }
  }

  @Post('copy-won-opportunity-documents')
  async copyWonOpportunityDocumentsToOneDrive(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
      documents: {
        proposalFiles?: { pptxPath?: string; pdfPath?: string };
        contractFiles?: { pptxPath?: string; pdfPath?: string };
        disclaimerPath?: string;
        emailConfirmationPath?: string;
        contractSubmissionId?: string;
        bookingConfirmationSubmissionId?: string;
      };
    }
  ) {
    try {
      this.logger.log(`Copying won opportunity documents to OneDrive for opportunity: ${body.opportunityId}`);
      
      const result = await this.oneDriveFileManagerService.copyWonOpportunityDocumentsToOneDrive(
        body.opportunityId,
        body.customerName,
        body.documents
      );

      return result;
    } catch (error) {
      this.logger.error(`Error copying won opportunity documents to OneDrive: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy won opportunity documents to OneDrive',
        error: error.message,
        results: {}
      };
    }
  }

  @Post('copy-documents-by-opportunity-id')
  async copyDocumentsByOpportunityId(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
    }
  ) {
    try {
      this.logger.log(`Copying documents by opportunity ID to OneDrive for opportunity: ${body.opportunityId}`);
      
      const result = await this.oneDriveFileManagerService.copyWonOpportunityDocumentsToOneDrive(
        body.opportunityId,
        body.customerName,
        {} // Empty documents object - will trigger automatic file finding by opportunity ID
      );

      return result;
    } catch (error) {
      this.logger.error(`Error copying documents by opportunity ID to OneDrive: ${error.message}`);
      return {
        success: false,
        message: 'Failed to copy documents by opportunity ID to OneDrive',
        error: error.message,
        results: {}
      };
    }
  }

  @Post('organize-by-outcome')
  async organizeFilesByOutcome(
    @Request() req,
    @Body() body: {
      opportunityId: string;
      customerName: string;
      postcode: string;
      outcome: 'won' | 'lost';
      files: {
        surveyPath?: string;
        calculatorPath?: string;
        contractPath?: string;
        proposalPath?: string;
        disclaimerPath?: string;
        emailConfirmationPath?: string;
      };
    }
  ) {
    try {
      this.logger.log(`Organizing files by outcome (${body.outcome}) for opportunity: ${body.opportunityId}`);
      
      const userId = req.user?.id || req.user?.userId || undefined;
      
      const result = await this.oneDriveFileManagerService.organizeFilesByOutcome(
        body.opportunityId,
        body.customerName,
        body.postcode,
        body.outcome,
        body.files,
        userId
      );

      return {
        success: result.success,
        message: result.message,
        error: result.error,
        folderPath: result.folderPath
      };
    } catch (error) {
      this.logger.error(`Error organizing files by outcome: ${error.message}`);
      return {
        success: false,
        message: 'Failed to organize files by outcome',
        error: error.message
      };
    }
  }
}

