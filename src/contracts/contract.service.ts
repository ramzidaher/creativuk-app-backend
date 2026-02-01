import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { DocuSealService } from '../integrations/docuseal.service';
import { PdfGeneratorService } from './pdf-generator.service';
import { EPVSAutomationService } from '../epvs-automation/epvs-automation.service';
import { ExcelAutomationService } from '../excel-automation/excel-automation.service';
import * as path from 'path';
import * as fs from 'fs';

export interface ContractData {
  opportunityId: string;
  customerName: string;
  customerEmail: string;
  date: string;
  postcode: string;
  contractType: 'solar_installation' | 'maintenance' | 'warranty';
  calculatorType?: 'flux' | 'off-peak' | 'epvs';
  solarData?: any;
}

export interface ContractSigningResult {
  templateId: string;
  submissionId: string;
  signingUrl: string;
  status: 'pending' | 'completed' | 'failed';
}

@Injectable()
export class ContractService {
  private readonly logger = new Logger(ContractService.name);
  private readonly contractTemplatesDir = path.join(process.cwd(), 'dist', 'contracts', 'templates');
  private readonly outputDir = path.join(process.cwd(), 'dist', 'contracts', 'output');

  constructor(
    private readonly docuSealService: DocuSealService,
    private readonly pdfGeneratorService: PdfGeneratorService,
    @Inject(forwardRef(() => EPVSAutomationService))
    private readonly epvsAutomationService: EPVSAutomationService,
    @Inject(forwardRef(() => ExcelAutomationService))
    private readonly excelAutomationService: ExcelAutomationService
  ) {
    // Ensure directories exist
    if (!fs.existsSync(this.contractTemplatesDir)) {
      fs.mkdirSync(this.contractTemplatesDir, { recursive: true });
    }
    if (!fs.existsSync(this.outputDir)) {
      fs.mkdirSync(this.outputDir, { recursive: true });
    }
  }

  /**
   * Create a contract signing workflow for an opportunity
   */
  async createContractSigningWorkflow(contractData: ContractData): Promise<ContractSigningResult> {
    try {
      this.logger.log(`Creating contract signing workflow for opportunity: ${contractData.opportunityId}`);

      // 1. Generate contract PDF using the appropriate service based on calculator type
      let contractPdfPath: string;
      
      if (contractData.calculatorType === 'epvs') {
        // Use EPVS automation service to generate PDF
        const pdfResult = await this.epvsAutomationService.generatePDF(
          contractData.opportunityId,
          undefined, // excelFilePath - will be determined by the service
          undefined  // signatureData - not needed for contract generation
        );
        
        if (!pdfResult.success || !pdfResult.pdfPath) {
          throw new Error(`Failed to generate PDF from EPVS: ${pdfResult.error || 'Unknown error'}`);
        }
        
        contractPdfPath = pdfResult.pdfPath;
      } else {
        // Use Excel automation service for flux/off-peak calculators
        const pdfResult = await this.excelAutomationService.generatePDF(
          contractData.opportunityId,
          undefined, // excelFilePath - will be determined by the service
          undefined  // signatureData - not needed for contract generation
        );
        
        if (!pdfResult.success || !pdfResult.pdfPath) {
          throw new Error(`Failed to generate PDF from Excel automation: ${pdfResult.error || 'Unknown error'}`);
        }
        
        contractPdfPath = pdfResult.pdfPath;
      }

      // 2. Create DocuSeal template and submission - CONTRACT ONLY (booking confirmation is separate)
      this.logger.log(`Creating contract-only signing workflow (booking confirmation will be separate)`);
      const signingWorkflow = await this.docuSealService.createContractSigningWorkflowWithBase64(
        contractPdfPath,
        contractData.opportunityId,
        {
          name: contractData.customerName,
          email: contractData.customerEmail,
        },
        {
          customerName: contractData.customerName,
          date: contractData.date,
          postcode: contractData.postcode,
          ...contractData.solarData,
        }
      );

      this.logger.log(`Contract signing workflow created successfully for opportunity: ${contractData.opportunityId}`);

      return {
        templateId: signingWorkflow.templateId,
        submissionId: signingWorkflow.submissionId,
        signingUrl: signingWorkflow.signingUrl,
        status: 'pending',
      };
    } catch (error) {
      this.logger.error(`Failed to create contract signing workflow: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get contract signing status
   */
  async getContractSigningStatus(submissionId: string): Promise<{
    status: string;
    completed: boolean;
    signedDocumentUrl?: string;
  }> {
    try {
      const submission = await this.docuSealService.getSubmissionStatus(submissionId);
      
      const isCompleted = submission.status === 'completed';
      let signedDocumentUrl: string | undefined;

      if (isCompleted) {
        // Get the signed document
        const signedDocument = await this.docuSealService.getSignedDocument(submissionId);
        
        // Save the signed document
        const signedDocPath = path.join(this.outputDir, `signed_contract_${submissionId}.pdf`);
        fs.writeFileSync(signedDocPath, signedDocument);
        
        signedDocumentUrl = `/contracts/download/${submissionId}`;
      }

      return {
        status: submission.status,
        completed: isCompleted,
        signedDocumentUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to get contract signing status: ${error.message}`);
      throw error;
    }
  }

  /**
   * Download signed contract
   */
  async downloadSignedContract(submissionId: string): Promise<Buffer> {
    try {
      return await this.docuSealService.getSignedDocument(submissionId);
    } catch (error) {
      this.logger.error(`Failed to download signed contract: ${error.message}`);
      throw error;
    }
  }

}
