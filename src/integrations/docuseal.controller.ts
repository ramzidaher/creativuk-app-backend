import { Controller, Post, Get, Body, Param, Res, HttpException, HttpStatus, Logger, Inject, forwardRef, Headers, HttpCode } from '@nestjs/common';
import { Response } from 'express';
import { DocuSealService } from './docuseal.service';
import * as fs from 'fs';
import * as path from 'path';
import { OpportunitiesService } from '../opportunities/opportunities.service';

@Controller('docuseal')
export class DocuSealController {
  private readonly logger = new Logger(DocuSealController.name);

  constructor(
    private readonly docuSealService: DocuSealService,
    @Inject(forwardRef(() => OpportunitiesService))
    private readonly opportunitiesService: OpportunitiesService
  ) {}

  /**
   * Create a signing workflow for Flux Contract
   * POST /docuseal/contract/flux
   */
  @Post('contract/flux')
  async createFluxContractSigning(
    @Body() body: {
      contractPdfPath?: string;
      opportunityId: string;
      customerData: {
        name: string;
        email: string;
      };
      contractData: {
        customerName: string;
        date: string;
        postcode: string;
        [key: string]: any;
      };
    }
  ) {
    try {
      const { contractPdfPath, opportunityId, customerData, contractData } = body;

      if (!opportunityId || !customerData || !contractData) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId, customerData, and contractData are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Creating Flux contract signing workflow for opportunity: ${opportunityId}`);

      // Always prioritize auto-detection for Flux PDFs to ensure correct calculator type
      // Try multiple possible Flux/EPVS PDF naming patterns
      const possibleFluxPaths = [
        path.join(
          process.cwd(),
          'src',
          'excel-file-calculator',
          'epvs-opportunities',
          'pdfs',
          `EPVS Calculator - ${opportunityId}.pdf`
        ),
        path.join(
          process.cwd(),
          'src',
          'excel-file-calculator',
          'epvs-opportunities',
          'pdfs',
          `EPVS Calculator Creativ - 06.02 - ${opportunityId}.pdf`
        ),
      ];

      let finalContractPdfPath: string | undefined;
      
      // First, try to find Flux PDF in default location
      for (const possiblePath of possibleFluxPaths) {
        if (fs.existsSync(possiblePath)) {
          finalContractPdfPath = possiblePath;
          this.logger.log(`✅ Auto-detected Flux PDF path: ${finalContractPdfPath}`);
          break;
        }
      }

      // If auto-detection failed, use provided path as fallback
      if (!finalContractPdfPath) {
        if (contractPdfPath && fs.existsSync(contractPdfPath)) {
          finalContractPdfPath = contractPdfPath;
          this.logger.warn(`⚠️ Using provided contractPdfPath (auto-detection failed): ${finalContractPdfPath}`);
        } else {
          throw new HttpException(
            {
              success: false,
              message: `Flux contract PDF file not found. Searched in: ${possibleFluxPaths.join(', ')}. Please ensure the PDF exists in the default location.`,
            },
            HttpStatus.NOT_FOUND
          );
        }
      }

      // Validate PDF file exists
      if (!fs.existsSync(finalContractPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Contract PDF file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Get booking confirmation PDF path
      const bookConfirmationPdfPath = path.join(
        process.cwd(),
        'src',
        'email_confirmation',
        'Confirmation of Booking Letter.pdf'
      );

      // Verify booking confirmation exists (required for all contract signings)
      if (!fs.existsSync(bookConfirmationPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Booking confirmation PDF is required but not found at ${bookConfirmationPdfPath}`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Always use combined method with booking confirmation
      this.logger.log(`Using combined Flux contract + booking confirmation signing workflow`);
      const result = await this.docuSealService.createContractAndBookingConfirmationSigningWorkflow(
        finalContractPdfPath,
        bookConfirmationPdfPath,
        opportunityId,
        customerData,
        contractData,
        'flux'
      );

      return {
        success: true,
        message: 'Flux contract signing workflow created successfully. Email sent to customer.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create Flux contract signing workflow: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during Flux contract signing workflow creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a template (without submission) for Flux Contract - for field position verification
   * POST /docuseal/contract/flux/template
   */
  @Post('contract/flux/template')
  async createFluxContractTemplate(
    @Body() body: {
      contractPdfPath?: string;
      opportunityId: string;
      contractData: {
        customerName: string;
        date: string;
        postcode: string;
        [key: string]: any;
      };
    }
  ) {
    try {
      const { contractPdfPath, opportunityId, contractData } = body;

      if (!opportunityId || !contractData) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId and contractData are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Creating Flux contract template (verification only) for opportunity: ${opportunityId}`);

      // Auto-detect Flux PDF path
      const possibleFluxPaths = [
        path.join(
          process.cwd(),
          'src',
          'excel-file-calculator',
          'epvs-opportunities',
          'pdfs',
          `EPVS Calculator - ${opportunityId}.pdf`
        ),
        path.join(
          process.cwd(),
          'src',
          'excel-file-calculator',
          'epvs-opportunities',
          'pdfs',
          `EPVS Calculator Creativ - 06.02 - ${opportunityId}.pdf`
        ),
      ];

      let finalContractPdfPath: string | undefined;
      
      for (const possiblePath of possibleFluxPaths) {
        if (fs.existsSync(possiblePath)) {
          finalContractPdfPath = possiblePath;
          this.logger.log(`✅ Auto-detected Flux PDF path: ${finalContractPdfPath}`);
          break;
        }
      }

      if (!finalContractPdfPath) {
        if (contractPdfPath && fs.existsSync(contractPdfPath)) {
          finalContractPdfPath = contractPdfPath;
          this.logger.warn(`⚠️ Using provided contractPdfPath (auto-detection failed): ${finalContractPdfPath}`);
        } else {
          throw new HttpException(
            {
              success: false,
              message: `Flux contract PDF file not found. Searched in: ${possibleFluxPaths.join(', ')}`,
            },
            HttpStatus.NOT_FOUND
          );
        }
      }

      if (!fs.existsSync(finalContractPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Contract PDF file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Get booking confirmation PDF path
      const bookConfirmationPdfPath = path.join(
        process.cwd(),
        'src',
        'email_confirmation',
        'Confirmation of Booking Letter.pdf'
      );

      if (!fs.existsSync(bookConfirmationPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Booking confirmation PDF is required but not found at ${bookConfirmationPdfPath}`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      const result = await this.docuSealService.createContractAndBookingConfirmationTemplate(
        finalContractPdfPath,
        bookConfirmationPdfPath,
        opportunityId,
        contractData,
        'flux'
      );

      return {
        success: true,
        message: 'Flux contract template created successfully. Use previewUrl to verify field positions.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create Flux contract template: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during Flux contract template creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a template (without submission) for Off Peak Contract - for field position verification
   * POST /docuseal/contract/off-peak/template
   */
  @Post('contract/off-peak/template')
  async createOffPeakContractTemplate(
    @Body() body: {
      contractPdfPath?: string;
      opportunityId: string;
      contractData: {
        customerName: string;
        date: string;
        postcode: string;
        [key: string]: any;
      };
    }
  ) {
    try {
      const { contractPdfPath, opportunityId, contractData } = body;

      if (!opportunityId || !contractData) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId and contractData are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Creating Off Peak contract template (verification only) for opportunity: ${opportunityId}`);

      // Auto-detect Off Peak PDF path
      const defaultOffPeakPath = path.join(
        process.cwd(),
        'src',
        'excel-file-calculator',
        'opportunities',
        'pdfs',
        `Off Peak Calculator - ${opportunityId}.pdf`
      );

      let finalContractPdfPath: string | undefined;
      
      if (fs.existsSync(defaultOffPeakPath)) {
        finalContractPdfPath = defaultOffPeakPath;
        this.logger.log(`✅ Auto-detected Off Peak PDF path: ${finalContractPdfPath}`);
      } else {
        if (contractPdfPath && fs.existsSync(contractPdfPath)) {
          finalContractPdfPath = contractPdfPath;
          this.logger.warn(`⚠️ Using provided contractPdfPath (auto-detection failed): ${finalContractPdfPath}`);
        } else {
          throw new HttpException(
            {
              success: false,
              message: `Off Peak contract PDF file not found. Searched in: ${defaultOffPeakPath}`,
            },
            HttpStatus.NOT_FOUND
          );
        }
      }

      if (!fs.existsSync(finalContractPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Contract PDF file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Get booking confirmation PDF path
      const bookConfirmationPdfPath = path.join(
        process.cwd(),
        'src',
        'email_confirmation',
        'Confirmation of Booking Letter.pdf'
      );

      if (!fs.existsSync(bookConfirmationPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Booking confirmation PDF is required but not found at ${bookConfirmationPdfPath}`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      const result = await this.docuSealService.createContractAndBookingConfirmationTemplate(
        finalContractPdfPath,
        bookConfirmationPdfPath,
        opportunityId,
        contractData,
        'off-peak'
      );

      return {
        success: true,
        message: 'Off Peak contract template created successfully. Use previewUrl to verify field positions.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create Off Peak contract template: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during Off Peak contract template creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a submission from a verified template
   * POST /docuseal/template/:templateId/submit
   */
  @Post('template/:templateId/submit')
  async createSubmissionFromTemplate(
    @Param('templateId') templateId: string,
    @Body() body: {
      opportunityId: string;
      customerData: {
        name: string;
        email: string;
      };
      contractData: {
        customerName: string;
        date: string;
        postcode: string;
        [key: string]: any;
      };
    }
  ) {
    try {
      const { opportunityId, customerData, contractData } = body;

      if (!templateId || !opportunityId || !customerData || !contractData) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: templateId, opportunityId, customerData, and contractData are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Creating submission from verified template ${templateId} for opportunity: ${opportunityId}`);

      const result = await this.docuSealService.createSubmissionFromVerifiedTemplate(
        templateId,
        opportunityId,
        customerData,
        contractData
      );

      return {
        success: true,
        message: 'Submission created from verified template successfully. Email sent to customer.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create submission from template: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during submission creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a signing workflow for Off Peak Contract
   * POST /docuseal/contract/off-peak
   */
  @Post('contract/off-peak')
  async createOffPeakContractSigning(
    @Body() body: {
      contractPdfPath?: string;
      opportunityId: string;
      customerData: {
        name: string;
        email: string;
      };
      contractData: {
        customerName: string;
        date: string;
        postcode: string;
        [key: string]: any;
      };
    }
  ) {
    try {
      const { contractPdfPath, opportunityId, customerData, contractData } = body;

      if (!opportunityId || !customerData || !contractData) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId, customerData, and contractData are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Creating Off Peak contract signing workflow for opportunity: ${opportunityId}`);

      // Always prioritize auto-detection for Off Peak PDFs to ensure correct calculator type
      const defaultOffPeakPath = path.join(
        process.cwd(),
        'src',
        'excel-file-calculator',
        'opportunities',
        'pdfs',
        `Off Peak Calculator - ${opportunityId}.pdf`
      );

      let finalContractPdfPath: string | undefined;
      
      // First, try to find Off Peak PDF in default location
      if (fs.existsSync(defaultOffPeakPath)) {
        finalContractPdfPath = defaultOffPeakPath;
        this.logger.log(`✅ Auto-detected Off Peak PDF path: ${finalContractPdfPath}`);
      } else {
        // If auto-detection failed, use provided path as fallback
        if (contractPdfPath && fs.existsSync(contractPdfPath)) {
          finalContractPdfPath = contractPdfPath;
          this.logger.warn(`⚠️ Using provided contractPdfPath (auto-detection failed): ${finalContractPdfPath}`);
        } else {
          throw new HttpException(
            {
              success: false,
              message: `Off Peak contract PDF file not found. Searched in: ${defaultOffPeakPath}. Please ensure the PDF exists in the default location.`,
            },
            HttpStatus.NOT_FOUND
          );
        }
      }

      // Validate PDF file exists
      if (!fs.existsSync(finalContractPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'Contract PDF file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Get booking confirmation PDF path
      const bookConfirmationPdfPath = path.join(
        process.cwd(),
        'src',
        'email_confirmation',
        'Confirmation of Booking Letter.pdf'
      );

      // Verify booking confirmation exists (required for all contract signings)
      if (!fs.existsSync(bookConfirmationPdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Booking confirmation PDF is required but not found at ${bookConfirmationPdfPath}`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Always use combined method with booking confirmation
      this.logger.log(`Using combined Off Peak contract + booking confirmation signing workflow`);
      const result = await this.docuSealService.createContractAndBookingConfirmationSigningWorkflow(
        finalContractPdfPath,
        bookConfirmationPdfPath,
        opportunityId,
        customerData,
        contractData,
        'off-peak'
      );

      return {
        success: true,
        message: 'Off Peak contract signing workflow created successfully. Email sent to customer.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create Off Peak contract signing workflow: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during Off Peak contract signing workflow creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a signing workflow for Contract only (legacy route - kept for backward compatibility)
   * POST /docuseal/contract
   * @deprecated Use /docuseal/contract/flux or /docuseal/contract/off-peak instead
   */
  @Post('contract')
  async createContractSigning(
    @Body() body: {
      contractPdfPath: string;
      opportunityId: string;
      customerData: {
        name: string;
        email: string;
      };
      contractData: {
        customerName: string;
        date: string;
        postcode: string;
        [key: string]: any;
      };
    }
  ) {
    // Redirect to flux route for backward compatibility
    this.logger.warn(`Using deprecated /docuseal/contract route. Please use /docuseal/contract/flux instead.`);
    return this.createFluxContractSigning(body);
  }

  /**
   * Create a signing workflow for Disclaimer only
   * POST /docuseal/disclaimer
   */
  @Post('disclaimer')
  async createDisclaimerSigning(
    @Body() body: {
      disclaimerPdfPath?: string;
      opportunityId: string;
      customerData?: {
        name: string;
        email: string;
      };
      customerName?: string;
      installerName?: string;
      userId?: string;
      formData?: {
        unitRate?: string;
        annualGridConsumption?: string;
        annualElectricitySpend?: string;
        standingCharge?: string;
        utilityBillReason?: string;
      };
    }
  ) {
    try {
      const { disclaimerPdfPath, opportunityId, customerData, customerName, installerName, userId, formData } = body;

      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Always use the disclaimer template PDF (DocuSeal will handle signing)
      const templatePath = path.join(process.cwd(), 'src', 'disclaimer', 'EPVS_Disclaimer_Template.pdf');
      const pdfPath = disclaimerPdfPath || templatePath;

      if (!fs.existsSync(pdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Disclaimer template PDF not found at: ${templatePath}. Please ensure the file exists.`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      this.logger.log(`Using disclaimer template PDF: ${pdfPath}`);

      // Auto-fetch customer data if userId is provided
      let finalCustomerData: { name: string; email: string } | undefined = customerData;
      let finalCustomerName: string | undefined = customerName;
      
      if (userId && (!customerData || !customerName)) {
        try {
          const opportunity = await this.opportunitiesService.getOpportunityById(opportunityId, userId);
          if (opportunity) {
            const fetchedName = opportunity.contactName || customerName || 'Customer';
            finalCustomerName = fetchedName;
            finalCustomerData = customerData || {
              name: fetchedName,
              email: opportunity.contactEmail || '',
            };
            this.logger.log(`Auto-fetched customer data: ${finalCustomerName}, ${finalCustomerData.email}`);
          }
        } catch (error) {
          this.logger.warn(`Failed to fetch customer data: ${error.message}`);
        }
      }

      if (!finalCustomerData || !finalCustomerName) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: customerData and customerName are required (or provide userId to auto-fetch)',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // TypeScript now knows these are defined after the check above
      const validatedCustomerData = finalCustomerData;
      const validatedCustomerName = finalCustomerName;

      // Call the service method
      const result = await this.docuSealService.createDisclaimerSigningWorkflow(
        pdfPath,
        opportunityId,
        validatedCustomerData,
        validatedCustomerName,
        installerName,
        formData
      );

      return {
        success: true,
        message: 'Disclaimer signing workflow created successfully. Email sent to customer.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create disclaimer signing workflow: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during disclaimer signing workflow creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a signing workflow for Express Consent only
   * POST /docuseal/express-consent
   */
  @Post('express-consent')
  async createExpressConsentSigning(
    @Body() body: {
      expressConsentPdfPath?: string;
      opportunityId: string;
      customerData?: {
        name: string;
        email: string;
        address?: string;
      };
      customerName?: string;
      userId?: string;
    }
  ) {
    try {
      const { expressConsentPdfPath, opportunityId, customerData, customerName, userId } = body;

      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Always use the express consent template PDF (DocuSeal will handle signing)
      const templatePath = path.join(process.cwd(), 'src', 'expressform', 'Express Consent.pdf');
      const pdfPath = expressConsentPdfPath || templatePath;

      if (!fs.existsSync(pdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Express consent template PDF not found at: ${templatePath}. Please ensure the file exists.`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      this.logger.log(`Using express consent template PDF: ${pdfPath}`);

      // Auto-fetch customer data if userId is provided
      let finalCustomerData: { name: string; email: string; address?: string } | undefined = customerData;
      let finalCustomerName: string | undefined = customerName;
      
      if (userId && (!customerData || !customerName)) {
        try {
          const opportunity = await this.opportunitiesService.getOpportunityById(opportunityId, userId);
          if (opportunity) {
            const fetchedName = opportunity.contactName || customerName || 'Customer';
            finalCustomerName = fetchedName;
            
            // Fetch full customer details including address
            let customerAddress: string | undefined;
            try {
              const opportunityDetails = await this.opportunitiesService.getOpportunityDetails(opportunityId, userId);
              if (opportunityDetails.contactAddress) {
                customerAddress = opportunityDetails.contactAddress;
              } else if (opportunityDetails.address) {
                customerAddress = opportunityDetails.address;
              } else {
                // Construct address from available fields
                const addressParts = [
                  opportunityDetails.contactAddressLine2,
                  opportunityDetails.contactCity,
                  opportunityDetails.contactState,
                  opportunityDetails.contactPostcode
                ].filter(Boolean);
                if (addressParts.length > 0) {
                  customerAddress = addressParts.join(', ');
                }
              }
            } catch (addressError) {
              this.logger.warn(`Failed to fetch customer address: ${addressError.message}`);
            }
            
            finalCustomerData = customerData ? {
              ...customerData,
              address: customerData.address || customerAddress,
            } : {
              name: fetchedName,
              email: opportunity.contactEmail || '',
              address: customerAddress,
            };
            this.logger.log(`Auto-fetched customer data: ${finalCustomerName}, ${finalCustomerData.email}, address: ${customerAddress || 'not available'}`);
          }
        } catch (error) {
          this.logger.warn(`Failed to fetch customer data: ${error.message}`);
        }
      }

      if (!finalCustomerData || !finalCustomerName) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: customerData and customerName are required (or provide userId to auto-fetch)',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // TypeScript now knows these are defined after the check above
      const validatedCustomerData = finalCustomerData;
      const validatedCustomerName = finalCustomerName;

      // Call the service method
      const result = await this.docuSealService.createExpressConsentSigningWorkflow(
        pdfPath,
        opportunityId,
        validatedCustomerData,
        validatedCustomerName
      );

      return {
        success: true,
        message: 'Express consent signing workflow created successfully. Email sent to customer.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create express consent signing workflow: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during express consent signing workflow creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a disclaimer template for verification (without submission)
   * POST /docuseal/disclaimer/template
   * Creates a template with fields but doesn't send it to the customer
   * Use this to verify field positions before creating the actual submission
   */
  @Post('disclaimer/template')
  async createDisclaimerTemplate(
    @Body() body: {
      disclaimerPdfPath?: string;
      opportunityId: string;
      customerName: string;
      installerName?: string;
      formData?: {
        unitRate?: string;
        annualGridConsumption?: string;
        annualElectricitySpend?: string;
        standingCharge?: string;
        utilityBillReason?: string;
      };
    }
  ) {
    try {
      const { disclaimerPdfPath, opportunityId, customerName, installerName, formData } = body;

      if (!opportunityId || !customerName) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId and customerName are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Creating disclaimer template (verification only) for opportunity: ${opportunityId}`);

      // Always use the disclaimer template PDF (DocuSeal will handle signing)
      const templatePath = path.join(process.cwd(), 'src', 'disclaimer', 'EPVS_Disclaimer_Template.pdf');
      const pdfPath = disclaimerPdfPath || templatePath;

      if (!fs.existsSync(pdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Disclaimer template PDF not found at: ${templatePath}. Please ensure the file exists.`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      this.logger.log(`Using disclaimer template PDF: ${pdfPath}`);

      const result = await this.docuSealService.createDisclaimerTemplate(
        pdfPath,
        opportunityId,
        customerName,
        installerName,
        formData
      );

      return {
        success: true,
        message: 'Disclaimer template created successfully. Use previewUrl to verify field positions.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create disclaimer template: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during disclaimer template creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a signing workflow for Booking Confirmation only
   * POST /docuseal/booking-confirmation
   * 
   * Accepts either:
   * - Full body with all fields, OR
   * - Just opportunityId (will auto-find PDF and customer data)
   */
  @Post('booking-confirmation')
  async createBookingConfirmationSigning(
    @Body() body: {
      bookConfirmationPdfPath?: string;
      opportunityId: string;
      userId?: string;
      customerData?: {
        name: string;
        email: string;
      };
      customerName?: string;
    }
  ) {
    try {
      const { bookConfirmationPdfPath, opportunityId, customerData, customerName } = body;

      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      // Use template PDF directly if path not provided
      // DocuSeal handles each template separately, so we can use the same template PDF
      let pdfPath = bookConfirmationPdfPath;
      if (!pdfPath) {
        this.logger.log(`PDF path not provided, using template PDF directly for opportunity: ${opportunityId}`);
        pdfPath = path.join(process.cwd(), 'src', 'email_confirmation', 'Confirmation of Booking Letter.pdf');
        this.logger.log(`Using template PDF: ${pdfPath}`);
      }

      // Validate PDF file exists
      if (!fs.existsSync(pdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: `Booking confirmation PDF file not found: ${pdfPath}`,
          },
          HttpStatus.NOT_FOUND
        );
      }

      // Fetch customer data if not provided
      let finalCustomerData = customerData;
      let finalCustomerName = customerName;

      if ((!finalCustomerData || !finalCustomerData.email) && body.userId) {
        this.logger.log(`Customer data not provided, fetching from opportunity: ${opportunityId}`);
        try {
          const customerDetails = await this.opportunitiesService.getCustomerDetails(opportunityId, body.userId);
          if (customerDetails) {
            finalCustomerData = finalCustomerData || {
              name: customerDetails.name || 'Customer',
              email: customerDetails.email || '',
            };
            finalCustomerName = finalCustomerName || customerDetails.name || 'Customer';
          }
        } catch (error) {
          this.logger.warn(`Failed to fetch customer details: ${error.message}`);
        }
      }

      if (!finalCustomerData || !finalCustomerData.email) {
        throw new HttpException(
          {
            success: false,
            message: 'Customer email is required. Please provide customerData.email or ensure opportunity has customer email.',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      if (!finalCustomerName) {
        finalCustomerName = finalCustomerData.name || 'Customer';
      }

      this.logger.log(`Creating booking confirmation signing workflow for opportunity: ${opportunityId}`);

      const result = await this.docuSealService.createBookingConfirmationSigningWorkflow(
        pdfPath,
        opportunityId,
        finalCustomerData,
        finalCustomerName
      );

      return {
        success: true,
        message: 'Booking confirmation signing workflow created successfully. Email sent to customer.',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to create booking confirmation signing workflow: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during booking confirmation signing workflow creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Get submission documents with URLs
   * GET /docuseal/submissions/:submissionId/documents
   */
  @Get('submissions/:submissionId/documents')
  async getSubmissionDocuments(@Param('submissionId') submissionId: string) {
    try {
      if (!submissionId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: submissionId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Getting documents for submission: ${submissionId}`);

      const result = await this.docuSealService.getSubmissionDocuments(submissionId);

      return {
        success: true,
        message: 'Submission documents retrieved successfully',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to get submission documents: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error while retrieving submission documents',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Download signed PDF from submission
   * GET /docuseal/submissions/:submissionId/download
   */
  @Get('submissions/:submissionId/download')
  async downloadSignedDocument(
    @Param('submissionId') submissionId: string,
    @Res() res: Response
  ) {
    try {
      if (!submissionId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: submissionId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Downloading signed document for submission: ${submissionId}`);

      const pdfBuffer = await this.docuSealService.getSignedDocument(submissionId);

      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="signed-document-${submissionId}.pdf"`);
      res.send(pdfBuffer);
    } catch (error) {
      this.logger.error(`Failed to download signed document: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error while downloading signed document',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Get all submission IDs for an opportunity
   * GET /docuseal/submissions/opportunity/:opportunityId
   * Automatically syncs status from DocuSeal API
   */
  @Get('submissions/opportunity/:opportunityId')
  async getSubmissionsByOpportunity(@Param('opportunityId') opportunityId: string) {
    try {
      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Getting submissions for opportunity: ${opportunityId}`);

      const result = await this.docuSealService.getSubmissionsByOpportunity(opportunityId);

      return {
        success: true,
        message: 'Submissions retrieved successfully (status synced from DocuSeal)',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to get submissions: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error while retrieving submissions',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Verify DocuSeal API configuration
   * GET /docuseal/verify-config
   * Returns API key status (first/last chars only for security)
   * Compares the service's loaded key with the current environment variable
   */
  @Get('verify-config')
  async verifyConfig() {
    try {
      const config = this.docuSealService.getDocuSealConfig();
      
      return {
        success: true,
        message: 'DocuSeal configuration loaded',
        ...config,
      };
    } catch (error) {
      this.logger.error(`Failed to verify config: ${error.message}`);
      throw new HttpException(
        {
          success: false,
          message: 'Error verifying configuration',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Process completed submissions for an opportunity and upload to OneDrive
   * POST /docuseal/opportunity/:opportunityId/process-completed
   * Checks all completed submissions in DocuSeal for the opportunity and uploads signed documents
   */
  @Post('opportunity/:opportunityId/process-completed')
  async processCompletedSubmissions(@Param('opportunityId') opportunityId: string) {
    try {
      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Processing completed submissions for opportunity: ${opportunityId}`);

      const result = await this.docuSealService.processCompletedSubmissionsForOpportunity(opportunityId);

      return {
        success: result.success,
        message: result.message,
        data: {
          processed: result.processed,
          errors: result.errors,
        },
      };
    } catch (error) {
      this.logger.error(`Failed to process completed submissions: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error while processing completed submissions',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Refresh status for a specific submission
   * GET /docuseal/submissions/:submissionId/refresh-status
   * Fetches fresh status from DocuSeal API and updates database
   */
  @Get('submissions/:submissionId/refresh-status')
  async refreshSubmissionStatus(@Param('submissionId') submissionId: string) {
    try {
      if (!submissionId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: submissionId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      this.logger.log(`Refreshing status for submission: ${submissionId}`);

      const result = await this.docuSealService.refreshSubmissionStatus(submissionId);

      return {
        success: true,
        message: 'Submission status refreshed successfully',
        data: result,
      };
    } catch (error) {
      this.logger.error(`Failed to refresh submission status: ${error.message}`);
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        {
          success: false,
          message: 'Internal server error while refreshing submission status',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Webhook endpoint for DocuSeal events
   * POST /docuseal/webhook
   * Receives real-time notifications from DocuSeal about document events
   * No authentication required - uses webhook secret for verification
   * 
   * Event types handled:
   * - submission.completed: Document has been signed
   * - submission.declined: Document signing was declined
   * - submission.expired: Document signing expired
   * - submission.viewed: Document was viewed
   * - submission.started: Document signing was started
   */
  @Post('webhook')
  @HttpCode(HttpStatus.OK)
  async handleWebhook(
    @Body() payload: any,
    @Headers() headers: Record<string, string>
  ) {
    try {
      this.logger.log(`📥 Received DocuSeal webhook event: ${payload.event || 'unknown'}`);
      this.logger.debug(`Webhook payload: ${JSON.stringify(payload, null, 2)}`);

      // Verify webhook secret if configured
      const secretVerified = await this.docuSealService.verifyWebhookSecret(headers, payload);
      if (!secretVerified) {
        this.logger.warn('⚠️ Webhook secret verification failed - request may not be from DocuSeal');
        // Continue processing but log warning
        // In production, you might want to reject here: throw new HttpException('Unauthorized', HttpStatus.UNAUTHORIZED);
      }

      // Handle different event types
      const eventType = payload.event || payload.type;
      const submissionId = payload.submission?.id?.toString() || payload.submission_id?.toString();

      if (!submissionId) {
        this.logger.warn('⚠️ Webhook payload missing submission ID');
        return {
          success: true,
          message: 'Webhook received but no submission ID found',
        };
      }

      this.logger.log(`Processing webhook event: ${eventType} for submission: ${submissionId}`);

      // Process the webhook event
      await this.docuSealService.handleWebhookEvent(eventType, payload, submissionId);

      return {
        success: true,
        message: 'Webhook processed successfully',
      };
    } catch (error) {
      this.logger.error(`❌ Error processing webhook: ${error.message}`);
      this.logger.error(`Stack trace: ${error.stack}`);
      
      // Return 200 to prevent DocuSeal from retrying
      // Log the error for manual investigation
      return {
        success: false,
        message: 'Webhook processing error (logged for investigation)',
        error: error.message,
      };
    }
  }
}

