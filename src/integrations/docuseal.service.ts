import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import axios from 'axios';
import * as FormData from 'form-data';
import * as fs from 'fs';
import * as path from 'path';
import { PDFDocument } from 'pdf-lib';
import * as jwt from 'jsonwebtoken';
import { OneDriveFileManagerService } from '../onedrive/onedrive-file-manager.service';

export interface DocuSealTemplate {
  id: string;
  name: string;
  slug?: string;
  template_slug?: string;
  created_at: string;
  updated_at: string;
}

export interface DocuSealSubmission {
  id: number;
  template_id?: number;
  status: string;
  created_at: string;
  completed_at?: string;
  audit_log_url?: string;
  combined_document_url?: string;
  submitters: Array<{
    id: number;
    email: string;
    name: string;
    status: string;
    slug?: string;
    role?: string;
  }>;
  signers?: Array<{
    id: number;
    email: string;
    name: string;
    status: string;
    slug?: string;
  }>;
}

export interface DocuSealSubmitter {
  id: number;
  slug: string;
  uuid: string;
  name: string;
  email: string;
  phone: string | null;
  completed_at: string | null;
  declined_at: string | null;
  external_id: string | null;
  submission_id: number;
  metadata: any;
  opened_at: string | null;
  sent_at: string | null;
  created_at: string;
  updated_at: string;
  status: string;
  application_key: string | null;
  values: any[];
  preferences: {
    send_email: boolean;
    send_sms: boolean;
  };
  role: string;
  embed_src: string;
}

export interface DocuSealField {
  id: string;
  name: string;
  type: 'text' | 'signature' | 'date' | 'checkbox';
  x: number;
  y: number;
  width: number;
  height: number;
  page: number;
  required: boolean;
}

@Injectable()
export class DocuSealService {
  private readonly logger = new Logger(DocuSealService.name);
  private readonly baseUrl = process.env.DOCUSEAL_BASE_URL || 'https://api.docuseal.com';
  private readonly apiKey: string;

  constructor(
    private readonly prisma: PrismaService,
    @Inject(forwardRef(() => OneDriveFileManagerService))
    private readonly oneDriveFileManagerService: OneDriveFileManagerService
  ) {
    const apiKey = process.env.DOCUSEAL_API_KEY;
    if (!apiKey) {
      this.logger.error('DOCUSEAL_API_KEY is not set in environment variables - DocuSeal API calls will fail');
      throw new Error('DOCUSEAL_API_KEY environment variable is required');
    }
    // Trim whitespace and newlines that might be in .env file
    this.apiKey = apiKey.trim();
    // Log first and last 4 characters for verification (without exposing full key)
    const keyPreview = apiKey.length > 8 
      ? `${apiKey.substring(0, 4)}...${apiKey.substring(apiKey.length - 4)}`
      : '***';
    this.logger.log(`DocuSeal service initialized with base URL: ${this.baseUrl}`);
    this.logger.log(`DocuSeal API Key loaded: ${keyPreview} (length: ${apiKey.length})`);
  }

  /**
   * Get the correct API endpoint path
   * Cloud API (both .com and .eu) uses paths without /api/ prefix, self-hosted uses /api/ prefix
   */
  private getApiPath(path: string): string {
    if (this.baseUrl.includes('api.docuseal.com') || this.baseUrl.includes('api.docuseal.eu')) {
      // Cloud API: no /api/ prefix
      return `${this.baseUrl}${path}`;
    } else {
      // Self-hosted: use /api/ prefix
      return `${this.baseUrl}/api${path}`;
    }
  }

  /**
   * Create a template from PDF file
   * Uploads a PDF and creates a template for signing
   */
  async createTemplateFromPdf(
    pdfPath: string,
    templateName: string,
    opportunityId: string
  ): Promise<DocuSealTemplate> {
    try {
      this.logger.log(`Creating template from PDF: ${templateName}`);

      // Read PDF file
      const pdfBuffer = fs.readFileSync(pdfPath);
      const fileName = path.basename(pdfPath);

      // Create form data
      const formData = new FormData();
      formData.append('template[name]', `${templateName} - ${opportunityId}`);
      formData.append('template[source]', pdfBuffer, {
        filename: fileName,
        contentType: 'application/pdf',
      });

      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates'),
        formData,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            ...formData.getHeaders(),
          },
        }
      );

      this.logger.log(`Template created successfully from PDF: ${response.data.id}`);
      return response.data;
    } catch (error) {
      this.logger.error(`Failed to create template from PDF: ${error.message}`);
      throw error;
    }
  }

  /**
   * Create a template from HTML content with signature fields
   * This works with the free version of DocuSeal v2.1.5
   */
  async createTemplateFromHtml(
    htmlContent: string,
    templateName: string,
    opportunityId: string
  ): Promise<DocuSealTemplate> {
    try {
      this.logger.log(`Creating template from HTML: ${templateName}`);

      const requestBody = {
        name: `${templateName} - ${opportunityId}`,
        html: htmlContent,
        size: "Letter"
      };

      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/html'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      this.logger.log(`Template created successfully: ${response.data.id}`);
      return response.data;
    } catch (error) {
      this.logger.error(`Failed to create template from HTML: ${error.message}`);
      throw error;
    }
  }

  /**
   * Create a submission using an existing template
   * This works with the free version of DocuSeal v2.1.5
   */
  async createSubmissionFromTemplate(
    templateId: string,
    submissionName: string,
    signers: Array<{
      email: string;
      name: string;
      role: string;
    }>,
    opportunityId: string,
    fieldValues?: Record<string, any>
  ): Promise<DocuSealSubmitter[]> {
    try {
      this.logger.log(`Creating submission from template: ${templateId}`);
      this.logger.log(`Sending email to customer(s): ${signers.map(s => s.email).join(', ')}`);

      // Build values object - only include fields that are explicitly provided
      // Don't hardcode field names as they vary by template type
      const values: Record<string, any> = fieldValues || {};

      const requestBody = {
        template_id: parseInt(templateId, 10), // Ensure it's a proper integer
        submitters: signers.map(signer => ({
          name: signer.name,
          role: signer.role || "Signer1", // Use the role from signer or default to Signer1 (must match template)
          email: signer.email,
          send_email: true, // Ensure email is sent to customer (per submitter)
          ...(Object.keys(values).length > 0 ? { values } : {}) // Only include values if there are any
        })),
        send_email: true // Also at root level for backward compatibility
      };

      this.logger.log(`Creating submission with template_id: ${templateId}, role: ${signers[0]?.role || "Signer1"}`);
      this.logger.log(`Request body: ${JSON.stringify(requestBody, null, 2)}`);

      const response = await axios.post<DocuSealSubmitter[]>(
        this.getApiPath('/submissions'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      this.logger.log(`Submission created successfully: ${response.data[0].submission_id}`);
      return response.data;
    } catch (error: any) {
      this.logger.error(`Failed to create submission from template: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data, null, 2)}`);
        this.logger.error(`Template ID: ${templateId}, Submitter emails: ${signers.map(s => s.email).join(', ')}`);
      }
      throw error;
    }
  }

  /**
   * Add fields to a template (text fields, signature boxes, etc.)
   */
  async addFieldsToTemplate(
    templateId: string,
    fields: DocuSealField[]
  ): Promise<void> {
    try {
      this.logger.log(`Adding ${fields.length} fields to template: ${templateId}`);

      for (const field of fields) {
        await axios.post(
          this.getApiPath(`/templates/${templateId}/fields`),
          {
            field: {
              name: field.name,
              type: field.type,
              x: field.x,
              y: field.y,
              width: field.width,
              height: field.height,
              page: field.page,
              required: field.required,
            },
          },
          {
            headers: {
              'X-Auth-Token': this.apiKey,
              'Content-Type': 'application/json',
            },
          }
        );
      }

      this.logger.log(`Successfully added fields to template: ${templateId}`);
    } catch (error) {
      this.logger.error(`Failed to add fields to template: ${error.message}`);
      throw error;
    }
  }

  /**
   * Create a submission (signing request) from a template
   */
  async createSubmission(
    templateId: string,
    signers: Array<{
      email: string;
      name: string;
      role: string;
    }>,
    opportunityId: string
  ): Promise<DocuSealSubmission> {
    try {
      this.logger.log(`Creating submission for template: ${templateId}`);

      // Convert template_id to integer as required by API
      const templateIdInt = parseInt(templateId, 10);
      if (isNaN(templateIdInt)) {
        throw new Error(`Invalid template_id: ${templateId}. Must be a number.`);
      }

      const response = await axios.post<DocuSealSubmission>(
        this.getApiPath('/submissions'),
        {
          template_id: templateIdInt,
          send_email: true,
          submitters: signers.map(signer => ({
            email: signer.email,
            name: signer.name,
            role: signer.role,
            external_id: opportunityId,
          })),
        },
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      this.logger.log(`Submission created successfully: ${response.data.id}`);
      return response.data;
    } catch (error) {
      this.logger.error(`Failed to create submission: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Get submission status
   */
  async getSubmissionStatus(submissionId: string): Promise<DocuSealSubmission> {
    try {
      const response = await axios.get<DocuSealSubmission>(
        this.getApiPath(`/submissions/${submissionId}`),
        {
          headers: {
            'X-Auth-Token': this.apiKey,
          },
        }
      );

      return response.data;
    } catch (error) {
      this.logger.error(`Failed to get submission status: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get submission documents with URLs
   * Returns list of documents with downloadable URLs
   */
  async getSubmissionDocuments(submissionId: string): Promise<{
    id: number;
    documents: Array<{
      name: string;
      url: string;
    }>;
  }> {
    try {
      const response = await axios.get<{
        id: number;
        documents: Array<{
          name: string;
          url: string;
        }>;
      }>(
        this.getApiPath(`/submissions/${submissionId}/documents`),
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Accept': 'application/json',
          },
        }
      );

      this.logger.log(`Retrieved ${response.data.documents.length} document(s) for submission ${submissionId}`);
      return response.data;
    } catch (error) {
      this.logger.error(`Failed to get submission documents: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Get signed document (download as Buffer)
   */
  async getSignedDocument(submissionId: string): Promise<Buffer> {
    try {
      const response = await axios.get<ArrayBuffer>(
        this.getApiPath(`/submissions/${submissionId}/download`),
        {
          headers: {
            'X-Auth-Token': this.apiKey,
          },
          responseType: 'arraybuffer',
        }
      );

      return Buffer.from(response.data);
    } catch (error) {
      this.logger.error(`Failed to get signed document: ${error.message}`);
      throw error;
    }
  }

  /**
   * Download document from URL (download as Buffer)
   * Downloads a document from a direct URL (e.g., from webhook payload)
   */
  async downloadDocumentFromUrl(documentUrl: string): Promise<Buffer> {
    try {
      const response = await axios.get<ArrayBuffer>(
        documentUrl,
        {
          responseType: 'arraybuffer',
        }
      );

      return Buffer.from(response.data);
    } catch (error) {
      this.logger.error(`Failed to download document from URL: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get audit log (download as Buffer)
   * Downloads the audit log PDF from the audit_log_url
   */
  async getAuditLog(auditLogUrl: string): Promise<Buffer> {
    try {
      const response = await axios.get<ArrayBuffer>(
        auditLogUrl,
        {
          responseType: 'arraybuffer',
        }
      );

      return Buffer.from(response.data);
    } catch (error) {
      this.logger.error(`Failed to get audit log: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get signing URL for a specific signer
   */
  async getSigningUrl(submissionId: string, signerId: string): Promise<string> {
    try {
      // Get the submission details to find the submitter slug
      const submission = await this.getSubmissionStatus(submissionId);
      const submitter = submission.submitters.find(s => s.id.toString() === signerId);
      
      if (!submitter) {
        throw new Error(`Signer with ID ${signerId} not found in submission ${submissionId}`);
      }

      // Return the embed URL format for DocuSeal
      return `${this.baseUrl}/s/${submitter.slug || signerId}`;
    } catch (error) {
      this.logger.error(`Failed to get signing URL: ${error.message}`);
      throw error;
    }
  }

  /**
   * Create a complete signing workflow for a contract
   * Uploads PDF, creates template, creates submission, and sends email to customer
   * This is the main method for your workflow:
   * 1. Backend generates contract PDF
   * 2. Backend uploads to DocuSeal (this method)
   * 3. DocuSeal sends email to customer (automatic when send_email: true)
   * 4. Customer signs via email link
   */
  async createContractSigningWorkflow(
    pdfPath: string,
    opportunityId: string,
    customerData: {
      name: string;
      email: string;
    },
    contractData: {
      customerName: string;
      date: string;
      postcode: string;
      [key: string]: any;
    }
  ): Promise<{
    templateId: string;
    submissionId: string;
    signingUrl: string;
  }> {
    try {
      this.logger.log(`Creating contract signing workflow for opportunity: ${opportunityId}`);

      // Step 1: Upload PDF and create template
      this.logger.log(`Uploading PDF and creating template: ${pdfPath}`);
      const template = await this.createTemplateFromPdf(
        pdfPath,
        `Contract - ${contractData.customerName}`,
        opportunityId
      );

      const templateId = template.id.toString();

      // Step 2: Create submission with email sending enabled
      // DocuSeal will automatically send email to customer
      this.logger.log(`Creating submission and sending email to ${customerData.email}`);
      const submitters = await this.createSubmissionFromTemplate(
        templateId,
        `Contract - ${contractData.customerName}`,
        [
          {
            email: customerData.email,
            name: customerData.name,
            role: 'customer',
          },
        ],
        opportunityId
      );

      // Get the first submitter and create signing URL
      const firstSubmitter = submitters[0];
      const signingUrl = firstSubmitter.embed_src;

      this.logger.log(`Contract signing workflow created successfully for opportunity: ${opportunityId}`);
      this.logger.log(`Email sent to: ${customerData.email}`);
      this.logger.log(`Signing URL: ${signingUrl}`);

      return {
        templateId: templateId,
        submissionId: firstSubmitter.submission_id.toString(),
        signingUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to create contract signing workflow: ${error.message}`);
      throw error;
    }
  }

  /**
   * Convert PDF file to base64 string
   */
  private convertPdfToBase64(pdfPath: string): string {
    try {
      this.logger.log(`Converting PDF to base64: ${pdfPath}`);
      const pdfBuffer = fs.readFileSync(pdfPath);
      const base64String = pdfBuffer.toString('base64');
      this.logger.log(`PDF converted to base64 successfully (${base64String.length} characters)`);
      return base64String;
    } catch (error) {
      this.logger.error(`Failed to convert PDF to base64: ${error.message}`);
      throw error;
    }
  }

  /**
   * Create a template from base64 PDF with signature fields
   * Uses /templates/pdf endpoint (cloud API) or /api/templates/pdf (self-hosted)
   */
  async createTemplateFromBase64Pdf(
    pdfPath: string,
    templateName: string,
    fields: Array<{
      name: string;
      type: 'signature' | 'text' | 'date' | 'checkbox';
      role: string;
      required?: boolean;
      areas: Array<{
        page: number;
        x: number; // Normalized coordinates (0-1)
        y: number; // Normalized coordinates (0-1)
        w: number; // Normalized width (0-1)
        h: number; // Normalized height (0-1)
      }>;
    }>,
    opportunityId: string
  ): Promise<DocuSealTemplate> {
    try {
      this.logger.log(`Creating template from base64 PDF: ${templateName}`);

      // Convert PDF to base64
      const pdfBase64 = this.convertPdfToBase64(pdfPath);

      // Prepare the request body according to Postman collection format
      const requestBody = {
        name: `${templateName} - ${opportunityId}`,
        documents: [
          {
            name: templateName,
            file: pdfBase64,
            fields: fields.map(field => ({
              name: field.name,
              type: field.type,
              role: field.role,
              required: field.required !== undefined ? field.required : true, // Default to required if not specified
              areas: field.areas.map(area => ({
                page: area.page,
                x: area.x,
                y: area.y,
                w: area.w,
                h: area.h,
              })),
            })),
          },
        ],
        external_id: opportunityId,
      };

      this.logger.log(`Sending template creation request with ${fields.length} field(s)`);

      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/pdf'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      this.logger.log(`Template created successfully from base64 PDF: ${response.data.id}`);
      return response.data;
    } catch (error) {
      this.logger.error(`Failed to create template from base64 PDF: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Create a complete signing workflow for a contract using base64 PDF
   * This method uses base64 format with signature fields
   */
  async createContractSigningWorkflowWithBase64(
    pdfPath: string,
    opportunityId: string,
    customerData: {
      name: string;
      email: string;
    },
    contractData: {
      customerName: string;
      date: string;
      postcode: string;
      [key: string]: any;
    }
  ): Promise<{
    templateId: string;
    submissionId: string;
    signingUrl: string;
  }> {
    try {
      this.logger.log(`Creating contract signing workflow with base64 PDF for opportunity: ${opportunityId}`);

      // Define contract fields based on exact coordinates from DocuSeal template
      // These coordinates match the exact template structure
      // Page numbers are incremented by 1 from the original template
      const contractFields = [
        {
          name: 'Signature',
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: 6,
              x: 0.4165854551733994,
              y: 0.8750314841982466,
              w: 0.2653658834317836,
              h: 0.03569636050032998,
            },
          ],
        },
        {
          name: 'Signature 2',
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: 19,
              x: 0.3356507240853658,
              y: 0.7162867757529695,
              w: 0.4640442787728659,
              h: 0.0292775886270974,
            },
          ],
        },
        {
          name: 'Signature 3',
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: 21,
              x: 0.3040649711794969,
              y: 0.721342383107089,
              w: 0.2334960044302592,
              h: 0.06083455956000194,
            },
          ],
        },
        {
          name: 'Signature 4',
          type: 'signature' as const,
          role: 'Signer1',
          required: false,
          areas: [
            {
              page: 21,
              x: 0.5388616496760671,
              y: 0.7218451176293835,
              w: 0.2627643864329268,
              h: 0.05681249926352749,
            },
          ],
        },
        {
          name: 'Signature 5',
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: 23,
              x: 0.1986991175209604,
              y: 0.5033307888209371,
              w: 0.3466667361375763,
              h: 0.04675716440422317,
            },
          ],
        },
      ];

      // Step 1: Create template from base64 PDF with contract fields
      this.logger.log(`Creating template from base64 PDF with contract fields: ${pdfPath}`);
      const template = await this.createTemplateFromBase64Pdf(
        pdfPath,
        `Contract - ${contractData.customerName}`,
        contractFields,
        opportunityId
      );

      const templateId = template.id.toString();

      // Step 2: Create submission using the template
      // According to Postman collection, /submissions returns an array of submitters
      this.logger.log(`Creating submission and sending email to ${customerData.email}`);
      const submitters = await this.createSubmissionFromTemplate(
        templateId,
        `Contract - ${contractData.customerName}`,
        [
          {
            email: customerData.email,
            name: customerData.name,
            role: 'Signer1', // Match the role in signature fields
          },
        ],
        opportunityId
      );

      if (!submitters || submitters.length === 0) {
        throw new Error('No submitters returned from submission creation');
      }

      const firstSubmitter = submitters[0];
      
      // Build signing URL - check multiple possible fields from Postman response
      let signingUrl: string;
      if ((firstSubmitter as any)?.embed_src) {
        signingUrl = (firstSubmitter as any).embed_src;
      } else if (firstSubmitter.slug) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.slug}`;
      } else if (firstSubmitter.uuid) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.uuid}`;
      } else {
        signingUrl = 'Email sent to customer';
      }

      // Get submission ID from the submitter
      const submissionId = firstSubmitter.submission_id?.toString() || 'unknown';

      // Save submission ID to database for tracking
      try {
        await this.saveSubmissionToDatabase(
          opportunityId,
          'CONTRACT',
          templateId,
          submissionId,
          signingUrl,
          contractData.customerName,
          customerData.email
        );
        this.logger.log(`Saved contract submission to database: ${submissionId}`);
      } catch (dbError) {
        this.logger.warn(`Failed to save submission to database: ${dbError.message}. Continuing anyway.`);
      }

      this.logger.log(`Contract signing workflow created successfully for opportunity: ${opportunityId}`);
      this.logger.log(`Email sent to: ${customerData.email}`);
      this.logger.log(`Template ID: ${templateId}`);
      this.logger.log(`Submission ID: ${submissionId}`);
      this.logger.log(`Signing URL: ${signingUrl}`);

      return {
        templateId: templateId,
        submissionId,
        signingUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to create contract signing workflow with base64: ${error.message}`);
      throw error;
    }
  }

  /**
   * Merge two PDFs into a single PDF document
   */
  private async mergePdfs(pdf1Path: string, pdf2Path: string, outputPath: string): Promise<void> {
    try {
      this.logger.log(`Merging PDFs: ${pdf1Path} + ${pdf2Path} -> ${outputPath}`);
      
      // Verify both PDFs exist
      if (!fs.existsSync(pdf1Path)) {
        throw new Error(`First PDF not found: ${pdf1Path}`);
      }
      if (!fs.existsSync(pdf2Path)) {
        throw new Error(`Second PDF not found: ${pdf2Path}`);
      }
      
      // Read both PDFs
      const pdf1Bytes = fs.readFileSync(pdf1Path);
      const pdf2Bytes = fs.readFileSync(pdf2Path);
      
      this.logger.log(`PDF 1 size: ${pdf1Bytes.length} bytes, PDF 2 size: ${pdf2Bytes.length} bytes`);

      // Create a new PDF document
      const mergedPdf = await PDFDocument.create();

      // Load and copy pages from first PDF
      this.logger.log(`Loading first PDF: ${pdf1Path}`);
      const pdf1 = await PDFDocument.load(pdf1Bytes);
      const pdf1PageCount = pdf1.getPageCount();
      this.logger.log(`First PDF has ${pdf1PageCount} pages`);
      const pdf1Pages = await mergedPdf.copyPages(pdf1, pdf1.getPageIndices());
      pdf1Pages.forEach((page) => mergedPdf.addPage(page));

      // Load and copy pages from second PDF
      this.logger.log(`Loading second PDF: ${pdf2Path}`);
      const pdf2 = await PDFDocument.load(pdf2Bytes);
      const pdf2PageCount = pdf2.getPageCount();
      this.logger.log(`Second PDF has ${pdf2PageCount} pages`);
      const pdf2Pages = await mergedPdf.copyPages(pdf2, pdf2.getPageIndices());
      pdf2Pages.forEach((page) => mergedPdf.addPage(page));

      // Save merged PDF
      this.logger.log(`Saving merged PDF to: ${outputPath}`);
      const mergedPdfBytes = await mergedPdf.save();
      fs.writeFileSync(outputPath, mergedPdfBytes);
      
      const totalPages = mergedPdf.getPageCount();
      this.logger.log(`Successfully merged PDFs. Total pages: ${totalPages} (${pdf1PageCount} + ${pdf2PageCount})`);
    } catch (error) {
      this.logger.error(`Failed to merge PDFs: ${error.message}`);
      this.logger.error(`Stack trace: ${error.stack}`);
      throw error;
    }
  }

  /**
   * Create a combined signing workflow for contract and booking confirmation
   * Merges both PDFs into a single document before sending to DocuSeal to reduce costs
   * @param calculatorType - 'flux' or 'off-peak' to determine which field coordinates to use
   */
  async createContractAndBookingConfirmationSigningWorkflow(
    contractPdfPath: string,
    bookConfirmationPdfPath: string,
    opportunityId: string,
    customerData: {
      name: string;
      email: string;
    },
    contractData: {
      customerName: string;
      date: string;
      postcode: string;
      [key: string]: any;
    },
    calculatorType: 'flux' | 'off-peak' = 'flux'
  ): Promise<{
    templateId: string;
    submissionId: string;
    signingUrl: string;
  }> {
    try {
      this.logger.log(`Creating combined signing workflow (contract + booking confirmation) for opportunity: ${opportunityId} (calculator type: ${calculatorType})`);

      // 1. Verify booking confirmation PDF exists
      if (!fs.existsSync(bookConfirmationPdfPath)) {
        throw new Error(`Booking confirmation PDF not found at ${bookConfirmationPdfPath}. Booking confirmation is required for all contract signings.`);
      }

      // 2. Get contract page count to calculate booking confirmation page number and adjust signature field pages
      const contractPdfBytes = fs.readFileSync(contractPdfPath);
      const contractPdf = await PDFDocument.load(contractPdfBytes);
      const contractPageCount = contractPdf.getPageCount();
      const bookingConfirmationPage = contractPageCount + 1; // Booking confirmation starts after contract pages
      
      this.logger.log(`Contract has ${contractPageCount} pages. Booking confirmation will be on page ${bookingConfirmationPage}`);

      // 3. Merge the two PDFs into a single document
      const tempDir = path.join(process.cwd(), 'temp');
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }
      
      const mergedPdfPath = path.join(tempDir, `merged_contract_${opportunityId}_${Date.now()}.pdf`);
      await this.mergePdfs(contractPdfPath, bookConfirmationPdfPath, mergedPdfPath);
      
      this.logger.log(`Merged PDF created at: ${mergedPdfPath}`);

      // 4. Convert merged PDF to base64
      const mergedPdfBase64 = this.convertPdfToBase64(mergedPdfPath);

      // 5. Define signature fields for the merged document based on calculator type
      // Base page numbers (for 23-page contracts)
      let contractSignatureFields: Array<{
        name: string;
        type: 'signature' | 'text';
        role: string;
        required?: boolean;
        areas: Array<{ page: number; x: number; y: number; w: number; h: number }>;
      }>;

      if (calculatorType === 'off-peak') {
        // Off Peak coordinates (from DocuSeal template - coordinates only)
        // All fields are signature type
        // Base pages are for 23-page contracts (6, 19, 21, 21, 23)
        // If contract is 24 pages, pages will be adjusted to (7, 20, 22, 22, 24)
        // Each field has unique name with page number for sidebar visibility
        contractSignatureFields = [
          {
            name: 'Signature - Page 6',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 6, x: 0.4780487804878049, y: 0.8754961427477769, w: 0.2439024390243902, h: 0.04386068198944992 },
            ],
          },
          {
            name: 'Signature - Page 19',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 19, x: 0.3346341463414634, y: 0.6974142803315749, w: 0.4653658536585366, h: 0.03240391861341374 },
            ],
          },
          {
            name: 'Signature - Page 21 (Left)',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 21, x: 0.3063414634146341, y: 0.7062452901281085, w: 0.2263414634146342, h: 0.05124340617935186 },
            ],
          },
          {
            name: 'Signature - Page 21 (Right)',
            type: 'signature' as const,
            role: 'Signer1',
            required: false,
            areas: [
              { page: 21, x: 0.5395121951219513, y: 0.705491710625471, w: 0.2604878048780488, h: 0.05124340617935186 },
            ],
          },
          {
            name: 'Signature - Page 23',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 23, x: 0.2, y: 0.4957846646571213, w: 0.3326829268292683, h: 0.0452147701582517 },
            ],
          },
        ];
      } else {
        // Flux coordinates (from DocuSeal template - coordinates only)
        // All fields are signature type
        // Base pages are for 23-page contracts
        // Each field has unique name with page number for sidebar visibility
        contractSignatureFields = [
          {
            name: 'Signature - Page 6',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 6, x: 0.4224390243902439, y: 0.8848124374300689, w: 0.2429268292682927, h: 0.05113452967315257 },
            ],
          },
          {
            name: 'Signature - Page 19',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 19, x: 0.335609756097561, y: 0.7278282780708365, w: 0.464390243902439, h: 0.02938960060286355 },
            ],
          },
          {
            name: 'Signature - Page 21 (Left)',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 21, x: 0.3092682926829268, y: 0.7313842313489073, w: 0.2243902439024391, h: 0.06028636021100231 },
            ],
          },
          {
            name: 'Signature - Page 21 (Right)',
            type: 'signature' as const,
            role: 'Signer1',
            required: false,
            areas: [
              { page: 21, x: 0.5404878048780488, y: 0.7306306518462697, w: 0.2604878048780488, h: 0.06254709871891484 },
            ],
          },
          {
            name: 'Signature - Page 23',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 23, x: 0.1941463414634146, y: 0.5111241976326483, w: 0.3082926829268293, h: 0.04898266767143933 },
            ],
          },
        ];
      }

      // 6. Adjust contract signature field page numbers based on contract page count
      // If contract is 24 pages (without booking confirmation), add 1 to all contract signature pages
      // If contract is 23 pages, keep pages as is
      if (contractPageCount === 24) {
        this.logger.log(`Contract is 24 pages - adjusting all contract signature field pages by +1`);
        contractSignatureFields.forEach(field => {
          field.areas.forEach(area => {
            area.page = area.page + 1;
          });
        });
        this.logger.log(`Adjusted contract signature field pages for 24-page contract`);
      } else if (contractPageCount === 23) {
        this.logger.log(`Contract is 23 pages - keeping contract signature field pages as is`);
      } else {
        this.logger.warn(`Contract has ${contractPageCount} pages (expected 23 or 24) - keeping signature field pages as defined`);
      }

      const currentDate = new Date().toLocaleDateString('en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });

      // Booking confirmation fields are on the page after the contract (dynamically calculated)
      // Use unique names with page number for clarity
      const bookConfirmationSignatureFields = [
        {
          name: `Signature - Page ${bookingConfirmationPage}`,
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: bookingConfirmationPage, // Dynamically calculated based on contract page count
              x: 0.3638886142240891,
              y: 0.6804700876227533,
              w: 0.2926829268292683,
              h: 0.03519668737060044,
            },
          ],
        },
        {
          name: `Full Name - Page ${bookingConfirmationPage}`,
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: bookingConfirmationPage, // Dynamically calculated based on contract page count
              x: 0.3873007373062174,
              y: 0.7563829362475439,
              w: 0.2829268292682927,
              h: 0.030303030303030304,
            },
          ],
        },
        {
          name: `Date Signed - Page ${bookingConfirmationPage}`,
          type: 'date' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: bookingConfirmationPage, // Dynamically calculated based on contract page count
              x: 0.3265432098765432,
              y: 0.8333333333333334,
              w: 0.24691358024691357,
              h: 0.030303030303030304,
            },
          ],
        },
      ];

      // 4. Combine all fields for the single merged document
      const allFields = [
        ...contractSignatureFields.map(field => ({
          name: field.name,
          type: field.type,
          role: field.role,
          areas: field.areas.map(area => ({
            page: area.page,
            x: area.x,
            y: area.y,
            w: area.w,
            h: area.h,
          })),
        })),
        ...bookConfirmationSignatureFields.map(field => ({
          name: field.name,
          type: field.type,
          role: field.role,
          required: field.required,
          areas: field.areas.map(area => ({
            page: area.page,
            x: area.x,
            y: area.y,
            w: area.w,
            h: area.h,
          })),
        })),
      ];

      // 5. Prepare request body with single merged document
      const requestBody = {
        name: `Contract & Booking Confirmation - ${contractData.customerName} - ${opportunityId}`,
        documents: [
          {
            name: `Contract & Booking Confirmation - ${contractData.customerName}`,
            file: mergedPdfBase64,
            fields: allFields,
          },
        ],
        external_id: opportunityId,
      };

      this.logger.log(`Creating template with merged PDF (single document)`);
      this.logger.log(`Merged PDF size: ${mergedPdfBase64.length} characters (base64)`);
      this.logger.log(`Total fields: ${allFields.length} (${contractSignatureFields.length} contract + ${bookConfirmationSignatureFields.length} booking confirmation)`);

      // 6. Create template with merged document
      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/pdf'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      const templateId = response.data.id.toString();
      this.logger.log(`Template created successfully with ID: ${templateId}`);

      // Clean up temporary merged PDF
      try {
        if (fs.existsSync(mergedPdfPath)) {
          fs.unlinkSync(mergedPdfPath);
          this.logger.log(`Cleaned up temporary merged PDF: ${mergedPdfPath}`);
        }
      } catch (cleanupError) {
        this.logger.warn(`Failed to clean up temporary PDF: ${cleanupError.message}`);
      }

      // Create submission
      this.logger.log(`Creating submission and sending email to ${customerData.email}`);
      const submitters = await this.createSubmissionFromTemplate(
        templateId,
        `Contract & Booking Confirmation - ${contractData.customerName}`,
        [
          {
            email: customerData.email,
            name: customerData.name,
            role: 'Signer1',
          },
        ],
        opportunityId,
        {
          "Full Name": customerData.name,
          "Date Signed": currentDate
        }
      );

      if (!submitters || submitters.length === 0) {
        throw new Error('No submitters returned from submission creation');
      }

      const firstSubmitter = submitters[0];
      
      // Build signing URL
      let signingUrl: string;
      if ((firstSubmitter as any)?.embed_src) {
        signingUrl = (firstSubmitter as any).embed_src;
      } else if (firstSubmitter.slug) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.slug}`;
      } else if (firstSubmitter.uuid) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.uuid}`;
      } else {
        signingUrl = 'Email sent to customer';
      }

      const submissionId = firstSubmitter.submission_id?.toString() || 'unknown';

      // Save submission ID to database for both CONTRACT and BOOKING_CONFIRMATION
      // Since they're in the same submission, we use the same submission ID for both
      try {
        await this.saveSubmissionToDatabase(
          opportunityId,
          'CONTRACT',
          templateId,
          submissionId,
          signingUrl,
          contractData.customerName,
          customerData.email
        );
        this.logger.log(`Saved contract submission to database: ${submissionId}`);

        await this.saveSubmissionToDatabase(
          opportunityId,
          'BOOKING_CONFIRMATION',
          templateId,
          submissionId,
          signingUrl,
          contractData.customerName,
          customerData.email
        );
        this.logger.log(`Saved booking confirmation submission to database: ${submissionId}`);
      } catch (dbError) {
        this.logger.warn(`Failed to save submission to database: ${dbError.message}. Continuing anyway.`);
      }

      this.logger.log(`Combined contract & booking confirmation signing workflow created successfully for opportunity: ${opportunityId}`);
      this.logger.log(`Email sent to: ${customerData.email}`);
      this.logger.log(`Template ID: ${templateId}`);
      this.logger.log(`Submission ID: ${submissionId}`);
      this.logger.log(`Signing URL: ${signingUrl}`);

      return {
        templateId,
        submissionId,
        signingUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to create combined contract & booking confirmation signing workflow: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Create a template with signature fields for contract and booking confirmation
   * WITHOUT creating a submission - allows verification of field positions before sending
   * Returns template ID and preview URL for checking field positions
   */
  async createContractAndBookingConfirmationTemplate(
    contractPdfPath: string,
    bookConfirmationPdfPath: string,
    opportunityId: string,
    contractData: {
      customerName: string;
      date: string;
      postcode: string;
      [key: string]: any;
    },
    calculatorType: 'flux' | 'off-peak' = 'flux'
  ): Promise<{
    templateId: string;
    templateSlug: string | null; // Slug for embedded form
    formBuilderToken: string; // JWT token for Form Builder component
    embeddedFormUrl: string | null; // URL for embedded signing form component (fallback)
    previewUrl: string; // Fallback preview URL
    mergedPdfPath: string; // Path to merged PDF for reference
    fields: Array<{
      name: string;
      type: 'signature' | 'text' | 'date';
      page: number;
      x: number;
      y: number;
      w: number;
      h: number;
      required: boolean;
    }>;
  }> {
    try {
      this.logger.log(`Creating template (without submission) for contract & booking confirmation - opportunity: ${opportunityId} (calculator type: ${calculatorType})`);

      // 1. Verify booking confirmation PDF exists
      if (!fs.existsSync(bookConfirmationPdfPath)) {
        throw new Error(`Booking confirmation PDF not found at ${bookConfirmationPdfPath}. Booking confirmation is required for all contract signings.`);
      }

      // 2. Get contract page count to calculate booking confirmation page number and adjust signature field pages
      const contractPdfBytes = fs.readFileSync(contractPdfPath);
      const contractPdf = await PDFDocument.load(contractPdfBytes);
      const contractPageCount = contractPdf.getPageCount();
      const bookingConfirmationPage = contractPageCount + 1; // Booking confirmation starts after contract pages
      
      this.logger.log(`Contract has ${contractPageCount} pages. Booking confirmation will be on page ${bookingConfirmationPage}`);

      // 3. Merge the two PDFs into a single document
      const tempDir = path.join(process.cwd(), 'temp');
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }
      
      const mergedPdfPath = path.join(tempDir, `merged_contract_${opportunityId}_${Date.now()}.pdf`);
      await this.mergePdfs(contractPdfPath, bookConfirmationPdfPath, mergedPdfPath);
      
      this.logger.log(`Merged PDF created at: ${mergedPdfPath}`);

      // 4. Convert merged PDF to base64
      const mergedPdfBase64 = this.convertPdfToBase64(mergedPdfPath);

      // 5. Define signature fields for the merged document based on calculator type
      // Base page numbers (for 23-page contracts)
      let contractSignatureFields: Array<{
        name: string;
        type: 'signature' | 'text';
        role: string;
        required?: boolean;
        areas: Array<{ page: number; x: number; y: number; w: number; h: number }>;
      }>;

      if (calculatorType === 'off-peak') {
        // Off Peak coordinates - each field has unique name with page number
        contractSignatureFields = [
          {
            name: 'Signature - Page 6',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 6, x: 0.4780487804878049, y: 0.8754961427477769, w: 0.2439024390243902, h: 0.04386068198944992 },
            ],
          },
          {
            name: 'Signature - Page 19',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 19, x: 0.3346341463414634, y: 0.6974142803315749, w: 0.4653658536585366, h: 0.03240391861341374 },
            ],
          },
          {
            name: 'Signature - Page 21 (Left)',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 21, x: 0.3063414634146341, y: 0.7062452901281085, w: 0.2263414634146342, h: 0.05124340617935186 },
            ],
          },
          {
            name: 'Signature - Page 21 (Right)',
            type: 'signature' as const,
            role: 'Signer1',
            required: false,
            areas: [
              { page: 21, x: 0.5395121951219513, y: 0.705491710625471, w: 0.2604878048780488, h: 0.05124340617935186 },
            ],
          },
          {
            name: 'Signature - Page 23',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 23, x: 0.2, y: 0.4957846646571213, w: 0.3326829268292683, h: 0.0452147701582517 },
            ],
          },
        ];
      } else {
        // Flux coordinates - each field has unique name with page number
        contractSignatureFields = [
          {
            name: 'Signature - Page 6',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 6, x: 0.4224390243902439, y: 0.8848124374300689, w: 0.2429268292682927, h: 0.05113452967315257 },
            ],
          },
          {
            name: 'Signature - Page 19',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 19, x: 0.335609756097561, y: 0.7278282780708365, w: 0.464390243902439, h: 0.02938960060286355 },
            ],
          },
          {
            name: 'Signature - Page 21 (Left)',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 21, x: 0.3092682926829268, y: 0.7313842313489073, w: 0.2243902439024391, h: 0.06028636021100231 },
            ],
          },
          {
            name: 'Signature - Page 21 (Right)',
            type: 'signature' as const,
            role: 'Signer1',
            required: false,
            areas: [
              { page: 21, x: 0.5404878048780488, y: 0.7306306518462697, w: 0.2604878048780488, h: 0.06254709871891484 },
            ],
          },
          {
            name: 'Signature - Page 23',
            type: 'signature' as const,
            role: 'Signer1',
            required: true,
            areas: [
              { page: 23, x: 0.1941463414634146, y: 0.5111241976326483, w: 0.3082926829268293, h: 0.04898266767143933 },
            ],
          },
        ];
      }

      // 6. Adjust contract signature field page numbers based on contract page count
      if (contractPageCount === 24) {
        this.logger.log(`Contract is 24 pages - adjusting all contract signature field pages by +1`);
        contractSignatureFields.forEach(field => {
          field.areas.forEach(area => {
            area.page = area.page + 1;
          });
        });
      } else if (contractPageCount === 23) {
        this.logger.log(`Contract is 23 pages - keeping contract signature field pages as is`);
      } else {
        this.logger.warn(`Contract has ${contractPageCount} pages (expected 23 or 24) - keeping signature field pages as defined`);
      }

      // Booking confirmation fields are on the page after the contract
      // Use unique names with page number for clarity
      const bookConfirmationSignatureFields = [
        {
          name: `Signature - Page ${bookingConfirmationPage}`,
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: bookingConfirmationPage,
              x: 0.3638886142240891,
              y: 0.6804700876227533,
              w: 0.2926829268292683,
              h: 0.03519668737060044,
            },
          ],
        },
        {
          name: `Full Name - Page ${bookingConfirmationPage}`,
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: bookingConfirmationPage,
              x: 0.3873007373062174,
              y: 0.7563829362475439,
              w: 0.2829268292682927,
              h: 0.030303030303030304,
            },
          ],
        },
        {
          name: `Date Signed - Page ${bookingConfirmationPage}`,
          type: 'date' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: bookingConfirmationPage,
              x: 0.3265432098765432,
              y: 0.8333333333333334,
              w: 0.24691358024691357,
              h: 0.030303030303030304,
            },
          ],
        },
      ];

      // 7. Combine all fields for the single merged document
      const allFields = [
        ...contractSignatureFields.map(field => ({
          name: field.name,
          type: field.type,
          role: field.role,
          required: field.required,
          areas: field.areas.map(area => ({
            page: area.page,
            x: area.x,
            y: area.y,
            w: area.w,
            h: area.h,
          })),
        })),
        ...bookConfirmationSignatureFields.map(field => ({
          name: field.name,
          type: field.type,
          role: field.role,
          required: field.required,
          areas: field.areas.map(area => ({
            page: area.page,
            x: area.x,
            y: area.y,
            w: area.w,
            h: area.h,
          })),
        })),
      ];

      // 8. Prepare request body with single merged document
      const requestBody = {
        name: `Contract & Booking Confirmation - ${contractData.customerName} - ${opportunityId}`,
        documents: [
          {
            name: `Contract & Booking Confirmation - ${contractData.customerName}`,
            file: mergedPdfBase64,
            fields: allFields,
          },
        ],
        external_id: `template-${opportunityId}`,
      };

      this.logger.log(`Creating template with merged PDF (for verification, no submission)`);
      this.logger.log(`Merged PDF size: ${mergedPdfBase64.length} characters (base64)`);
      this.logger.log(`Total fields: ${allFields.length} (${contractSignatureFields.length} contract + ${bookConfirmationSignatureFields.length} booking confirmation)`);

      // 9. Create template with merged document
      // Log API key preview for debugging (first 4 and last 4 chars only)
      const apiKeyPreview = this.apiKey.length > 8 
        ? `${this.apiKey.substring(0, 4)}...${this.apiKey.substring(this.apiKey.length - 4)}`
        : '***';
      this.logger.log(`Using API key: ${apiKeyPreview} (length: ${this.apiKey.length}) for template creation`);
      
      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/pdf'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      const templateId = response.data.id.toString();
      const templateSlug = response.data.slug || response.data.template_slug || null;
      
      this.logger.log(`Template created successfully with ID: ${templateId}`);
      if (templateSlug) {
        this.logger.log(`Template slug: ${templateSlug}`);
      }

      // 10. Get template slug for embedded form
      // If slug is not in initial response, fetch template details
      let finalTemplateSlug = templateSlug;
      if (!finalTemplateSlug) {
        try {
          const templateDetailsResponse = await axios.get<DocuSealTemplate>(
            this.getApiPath(`/templates/${templateId}`),
            {
              headers: {
                'X-Auth-Token': this.apiKey,
              },
            }
          );
          finalTemplateSlug = templateDetailsResponse.data.slug || templateDetailsResponse.data.template_slug || null;
          if (finalTemplateSlug) {
            this.logger.log(`Retrieved template slug from API: ${finalTemplateSlug}`);
          }
        } catch (error) {
          this.logger.warn(`Could not fetch template slug: ${error.message}`);
        }
      }

      // 11. Generate JWT token for Form Builder
      // Get user email from environment or use default
      const userEmail = process.env.DOCUSEAL_USER_EMAIL || 'operations@creativuk.co.uk';
      
      // Generate JWT token for Form Builder with preview mode
      const jwtPayload = {
        user_email: userEmail,
        integration_email: userEmail,
        external_id: `template-${opportunityId}`,
        name: `Contract & Booking Confirmation - ${contractData.customerName}`,
        template_id: parseInt(templateId, 10),
      };
      
      const formBuilderToken = jwt.sign(jwtPayload, this.apiKey, {
        algorithm: 'HS256',
      });
      
      this.logger.log(`JWT token generated for Form Builder (template ID: ${templateId})`);

      // 12. Build embedded form URL (for signing form - fallback)
      // Fix URL generation: https://api.docuseal.eu -> https://docuseal.eu
      let baseUrlForEmbed: string;
      if (this.baseUrl.includes('/api')) {
        // Self-hosted: remove /api
        baseUrlForEmbed = this.baseUrl.replace('/api', '');
      } else if (this.baseUrl.includes('api.docuseal.eu')) {
        // EU cloud: replace api.docuseal.eu with docuseal.eu
        baseUrlForEmbed = this.baseUrl.replace('api.docuseal.eu', 'docuseal.eu');
      } else if (this.baseUrl.includes('api.docuseal.com')) {
        // US cloud: replace api.docuseal.com with docuseal.com
        baseUrlForEmbed = this.baseUrl.replace('api.docuseal.com', 'docuseal.com');
      } else {
        // Fallback: try regex replacement
        baseUrlForEmbed = this.baseUrl.replace(/:\/\/api\./, '://');
      }
      
      const embeddedFormUrl = finalTemplateSlug 
        ? `${baseUrlForEmbed}/d/${finalTemplateSlug}`
        : null;
      
      // Build preview URL (fallback)
      const previewUrl = `${this.baseUrl}/templates/${templateId}`;
      
      this.logger.log(`Form Builder token generated. Embedded form URL: ${embeddedFormUrl || 'Not available (slug missing)'}`);

      // 11. Flatten fields for return (one entry per area)
      const flattenedFields = allFields.flatMap(field =>
        field.areas.map(area => ({
          name: field.name,
          type: field.type,
          page: area.page,
          x: area.x,
          y: area.y,
          w: area.w,
          h: area.h,
          required: field.required !== undefined ? field.required : true,
        }))
      );

      this.logger.log(`Template created for verification. Template ID: ${templateId}`);
      this.logger.log(`Template Slug: ${finalTemplateSlug || 'Not available'}`);
      this.logger.log(`Form Builder Token: Generated (${formBuilderToken.length} characters)`);
      this.logger.log(`Embedded Form URL: ${embeddedFormUrl || 'Not available'}`);
      this.logger.log(`Merged PDF saved at: ${mergedPdfPath} (for reference)`);

      return {
        templateId,
        templateSlug: finalTemplateSlug,
        formBuilderToken,
        embeddedFormUrl,
        previewUrl,
        mergedPdfPath,
        fields: flattenedFields,
      };
    } catch (error) {
      this.logger.error(`Failed to create contract & booking confirmation template: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Create a submission from a verified template (created via createContractAndBookingConfirmationTemplate)
   * Use this after verifying field positions are correct
   */
  async createSubmissionFromVerifiedTemplate(
    templateId: string,
    opportunityId: string,
    customerData: {
      name: string;
      email: string;
    },
    contractData: {
      customerName: string;
      date: string;
      postcode: string;
      [key: string]: any;
    }
  ): Promise<{
    templateId: string;
    submissionId: string;
    signingUrl: string;
  }> {
    try {
      this.logger.log(`Creating submission from verified template ${templateId} for opportunity: ${opportunityId}`);

      const currentDate = new Date().toLocaleDateString('en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });

      // Create submission from the verified template
      this.logger.log(`Creating submission and sending email to ${customerData.email}`);
      const submitters = await this.createSubmissionFromTemplate(
        templateId,
        `Contract & Booking Confirmation - ${contractData.customerName}`,
        [
          {
            email: customerData.email,
            name: customerData.name,
            role: 'Signer1',
          },
        ],
        opportunityId,
        {
          "Full Name": customerData.name,
          "Date Signed": currentDate
        }
      );

      if (!submitters || submitters.length === 0) {
        throw new Error('No submitters returned from submission creation');
      }

      const firstSubmitter = submitters[0];
      
      // Build signing URL
      let signingUrl: string;
      if ((firstSubmitter as any)?.embed_src) {
        signingUrl = (firstSubmitter as any).embed_src;
      } else if (firstSubmitter.slug) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.slug}`;
      } else if (firstSubmitter.uuid) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.uuid}`;
      } else {
        signingUrl = 'Email sent to customer';
      }

      const submissionId = firstSubmitter.submission_id?.toString() || 'unknown';

      // Save submission ID to database for both CONTRACT and BOOKING_CONFIRMATION
      try {
        await this.saveSubmissionToDatabase(
          opportunityId,
          'CONTRACT',
          templateId,
          submissionId,
          signingUrl,
          contractData.customerName,
          customerData.email
        );
        this.logger.log(`Saved contract submission to database: ${submissionId}`);

        await this.saveSubmissionToDatabase(
          opportunityId,
          'BOOKING_CONFIRMATION',
          templateId,
          submissionId,
          signingUrl,
          contractData.customerName,
          customerData.email
        );
        this.logger.log(`Saved booking confirmation submission to database: ${submissionId}`);
      } catch (dbError) {
        this.logger.warn(`Failed to save submission to database: ${dbError.message}. Continuing anyway.`);
      }

      this.logger.log(`Submission created from verified template successfully for opportunity: ${opportunityId}`);
      this.logger.log(`Email sent to: ${customerData.email}`);
      this.logger.log(`Template ID: ${templateId}`);
      this.logger.log(`Submission ID: ${submissionId}`);
      this.logger.log(`Signing URL: ${signingUrl}`);

      return {
        templateId,
        submissionId,
        signingUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to create submission from verified template: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Create a complete signing workflow for contract, disclaimer, and book confirmation
   * Handles all three documents in a single DocuSeal submission
   */
  async createCompleteSigningWorkflow(
    contractPdfPath: string,
    disclaimerPdfPath: string,
    bookConfirmationPdfPath: string,
    opportunityId: string,
    customerData: {
      name: string;
      email: string;
    },
    contractData: {
      customerName: string;
      date: string;
      postcode: string;
      [key: string]: any;
    }
  ): Promise<{
    templateId: string;
    submissionId: string;
    signingUrl: string;
  }> {
    try {
      this.logger.log(`Creating complete signing workflow (contract + disclaimer + book confirmation) for opportunity: ${opportunityId}`);

      // Convert all PDFs to base64
      const contractBase64 = this.convertPdfToBase64(contractPdfPath);
      const disclaimerBase64 = this.convertPdfToBase64(disclaimerPdfPath);
      const bookConfirmationBase64 = this.convertPdfToBase64(bookConfirmationPdfPath);

      // Define signature fields for each document
      // Contract: pages 6, 19, 21, 23 (existing coordinates)
      const contractSignatureFields = [
        {
          name: 'Contract Signature',
          type: 'signature' as const,
          role: 'Signer1',
          areas: [
            { page: 6, x: 0.33, y: 0.88, w: 0.34, h: 0.03 },
            { page: 19, x: 0.33, y: 0.71, w: 0.34, h: 0.03 },
            { page: 21, x: 0.33, y: 0.70, w: 0.24, h: 0.07 },
            { page: 23, x: 0.33, y: 0.50, w: 0.34, h: 0.08 },
          ],
        },
      ];

      // Disclaimer: page 1 (signature at bottom right)
      // Based on signature position {x: 550, y: height - 40} on ~612x792 page
      const disclaimerSignatureFields = [
        {
          name: 'Disclaimer Signature',
          type: 'signature' as const,
          role: 'Signer1',
          areas: [
            { page: 1, x: 0.85, y: 0.05, w: 0.12, h: 0.05 }, // Bottom right area
          ],
        },
      ];

      // Book Confirmation: page 1 (signature position from digital signature service)
      // Based on {x: 200, y: 180} on ~612x792 page
      // Includes signature, customer name, and date fields
      const bookConfirmationSignatureFields = [
        {
          name: 'Book Confirmation Signature',
          type: 'signature' as const,
          role: 'Signer1',
          areas: [
            { page: 1, x: 0.33, y: 0.77, w: 0.25, h: 0.06 }, // Signature area
          ],
        },
        {
          name: 'Full Name',
          type: 'text' as const,
          role: 'Signer1',
          areas: [
            { page: 1, x: 0.39, y: 0.76, w: 0.34, h: 0.03 }, // Customer name field
          ],
        },
        {
          name: 'Date Signed',
          type: 'date' as const,
          role: 'Signer1',
          areas: [
            { page: 1, x: 0.20, y: 0.83, w: 0.34, h: 0.03 }, // Date field
          ],
        },
      ];

      // Prepare request body with all three documents
      const requestBody = {
        name: `Complete Signing Package - ${contractData.customerName} - ${opportunityId}`,
        documents: [
          {
            name: `Contract - ${contractData.customerName}`,
            file: contractBase64,
            fields: contractSignatureFields.map(field => ({
              name: field.name,
              type: field.type,
              role: field.role,
              areas: field.areas.map(area => ({
                page: area.page,
                x: area.x,
                y: area.y,
                w: area.w,
                h: area.h,
              })),
            })),
          },
          {
            name: `Disclaimer - ${contractData.customerName}`,
            file: disclaimerBase64,
            fields: disclaimerSignatureFields.map(field => ({
              name: field.name,
              type: field.type,
              role: field.role,
              areas: field.areas.map(area => ({
                page: area.page,
                x: area.x,
                y: area.y,
                w: area.w,
                h: area.h,
              })),
            })),
          },
          {
            name: `Book Confirmation - ${contractData.customerName}`,
            file: bookConfirmationBase64,
            fields: bookConfirmationSignatureFields.map(field => ({
              name: field.name,
              type: field.type,
              role: field.role,
              areas: field.areas.map(area => ({
                page: area.page,
                x: area.x,
                y: area.y,
                w: area.w,
                h: area.h,
              })),
            })),
          },
        ],
        external_id: opportunityId,
      };

      this.logger.log(`Creating template with 3 documents (contract, disclaimer, book confirmation)`);

      // Create template with all documents
      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/pdf'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      const templateId = response.data.id.toString();
      this.logger.log(`Template created successfully with ID: ${templateId}`);

      // Create submission
      this.logger.log(`Creating submission and sending email to ${customerData.email}`);
      const submitters = await this.createSubmissionFromTemplate(
        templateId,
        `Complete Signing Package - ${contractData.customerName}`,
        [
          {
            email: customerData.email,
            name: customerData.name,
            role: 'Signer1',
          },
        ],
        opportunityId
      );

      if (!submitters || submitters.length === 0) {
        throw new Error('No submitters returned from submission creation');
      }

      const firstSubmitter = submitters[0];
      
      // Build signing URL
      let signingUrl: string;
      if ((firstSubmitter as any)?.embed_src) {
        signingUrl = (firstSubmitter as any).embed_src;
      } else if (firstSubmitter.slug) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.slug}`;
      } else if (firstSubmitter.uuid) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.uuid}`;
      } else {
        signingUrl = 'Email sent to customer';
      }

      const submissionId = firstSubmitter.submission_id?.toString() || 'unknown';

      this.logger.log(`Complete signing workflow created successfully for opportunity: ${opportunityId}`);
      this.logger.log(`Email sent to: ${customerData.email}`);
      this.logger.log(`Template ID: ${templateId}`);
      this.logger.log(`Submission ID: ${submissionId}`);
      this.logger.log(`Signing URL: ${signingUrl}`);

      return {
        templateId,
        submissionId,
        signingUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to create complete signing workflow: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Create a signing workflow for Disclaimer only
   * Sends disclaimer document to customer for signing
   */
  async createDisclaimerSigningWorkflow(
    disclaimerPdfPath: string,
    opportunityId: string,
    customerData: {
      name: string;
      email: string;
    },
    customerName: string,
    installerName?: string,
    formData?: {
      unitRate?: string;
      annualGridConsumption?: string;
      annualElectricitySpend?: string;
      standingCharge?: string;
      utilityBillReason?: string;
    }
  ): Promise<{
    templateId: string;
    submissionId: string;
    signingUrl: string;
  }> {
    try {
      this.logger.log(`Creating disclaimer signing workflow for opportunity: ${opportunityId}`);

      // Convert PDF to base64
      const disclaimerBase64 = this.convertPdfToBase64(disclaimerPdfPath);

      // Get current date for pre-filling
      const currentDate = new Date().toLocaleDateString('en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });

      // Disclaimer fields with correct coordinates (DocuSeal uses 1-indexed pages, so page 1 = first page)
      const disclaimerFields = [
        {
          name: 'Checkbox 1',
          type: 'checkbox' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.1229268292682927, y: 0.3898665746491833, w: 0.03333333333333333, h: 0.02357948010121923 },
          ],
        },
        {
          name: 'Checkbox 2',
          type: 'checkbox' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.1229268292682927, y: 0.4492178513917645, w: 0.03333333333333333, h: 0.02357948010121923 },
          ],
        },
        {
          name: 'Installers Name',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.2461787823932927, y: 0.1555095327731409, w: 0.2308943883384146, h: 0.01656314699792963 },
          ],
        },
        {
          name: 'Customer Name',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.6390243902439025, y: 0.1555095327731409, w: 0.2328455483041159, h: 0.01518288474810214 },
          ],
        },
        {
          name: 'Unit rate figures p per kWh',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.6286128048780488, y: 0.2392403289608416, w: 0.09626030154344511, h: 0.01472280435186982 },
          ],
        },
        {
          name: 'Annual Grid Consumption - figure 2 kWh',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.2448780487804878, y: 0.409477828863656, w: 0.1001625321551067, h: 0.01426263971111541 },
          ],
        },
        {
          name: 'Annual Electricity Spend - amount pounds',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.2955792682926829, y: 0.4701919656811494, w: 0.08781927293794833, h: 0.01357604941217505 },
          ],
        },
        {
          name: 'Standing Charge - per day',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.6559146341463414, y: 0.4681224872543842, w: 0.09919729250210818, h: 0.01702749913457491 },
          ],
        },
        {
          name: 'Utility Bill Reason',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1207164634146341, y: 0.6479946538473975, w: 0.7562204649390244, h: 0.03244189149844723 },
          ],
        },
        {
          name: 'cusomterName-sign',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1735677132955412, y: 0.8027460007440477, w: 0.298101494021532, h: 0.02009990052406829 },
          ],
        },
        {
          name: 'Date Signed',
          type: 'date' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1655917432831555, y: 0.8341300337786836, w: 0.297104521960747, h: 0.02009990052406829 },
          ],
        },
        {
          name: 'Signature',
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1810451600609756, y: 0.8665719252771308, w: 0.2861375762195122, h: 0.01904204206025706 },
          ],
        },
      ];

      // Prepare request body with all fields
      const requestBody = {
        name: `Disclaimer - ${customerName} - ${opportunityId}`,
        documents: [
          {
            name: `Disclaimer - ${customerName}`,
            file: disclaimerBase64,
            fields: disclaimerFields.map(field => {
              const fieldObj: any = {
                name: field.name,
                type: field.type,
                role: field.role,
                required: field.required,
                areas: field.areas.map(area => ({
                  page: area.page,
                  x: area.x,
                  y: area.y,
                  w: area.w,
                  h: area.h,
                })),
              };

              // Add date format preference for date fields
              if (field.type === 'date') {
                fieldObj.preferences = {
                  format: 'DD/MM/YYYY',
                };
              }
              
              return fieldObj;
            }),
          },
        ],
        external_id: opportunityId,
      };

      this.logger.log(`Creating disclaimer template with ${disclaimerFields.length} fields`);

      // Create template
      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/pdf'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      const templateId = response.data.id.toString();
      this.logger.log(`Disclaimer template created successfully with ID: ${templateId}`);

      // Prepare field values for pre-filling
      const fieldValues: Record<string, any> = {};
      
      // Auto-fill customer name (appears in two places)
      if (customerName) {
        fieldValues['Customer Name'] = customerName;
        fieldValues['cusomterName-sign'] = customerName;
      }
      
      // Auto-fill installer name
      if (installerName) {
        fieldValues['Installers Name'] = installerName;
      }
      
      // Auto-fill date
      fieldValues['Date Signed'] = currentDate;
      
      // Pre-fill form data if provided
      if (formData) {
        if (formData.unitRate) {
          fieldValues['Unit rate figures p per kWh'] = formData.unitRate;
        }
        if (formData.annualGridConsumption) {
          fieldValues['Annual Grid Consumption - figure 2 kWh'] = formData.annualGridConsumption;
        }
        if (formData.annualElectricitySpend) {
          fieldValues['Annual Electricity Spend - amount pounds'] = formData.annualElectricitySpend;
        }
        if (formData.standingCharge) {
          fieldValues['Standing Charge - per day'] = formData.standingCharge;
        }
        if (formData.utilityBillReason) {
          fieldValues['Utility Bill Reason'] = formData.utilityBillReason;
        }
      }

      // Create submission and send email with pre-filled field values
      this.logger.log(`Creating submission and sending email to ${customerData.email}`);
      const submitters = await this.createSubmissionFromTemplate(
        templateId,
        `Disclaimer - ${customerName}`,
        [
          {
            email: customerData.email,
            name: customerData.name,
            role: 'Signer1',
          },
        ],
        opportunityId,
        fieldValues
      );

      if (!submitters || submitters.length === 0) {
        throw new Error('No submitters returned from submission creation');
      }

      const firstSubmitter = submitters[0];
      
      // Build signing URL
      let signingUrl: string;
      if ((firstSubmitter as any)?.embed_src) {
        signingUrl = (firstSubmitter as any).embed_src;
      } else if (firstSubmitter.slug) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.slug}`;
      } else if (firstSubmitter.uuid) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.uuid}`;
      } else {
        signingUrl = 'Email sent to customer';
      }

      const submissionId = firstSubmitter.submission_id?.toString() || 'unknown';

      // Save submission ID to database for tracking
      try {
        await this.saveSubmissionToDatabase(
          opportunityId,
          'DISCLAIMER',
          templateId,
          submissionId,
          signingUrl,
          customerName,
          customerData.email
        );
        this.logger.log(`Saved disclaimer submission to database: ${submissionId}`);
      } catch (dbError) {
        this.logger.warn(`Failed to save submission to database: ${dbError.message}. Continuing anyway.`);
      }

      this.logger.log(`Disclaimer signing workflow created successfully for opportunity: ${opportunityId}`);
      this.logger.log(`Email sent to: ${customerData.email}`);
      this.logger.log(`Template ID: ${templateId}`);
      this.logger.log(`Submission ID: ${submissionId}`);
      this.logger.log(`Signing URL: ${signingUrl}`);

      return {
        templateId,
        submissionId,
        signingUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to create disclaimer signing workflow: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Create a template with signature fields for disclaimer
   * WITHOUT creating a submission - allows verification of field positions before sending
   * Returns template ID and preview URL for checking field positions
   */
  async createDisclaimerTemplate(
    disclaimerPdfPath: string,
    opportunityId: string,
    customerName: string,
    installerName?: string,
    formData?: {
      unitRate?: string;
      annualGridConsumption?: string;
      annualElectricitySpend?: string;
      standingCharge?: string;
      utilityBillReason?: string;
    }
  ): Promise<{
    templateId: string;
    templateSlug: string | null;
    formBuilderToken: string;
    embeddedFormUrl: string | null;
    previewUrl: string;
    fields: Array<{
      name: string;
      type: 'signature' | 'text' | 'date' | 'checkbox';
      page: number;
      x: number;
      y: number;
      w: number;
      h: number;
      required: boolean;
    }>;
  }> {
    try {
      this.logger.log(`Creating template (without submission) for disclaimer - opportunity: ${opportunityId}`);

      // 1. Convert PDF to base64
      const disclaimerBase64 = this.convertPdfToBase64(disclaimerPdfPath);

      // 2. Get current date for pre-filling
      const currentDate = new Date().toLocaleDateString('en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });

      // 3. Disclaimer fields with correct coordinates (DocuSeal uses 1-indexed pages, so page 1 = first page)
      const disclaimerFields = [
        {
          name: 'Checkbox 1',
          type: 'checkbox' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.1229268292682927, y: 0.3898665746491833, w: 0.03333333333333333, h: 0.02357948010121923 },
          ],
        },
        {
          name: 'Checkbox 2',
          type: 'checkbox' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.1229268292682927, y: 0.4492178513917645, w: 0.03333333333333333, h: 0.02357948010121923 },
          ],
        },
        {
          name: 'Installers Name',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.2461787823932927, y: 0.1555095327731409, w: 0.2308943883384146, h: 0.01656314699792963 },
          ],
        },
        {
          name: 'Customer Name',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.6390243902439025, y: 0.1555095327731409, w: 0.2328455483041159, h: 0.01518288474810214 },
          ],
        },
        {
          name: 'Unit rate figures p per kWh',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.6286128048780488, y: 0.2392403289608416, w: 0.09626030154344511, h: 0.01472280435186982 },
          ],
        },
        {
          name: 'Annual Grid Consumption - figure 2 kWh',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.2448780487804878, y: 0.409477828863656, w: 0.1001625321551067, h: 0.01426263971111541 },
          ],
        },
        {
          name: 'Annual Electricity Spend - amount pounds',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.2955792682926829, y: 0.4701919656811494, w: 0.08781927293794833, h: 0.01357604941217505 },
          ],
        },
        {
          name: 'Standing Charge - per day',
          type: 'text' as const,
          role: 'Signer1',
          required: false,
          areas: [
            { page: 1, x: 0.6559146341463414, y: 0.4681224872543842, w: 0.09919729250210818, h: 0.01702749913457491 },
          ],
        },
        {
          name: 'Utility Bill Reason',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1207164634146341, y: 0.6479946538473975, w: 0.7562204649390244, h: 0.03244189149844723 },
          ],
        },
        {
          name: 'cusomterName-sign',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1735677132955412, y: 0.8027460007440477, w: 0.298101494021532, h: 0.02009990052406829 },
          ],
        },
        {
          name: 'Date Signed',
          type: 'date' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1655917432831555, y: 0.8341300337786836, w: 0.297104521960747, h: 0.02009990052406829 },
          ],
        },
        {
          name: 'Signature',
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            { page: 1, x: 0.1810451600609756, y: 0.8665719252771308, w: 0.2861375762195122, h: 0.01904204206025706 },
          ],
        },
      ];

      // 4. Prepare request body
      const requestBody = {
        name: `Disclaimer - ${customerName} - ${opportunityId}`,
        documents: [
          {
            name: `Disclaimer - ${customerName}`,
            file: disclaimerBase64,
            fields: disclaimerFields.map(field => {
              const fieldObj: any = {
                name: field.name,
                type: field.type,
                role: field.role,
                required: field.required,
                areas: field.areas.map(area => ({
                  page: area.page,
                  x: area.x,
                  y: area.y,
                  w: area.w,
                  h: area.h,
                })),
              };

              // Add date format preference for date fields
              if (field.type === 'date') {
                fieldObj.preferences = {
                  format: 'DD/MM/YYYY',
                };
              }
              
              return fieldObj;
            }),
          },
        ],
        external_id: opportunityId,
      };

      // 5. Log API key preview for debugging
      const apiKeyPreview = this.apiKey.length > 8 
        ? `${this.apiKey.substring(0, 4)}...${this.apiKey.substring(this.apiKey.length - 4)}`
        : '***';
      this.logger.log(`Using API key: ${apiKeyPreview} (length: ${this.apiKey.length}) for disclaimer template creation`);

      // 6. Create template
      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/pdf'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      const templateId = response.data.id.toString();
      const templateSlug = response.data.slug || response.data.template_slug || null;
      
      this.logger.log(`Template created successfully with ID: ${templateId}`);
      if (templateSlug) {
        this.logger.log(`Template slug: ${templateSlug}`);
      }

      // 7. Get template slug if not in initial response
      let finalTemplateSlug = templateSlug;
      if (!finalTemplateSlug) {
        try {
          const templateDetailsResponse = await axios.get<DocuSealTemplate>(
            this.getApiPath(`/templates/${templateId}`),
            {
              headers: {
                'X-Auth-Token': this.apiKey,
              },
            }
          );
          finalTemplateSlug = templateDetailsResponse.data.slug || templateDetailsResponse.data.template_slug || null;
          if (finalTemplateSlug) {
            this.logger.log(`Retrieved template slug from API: ${finalTemplateSlug}`);
          }
        } catch (error) {
          this.logger.warn(`Could not fetch template slug: ${error.message}`);
        }
      }

      // 8. Generate JWT token for Form Builder
      const userEmail = process.env.DOCUSEAL_USER_EMAIL || 'operations@creativuk.co.uk';
      
      const jwtPayload = {
        user_email: userEmail,
        integration_email: userEmail,
        external_id: `template-${opportunityId}`,
        name: `Disclaimer - ${customerName}`,
        template_id: parseInt(templateId, 10),
      };
      
      const formBuilderToken = jwt.sign(jwtPayload, this.apiKey, {
        algorithm: 'HS256',
      });
      
      this.logger.log(`JWT token generated for Form Builder (template ID: ${templateId})`);

      // 9. Build embedded form URL
      // Fix URL generation: https://api.docuseal.eu -> https://docuseal.eu
      let baseUrlForEmbed: string;
      if (this.baseUrl.includes('/api')) {
        // Self-hosted: remove /api
        baseUrlForEmbed = this.baseUrl.replace('/api', '');
      } else if (this.baseUrl.includes('api.docuseal.eu')) {
        // EU cloud: replace api.docuseal.eu with docuseal.eu
        baseUrlForEmbed = this.baseUrl.replace('api.docuseal.eu', 'docuseal.eu');
      } else if (this.baseUrl.includes('api.docuseal.com')) {
        // US cloud: replace api.docuseal.com with docuseal.com
        baseUrlForEmbed = this.baseUrl.replace('api.docuseal.com', 'docuseal.com');
      } else {
        // Fallback: try regex replacement
        baseUrlForEmbed = this.baseUrl.replace(/:\/\/api\./, '://');
      }
      
      const embeddedFormUrl = finalTemplateSlug 
        ? `${baseUrlForEmbed}/d/${finalTemplateSlug}`
        : null;

      // 10. Build preview URL (fallback)
      const previewUrl = finalTemplateSlug 
        ? `${baseUrlForEmbed}/d/${finalTemplateSlug}?token=${formBuilderToken}`
        : `${this.baseUrl}/templates/${templateId}`;

      // 11. Flatten fields for response
      const flattenedFields = disclaimerFields.flatMap(field =>
        field.areas.map(area => ({
          name: field.name,
          type: field.type,
          page: area.page + 1, // Convert to 1-indexed for display
          x: area.x,
          y: area.y,
          w: area.w,
          h: area.h,
          required: field.required,
        }))
      );

      this.logger.log(`Template created for verification. Template ID: ${templateId}`);
      this.logger.log(`Template Slug: ${finalTemplateSlug}`);
      this.logger.log(`Form Builder Token: Generated (${formBuilderToken.length} characters)`);
      this.logger.log(`Embedded Form URL: ${embeddedFormUrl || 'N/A'}`);

      return {
        templateId,
        templateSlug: finalTemplateSlug,
        formBuilderToken,
        embeddedFormUrl,
        previewUrl,
        fields: flattenedFields,
      };
    } catch (error) {
      this.logger.error(`Failed to create disclaimer template: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Create a signing workflow for Booking Confirmation only
   * Sends booking confirmation document to customer for signing
   * Includes signature, customer name, and date fields
   */
  async createBookingConfirmationSigningWorkflow(
    bookConfirmationPdfPath: string,
    opportunityId: string,
    customerData: {
      name: string;
      email: string;
    },
    customerName: string
  ): Promise<{
    templateId: string;
    submissionId: string;
    signingUrl: string;
  }> {
    try {
      this.logger.log(`Creating booking confirmation signing workflow for opportunity: ${opportunityId}`);

      // Convert PDF to base64
      const bookConfirmationBase64 = this.convertPdfToBase64(bookConfirmationPdfPath);

      // Book Confirmation: page 1 (exact coordinates from DocuSeal template)
      // Page 0 in DocuSeal API = Page 1 in our system (1-indexed)
      // Includes signature, customer name, and date fields
      const bookConfirmationFields = [
        {
          name: 'Signature',
          type: 'signature' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: 1, // Page 0 in DocuSeal = Page 1 here
              x: 0.3638886142240891,
              y: 0.6804700876227533,
              w: 0.2926829268292683,
              h: 0.03519668737060044,
            },
          ],
        },
        {
          name: 'Full Name',
          type: 'text' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: 1,
              x: 0.3873007373062174,
              y: 0.7563829362475439,
              w: 0.2829268292682927,
              h: 0.03174603174603174,
            },
          ],
        },
        {
          name: 'Date Signed',
          type: 'date' as const,
          role: 'Signer1',
          required: true,
          areas: [
            {
              page: 1,
              x: 0.1921951219512195,
              y: 0.8313780179757592,
              w: 0.2809756990758384,
              h: 0.03036576949620429,
            },
          ],
        },
      ];

      // Prepare request body with exact coordinates and required fields
      const requestBody = {
        name: `Booking Confirmation - ${customerName} - ${opportunityId}`,
        documents: [
          {
            name: `Booking Confirmation - ${customerName}`,
            file: bookConfirmationBase64,
            fields: bookConfirmationFields.map(field => {
              const fieldObj: any = {
                name: field.name,
                type: field.type,
                role: field.role,
                required: field.required !== undefined ? field.required : true,
                areas: field.areas.map(area => ({
                  page: area.page,
                  x: area.x,
                  y: area.y,
                  w: area.w,
                  h: area.h,
                })),
              };
              
              // Add date format preference for date fields
              if (field.type === 'date') {
                fieldObj.preferences = {
                  format: 'DD/MM/YYYY',
                };
              }
              
              return fieldObj;
            }),
          },
        ],
        external_id: opportunityId,
      };

      this.logger.log(`Creating booking confirmation template`);

      // Create template
      const response = await axios.post<DocuSealTemplate>(
        this.getApiPath('/templates/pdf'),
        requestBody,
        {
          headers: {
            'X-Auth-Token': this.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      const templateId = response.data.id.toString();
      this.logger.log(`Booking confirmation template created successfully with ID: ${templateId}`);

      // Get current date for pre-filling
      const currentDate = new Date().toLocaleDateString('en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });

      // Create submission and send email with pre-filled field values
      this.logger.log(`Creating submission and sending email to ${customerData.email}`);
      const submitters = await this.createSubmissionFromTemplate(
        templateId,
        `Booking Confirmation - ${customerName}`,
        [
          {
            email: customerData.email,
            name: customerData.name,
            role: 'Signer1',
          },
        ],
        opportunityId,
        {
          "Full Name": customerData.name,
          "Date Signed": currentDate
        }
      );

      if (!submitters || submitters.length === 0) {
        throw new Error('No submitters returned from submission creation');
      }

      const firstSubmitter = submitters[0];
      
      // Build signing URL
      let signingUrl: string;
      if ((firstSubmitter as any)?.embed_src) {
        signingUrl = (firstSubmitter as any).embed_src;
      } else if (firstSubmitter.slug) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.slug}`;
      } else if (firstSubmitter.uuid) {
        signingUrl = `${this.baseUrl}/s/${firstSubmitter.uuid}`;
      } else {
        signingUrl = 'Email sent to customer';
      }

      const submissionId = firstSubmitter.submission_id?.toString() || 'unknown';

      // Save submission ID to database for tracking
      try {
        await this.saveSubmissionToDatabase(
          opportunityId,
          'BOOKING_CONFIRMATION',
          templateId,
          submissionId,
          signingUrl,
          customerData.name,
          customerData.email
        );
        this.logger.log(`Saved booking confirmation submission to database: ${submissionId}`);
      } catch (dbError) {
        this.logger.warn(`Failed to save submission to database: ${dbError.message}. Continuing anyway.`);
      }

      this.logger.log(`Booking confirmation signing workflow created successfully for opportunity: ${opportunityId}`);
      this.logger.log(`Email sent to: ${customerData.email}`);
      this.logger.log(`Template ID: ${templateId}`);
      this.logger.log(`Submission ID: ${submissionId}`);
      this.logger.log(`Signing URL: ${signingUrl}`);

      return {
        templateId,
        submissionId,
        signingUrl,
      };
    } catch (error) {
      this.logger.error(`Failed to create booking confirmation signing workflow: ${error.message}`);
      if (error.response) {
        this.logger.error(`Response status: ${error.response.status}`);
        this.logger.error(`Response data: ${JSON.stringify(error.response.data)}`);
      }
      throw error;
    }
  }

  /**
   * Save submission ID to database for tracking
   */
  private async saveSubmissionToDatabase(
    opportunityId: string,
    documentType: 'CONTRACT' | 'DISCLAIMER' | 'BOOKING_CONFIRMATION',
    templateId: string,
    submissionId: string,
    signingUrl: string,
    customerName?: string,
    customerEmail?: string
  ): Promise<void> {
    try {
      await this.prisma.docuSealSubmission.upsert({
        where: {
          opportunityId_documentType: {
            opportunityId,
            documentType,
          },
        },
        update: {
          templateId,
          submissionId,
          signingUrl,
          customerName: customerName || undefined,
          customerEmail: customerEmail || undefined,
          updatedAt: new Date(),
        },
        create: {
          opportunityId,
          documentType,
          templateId,
          submissionId,
          signingUrl,
          status: 'pending',
          customerName: customerName || undefined,
          customerEmail: customerEmail || undefined,
        },
      });
    } catch (error) {
      this.logger.error(`Error saving submission to database: ${error.message}`);
      throw error;
    }
  }

  /**
   * Sync submission status from DocuSeal API to database
   * Fetches fresh status from DocuSeal and updates the database
   * If completed, downloads signed document and audit log and uploads to OneDrive
   */
  async syncSubmissionStatusFromDocuSeal(submissionId: string): Promise<void> {
    try {
      // Get fresh status from DocuSeal API
      const docuSealSubmission = await this.getSubmissionStatus(submissionId);
      
      // Map DocuSeal status to our database status (lowercase to match schema default)
      let dbStatus = 'pending';
      if (docuSealSubmission.status === 'completed' || docuSealSubmission.status === 'approved') {
        dbStatus = 'completed';
      } else if (docuSealSubmission.status === 'declined' || docuSealSubmission.status === 'rejected') {
        dbStatus = 'declined';
      } else if (docuSealSubmission.status === 'expired') {
        dbStatus = 'expired';
      }

      // Get signed document URL if completed
      let signedDocumentUrl: string | undefined;
      if (dbStatus === 'completed') {
        try {
          const documents = await this.getSubmissionDocuments(submissionId);
          if (documents.documents && documents.documents.length > 0) {
            // Get the first document URL (usually the signed PDF)
            signedDocumentUrl = documents.documents[0].url;
          }
        } catch (docError) {
          this.logger.warn(`Failed to get signed document URL for submission ${submissionId}: ${docError.message}`);
        }
      }

      // Find the submission by submissionId (not unique, so use findFirst)
      const dbSubmission = await this.prisma.docuSealSubmission.findFirst({
        where: { submissionId },
      });

      if (!dbSubmission) {
        this.logger.warn(`Submission ${submissionId} not found in database, skipping sync`);
        return;
      }

      // Check if status changed to completed (wasn't completed before)
      const wasCompleted = dbSubmission.status === 'completed';
      const isNowCompleted = dbStatus === 'completed';

      // Update database with fresh status using the unique id
      await this.prisma.docuSealSubmission.update({
        where: { id: dbSubmission.id },
        data: {
          status: dbStatus,
          signedDocumentUrl: signedDocumentUrl || undefined,
          completedAt: dbStatus === 'completed' && docuSealSubmission.completed_at 
            ? new Date(docuSealSubmission.completed_at) 
            : undefined,
          updatedAt: new Date(),
        },
      });

      this.logger.log(`Synced submission ${submissionId} status: ${dbStatus}`);
    } catch (error) {
      this.logger.error(`Error syncing submission status from DocuSeal: ${error.message}`);
      // Don't throw - we'll still return database status if sync fails
    }
  }

  /**
   * Get submission IDs for an opportunity
   * Returns all document types (contract, disclaimer, booking_confirmation)
   * Automatically syncs status from DocuSeal API to get fresh status
   */
  async getSubmissionsByOpportunity(opportunityId: string): Promise<{
    contract?: { submissionId: string; templateId: string; status: string; signingUrl?: string };
    disclaimer?: { submissionId: string; templateId: string; status: string; signingUrl?: string };
    bookingConfirmation?: { submissionId: string; templateId: string; status: string; signingUrl?: string };
  }> {
    try {
      const submissions = await this.prisma.docuSealSubmission.findMany({
        where: { opportunityId },
      });

      const result: any = {};
      
      // Sync status from DocuSeal API for each submission
      for (const submission of submissions) {
        // Fetch fresh status from DocuSeal API and update database
        await this.syncSubmissionStatusFromDocuSeal(submission.submissionId);

        // Re-fetch from database to get updated status (use id since we already have it)
        const updatedSubmission = await this.prisma.docuSealSubmission.findUnique({
          where: { id: submission.id },
        });

        if (!updatedSubmission) {
          this.logger.warn(`Submission ${submission.submissionId} not found after sync, skipping`);
          continue;
        }

        const submissionData = {
          submissionId: updatedSubmission.submissionId,
          templateId: updatedSubmission.templateId,
          status: updatedSubmission.status.toLowerCase(), // Return lowercase for consistency
          signingUrl: updatedSubmission.signingUrl || undefined,
        };

        switch (updatedSubmission.documentType) {
          case 'CONTRACT':
            result.contract = submissionData;
            break;
          case 'DISCLAIMER':
            result.disclaimer = submissionData;
            break;
          case 'BOOKING_CONFIRMATION':
            result.bookingConfirmation = submissionData;
            break;
        }
      }

      return result;
    } catch (error) {
      this.logger.error(`Error getting submissions by opportunity: ${error.message}`);
      throw error;
    }
  }

  /**
   * Update submission status (e.g., when document is completed)
   */
  async updateSubmissionStatus(
    opportunityId: string,
    documentType: 'CONTRACT' | 'DISCLAIMER' | 'BOOKING_CONFIRMATION',
    status: string,
    signedDocumentUrl?: string
  ): Promise<void> {
    try {
      await this.prisma.docuSealSubmission.update({
        where: {
          opportunityId_documentType: {
            opportunityId,
            documentType,
          },
        },
        data: {
          status,
          signedDocumentUrl: signedDocumentUrl || undefined,
          completedAt: status === 'completed' ? new Date() : undefined,
          updatedAt: new Date(),
        },
      });
    } catch (error) {
      this.logger.error(`Error updating submission status: ${error.message}`);
      throw error;
    }
  }

  /**
   * Refresh status for a specific submission
   * Fetches fresh status from DocuSeal API and updates database
   */
  async refreshSubmissionStatus(submissionId: string): Promise<{
    submissionId: string;
    status: string;
    signedDocumentUrl?: string;
    completedAt?: Date;
  }> {
    try {
      await this.syncSubmissionStatusFromDocuSeal(submissionId);
      
      // Find by submissionId (not unique, so use findFirst)
      const submission = await this.prisma.docuSealSubmission.findFirst({
        where: { submissionId },
      });

      if (!submission) {
        throw new Error(`Submission ${submissionId} not found in database`);
      }

      return {
        submissionId: submission.submissionId,
        status: submission.status.toLowerCase(),
        signedDocumentUrl: submission.signedDocumentUrl || undefined,
        completedAt: submission.completedAt || undefined,
      };
    } catch (error) {
      this.logger.error(`Error refreshing submission status: ${error.message}`);
      throw error;
    }
  }

  /**
   * List all submissions from DocuSeal API
   * Can filter by status, template_id, or search query
   */
  async listAllSubmissions(params?: {
    status?: string;
    template_id?: number;
    q?: string;
    limit?: number;
    after?: number;
  }): Promise<{
    data: DocuSealSubmission[];
    pagination?: {
      next?: number;
      before?: number;
    };
  }> {
    try {
      const queryParams = new URLSearchParams();
      if (params?.status) queryParams.append('status', params.status);
      if (params?.template_id) queryParams.append('template_id', params.template_id.toString());
      if (params?.q) queryParams.append('q', params.q);
      if (params?.limit) queryParams.append('limit', params.limit.toString());
      if (params?.after) queryParams.append('after', params.after.toString());

      const queryString = queryParams.toString();
      const url = this.getApiPath(`/submissions${queryString ? `?${queryString}` : ''}`);

      const response = await axios.get<{
        data: DocuSealSubmission[];
        pagination?: {
          next?: number;
          before?: number;
        };
      }>(url, {
        headers: {
          'X-Auth-Token': this.apiKey,
        },
      });

      return response.data;
    } catch (error) {
      this.logger.error(`Failed to list submissions: ${error.message}`);
      throw error;
    }
  }

  /**
   * Extract opportunity ID from template name or external_id
   * Template names follow pattern: "{DocumentType} - {CustomerName} - {OpportunityId}"
   * Examples:
   * - "Contract & Booking Confirmation Template - Test 1 - ZXVh7ONnuHGMpIHnpKZM"
   * - "Disclaimer - Test 1 - ZXVh7ONnuHGMpIHnpKZM"
   */
  private extractOpportunityIdFromTemplate(template: { name?: string; external_id?: string }): string | null {
    try {
      // First try external_id
      if (template.external_id) {
        // external_id might be in format like "template-{opportunityId}" or just "{opportunityId}"
        const match = template.external_id.match(/(?:template-)?([A-Za-z0-9_-]+)/);
        if (match && match[1]) {
          this.logger.debug(`Extracted opportunity ID from external_id: ${match[1]}`);
          return match[1];
        }
      }

      // Try to extract from template name
      // Pattern: "{DocumentType} - {CustomerName} - {OpportunityId}"
      // The opportunity ID is always the last part after the last " - "
      if (template.name) {
        const parts = template.name.split(' - ');
        if (parts.length >= 2) {
          const lastPart = parts[parts.length - 1].trim();
          // GHL opportunity IDs are typically 20+ characters alphanumeric
          // Check if it looks like an opportunity ID (at least 15 chars, alphanumeric with possible dashes/underscores)
          if (lastPart.length >= 15 && /^[A-Za-z0-9_-]+$/.test(lastPart)) {
            this.logger.debug(`Extracted opportunity ID from template name (last part): ${lastPart}`);
            return lastPart;
          }
        }

        // Fallback: Try to find opportunity ID pattern anywhere in the name
        // Look for a long alphanumeric string (GHL opportunity IDs are typically 20+ chars)
        const opportunityIdMatch = template.name.match(/([A-Za-z0-9_-]{20,})/);
        if (opportunityIdMatch && opportunityIdMatch[1]) {
          this.logger.debug(`Extracted opportunity ID from template name (pattern match): ${opportunityIdMatch[1]}`);
          return opportunityIdMatch[1];
        }
      }

      this.logger.warn(`Could not extract opportunity ID from template: ${JSON.stringify({ name: template.name, external_id: template.external_id })}`);
      return null;
    } catch (error) {
      this.logger.warn(`Error extracting opportunity ID from template: ${error.message}`);
      return null;
    }
  }

  /**
   * Get template details from DocuSeal
   */
  async getTemplate(templateId: string): Promise<{
    id: number;
    name: string;
    external_id?: string;
    [key: string]: any;
  }> {
    try {
      const response = await axios.get(
        this.getApiPath(`/templates/${templateId}`),
        {
          headers: {
            'X-Auth-Token': this.apiKey,
          },
        }
      );

      return response.data;
    } catch (error) {
      this.logger.error(`Failed to get template: ${error.message}`);
      throw error;
    }
  }

  /**
   * Process completed submissions for an opportunity and upload to OneDrive
   * Checks all submissions in DocuSeal, filters by opportunity ID, and uploads if signed
   */
  async processCompletedSubmissionsForOpportunity(opportunityId: string): Promise<{
    success: boolean;
    message: string;
    processed: number;
    errors?: string[];
  }> {
    try {
      this.logger.log(`🔍 Processing completed submissions for opportunity: ${opportunityId}`);

      // First, check database for submissions for this opportunity
      const dbSubmissions = await this.prisma.docuSealSubmission.findMany({
        where: { opportunityId },
      });

      this.logger.log(`Found ${dbSubmissions.length} submission(s) in database for opportunity ${opportunityId}`);

      // Get all completed submissions from DocuSeal
      const allSubmissions = await this.listAllSubmissions({
        status: 'completed',
        limit: 100, // Get up to 100 completed submissions
      });

      const errors: string[] = [];
      let processed = 0;

      this.logger.log(`Found ${allSubmissions.data.length} completed submission(s) from DocuSeal API`);

      // Primary approach: Use database submissions (more reliable since we know the opportunity ID)
      // Check each database submission and verify if it's completed in DocuSeal
      if (dbSubmissions.length > 0) {
        this.logger.log(`Processing ${dbSubmissions.length} database submission(s) and verifying status in DocuSeal`);
        
        for (const dbSubmission of dbSubmissions) {
          try {
            // Get fresh status from DocuSeal (always check, even if marked completed in DB)
            const docuSealSubmission = await this.getSubmissionStatus(dbSubmission.submissionId);
            
            if (docuSealSubmission.status === 'completed' || docuSealSubmission.status === 'approved') {
              this.logger.log(`📄 Found completed submission ${dbSubmission.submissionId} in database, processing...`);
              
              // Get template to determine document type
              const template = await this.getTemplate(dbSubmission.templateId);
              
              // Determine document type from template name
              let documentType: 'CONTRACT' | 'DISCLAIMER' | 'BOOKING_CONFIRMATION' = dbSubmission.documentType;
              const templateName = template.name?.toLowerCase() || '';
              if (templateName.includes('disclaimer') || templateName.includes('epvs')) {
                documentType = 'DISCLAIMER';
              } else if (templateName.includes('booking') || templateName.includes('confirmation')) {
                documentType = 'BOOKING_CONFIRMATION';
              }

              // Download signed document
              const signedDocumentBuffer = await this.getSignedDocument(dbSubmission.submissionId);
              this.logger.log(`✅ Downloaded signed document for submission ${dbSubmission.submissionId}`);

              // Download audit log if available
              let auditLogBuffer: Buffer | null = null;
              if (docuSealSubmission.audit_log_url) {
                try {
                  auditLogBuffer = await this.getAuditLog(docuSealSubmission.audit_log_url);
                  this.logger.log(`✅ Downloaded audit log for submission ${dbSubmission.submissionId}`);
                } catch (auditError) {
                  this.logger.warn(`Failed to download audit log: ${auditError.message}`);
                }
              }

              // Get customer name
              const customerName = dbSubmission.customerName || 
                                   docuSealSubmission.submitters?.[0]?.name || 
                                   'Customer';

              // Upload to OneDrive
              const uploadResult = await this.oneDriveFileManagerService.uploadSignedDocumentAndAuditLogToOrders(
                opportunityId,
                customerName,
                signedDocumentBuffer,
                auditLogBuffer,
                documentType,
                dbSubmission.submissionId
              );

              if (uploadResult.success) {
                processed++;
                this.logger.log(`✅ Successfully uploaded submission ${dbSubmission.submissionId} to OneDrive`);
                
                // Update database status
                await this.updateSubmissionStatus(
                  opportunityId,
                  documentType,
                  'completed',
                  docuSealSubmission.combined_document_url
                );
              } else {
                errors.push(`Failed to upload submission ${dbSubmission.submissionId}: ${uploadResult.error}`);
              }
            } else {
              this.logger.debug(`Submission ${dbSubmission.submissionId} status is ${docuSealSubmission.status}, not completed yet`);
            }
          } catch (error) {
            const errorMsg = `Error processing database submission ${dbSubmission.submissionId}: ${error.message}`;
            this.logger.error(errorMsg);
            errors.push(errorMsg);
          }
        }
      }

      // Fallback: If no database submissions found, try API approach (filter by template name)
      if (processed === 0 && allSubmissions.data.length > 0 && dbSubmissions.length === 0) {
        this.logger.log(`No database submissions found, trying API approach with ${allSubmissions.data.length} completed submission(s)`);
        
        // Filter submissions by opportunity ID from API results
        for (const submission of allSubmissions.data) {
          try {
          // Get template to extract opportunity ID
          if (!submission.template_id) {
            this.logger.debug(`Skipping submission ${submission.id} - no template_id`);
            continue;
          }

          const template = await this.getTemplate(submission.template_id.toString());
          this.logger.debug(`Template for submission ${submission.id}: name="${template.name}", external_id="${template.external_id}"`);
          
          const extractedOpportunityId = this.extractOpportunityIdFromTemplate(template);

          // Skip if opportunity ID doesn't match
          if (!extractedOpportunityId) {
            this.logger.debug(`Skipping submission ${submission.id} - could not extract opportunity ID`);
            continue;
          }

          if (extractedOpportunityId !== opportunityId) {
            this.logger.debug(`Skipping submission ${submission.id} - opportunity ID mismatch: ${extractedOpportunityId} !== ${opportunityId}`);
            continue;
          }

          this.logger.log(`✅ Found matching submission ${submission.id} for opportunity ${opportunityId}`);

          // Check if we've already uploaded this submission (to avoid duplicates)
          const alreadyUploaded = dbSubmissions.some(
            dbSub => dbSub.submissionId === submission.id.toString() && dbSub.status === 'completed'
          );

          if (alreadyUploaded) {
            this.logger.log(`⏭️ Submission ${submission.id} already processed, skipping`);
            continue;
          }

          // Only process if status is completed
          if (submission.status === 'completed' || submission.status === 'approved') {
            this.logger.log(`📄 Processing completed submission ${submission.id} for opportunity ${opportunityId}`);

            // Determine document type from template name
            let documentType: 'CONTRACT' | 'DISCLAIMER' | 'BOOKING_CONFIRMATION' = 'CONTRACT';
            const templateName = template.name?.toLowerCase() || '';
            if (templateName.includes('disclaimer') || templateName.includes('epvs')) {
              documentType = 'DISCLAIMER';
            } else if (templateName.includes('booking') || templateName.includes('confirmation')) {
              documentType = 'BOOKING_CONFIRMATION';
            }

            // Download signed document
            const signedDocumentBuffer = await this.getSignedDocument(submission.id.toString());
            this.logger.log(`✅ Downloaded signed document for submission ${submission.id}`);

            // Download audit log if available
            let auditLogBuffer: Buffer | null = null;
            if (submission.audit_log_url) {
              try {
                auditLogBuffer = await this.getAuditLog(submission.audit_log_url);
                this.logger.log(`✅ Downloaded audit log for submission ${submission.id}`);
              } catch (auditError) {
                this.logger.warn(`Failed to download audit log: ${auditError.message}`);
              }
            }

            // Get customer name from submission or template
            const customerName = submission.submitters?.[0]?.name || 
                                 template.name?.split(' - ')[1]?.trim() || 
                                 'Customer';

            // Upload to OneDrive
            const uploadResult = await this.oneDriveFileManagerService.uploadSignedDocumentAndAuditLogToOrders(
              opportunityId,
              customerName,
              signedDocumentBuffer,
              auditLogBuffer,
              documentType,
              submission.id.toString()
            );

            if (uploadResult.success) {
              processed++;
              this.logger.log(`✅ Successfully uploaded submission ${submission.id} to OneDrive`);
            } else {
              errors.push(`Failed to upload submission ${submission.id}: ${uploadResult.error}`);
            }
          }
        } catch (error) {
          const errorMsg = `Error processing submission ${submission.id}: ${error.message}`;
          this.logger.error(errorMsg);
          errors.push(errorMsg);
          }
        }
      }

      return {
        success: errors.length === 0,
        message: `Processed ${processed} completed submission(s) for opportunity ${opportunityId}`,
        processed,
        errors: errors.length > 0 ? errors : undefined,
      };
    } catch (error) {
      this.logger.error(`Error processing completed submissions: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get DocuSeal configuration for debugging
   * Returns information about the loaded API key and base URL
   */
  public getDocuSealConfig() {
    const apiKeySet = !!this.apiKey;
    const apiKeyPreview = apiKeySet ? `${this.apiKey.substring(0, 4)}...${this.apiKey.substring(this.apiKey.length - 4)}` : 'Not set';
    
    // Also check what's in process.env for comparison
    const envApiKey = process.env.DOCUSEAL_API_KEY;
    const envApiKeyPreview = envApiKey ? `${envApiKey.substring(0, 4)}...${envApiKey.substring(envApiKey.length - 4)}` : 'Not set';
    const keysMatch = this.apiKey === envApiKey;
    
    return {
      apiKeySet,
      apiKeyPreview,
      apiKeyLength: apiKeySet ? this.apiKey.length : 0,
      baseUrl: this.baseUrl,
      envApiKeySet: !!envApiKey,
      envApiKeyPreview,
      envApiKeyLength: envApiKey ? envApiKey.length : 0,
      keysMatch,
      note: 'If you changed the API key in .env, restart the server for changes to take effect. If keysMatch is false, the service is using a cached key from startup.',
    };
  }

  /**
   * Verify webhook secret from request headers
   * DocuSeal includes the secret key-value pair in request headers when configured
   * @param headers Request headers from webhook
   * @param payload Webhook payload (for additional verification if needed)
   * @returns true if secret is verified or not configured, false if verification fails
   */
  async verifyWebhookSecret(headers: Record<string, string>, payload: any): Promise<boolean> {
    try {
      const webhookSecret = process.env.DOCUSEAL_WEBHOOK_SECRET;
      
      // If no secret is configured, skip verification (allow webhook)
      if (!webhookSecret) {
        this.logger.debug('No DOCUSEAL_WEBHOOK_SECRET configured - skipping verification');
        return true;
      }

      // DocuSeal sends the secret in headers
      // The exact header name depends on how you configured it in DocuSeal console
      // Common patterns: X-Webhook-Secret, X-DocuSeal-Secret, or custom header name
      const secretHeader = headers['x-webhook-secret'] || 
                          headers['x-docuseal-secret'] || 
                          headers['x-secret'] ||
                          headers['webhook-secret'];

      if (!secretHeader) {
        this.logger.warn('Webhook secret header not found in request');
        return false;
      }

      // Compare secrets (use constant-time comparison in production)
      const isValid = secretHeader === webhookSecret;
      
      if (!isValid) {
        this.logger.warn('Webhook secret verification failed - secret mismatch');
      } else {
        this.logger.debug('✅ Webhook secret verified successfully');
      }

      return isValid;
    } catch (error) {
      this.logger.error(`Error verifying webhook secret: ${error.message}`);
      return false;
    }
  }

  /**
   * Handle webhook events from DocuSeal
   * Processes different event types and updates the database accordingly
   * 
   * DocuSeal webhook payload structure:
   * {
   *   "event_type": "form.completed" | "submission.completed" | etc.,
   *   "data": {
   *     "id": 616343,  // submitter ID
   *     "submission": { "id": 473508 },  // submission ID
   *     "template": { "external_id": "template-opportunityId" },
   *     "status": "completed",
   *     "documents": [{ "url": "..." }],
   *     "audit_log_url": "..."
   *   }
   * }
   * 
   * @param eventType The type of event (e.g., 'form.completed', 'submission.completed', 'submission.declined')
   * @param payload The full webhook payload
   * @param submissionId The submission ID from the webhook
   */
  async handleWebhookEvent(eventType: string, payload: any, submissionId: string): Promise<void> {
    try {
      this.logger.log(`🔄 Processing webhook event: ${eventType} for submission: ${submissionId}`);

      // Extract opportunity ID from template external_id if available
      // Format: "template-{opportunityId}" or just "{opportunityId}"
      const templateExternalId = payload.data?.template?.external_id || payload.template?.external_id;
      let opportunityId: string | null = null;
      
      if (templateExternalId) {
        // Remove "template-" prefix if present
        opportunityId = templateExternalId.replace(/^template-/, '');
        this.logger.log(`📋 Extracted opportunity ID from template external_id: ${opportunityId}`);
      }

      // Find the submission in database
      const dbSubmission = await this.prisma.docuSealSubmission.findFirst({
        where: { submissionId },
      });

      if (!dbSubmission) {
        this.logger.warn(`⚠️ Submission ${submissionId} not found in database`);
        
        // If we have opportunity ID from template, try to find by that
        if (opportunityId) {
          const submissionsByOpportunity = await this.prisma.docuSealSubmission.findMany({
            where: { opportunityId },
          });
          
          if (submissionsByOpportunity.length > 0) {
            this.logger.log(`📋 Found ${submissionsByOpportunity.length} submission(s) for opportunity ${opportunityId}, syncing status...`);
            // Try to sync status - it will update the submission if found
            await this.syncSubmissionStatusFromDocuSeal(submissionId);
            return;
          }
        }
        
        this.logger.warn(`⚠️ Submission ${submissionId} not found in database - may be from external system or not yet saved`);
        // Still try to sync status in case it's a new submission
        await this.syncSubmissionStatusFromDocuSeal(submissionId);
        return;
      }

      // Handle different event types
      // Note: DocuSeal sends both "form.completed" and "submission.completed" events
      switch (eventType) {
        case 'form.completed':
        case 'submission.completed':
        case 'submission.approved':
          // Check if we've already processed this submission (avoid duplicate uploads)
          // The upload method will check if files exist, but we can skip early if already completed
          if (dbSubmission.status === 'completed') {
            this.logger.log(`ℹ️ Submission ${submissionId} already marked as completed in database`);
            // Still sync status in case it changed, but skip upload (upload method will check if files exist)
            await this.syncSubmissionStatusFromDocuSeal(submissionId);
            
            // Check if we should still upload (in case files were deleted or missing)
            // The upload method will handle checking if files exist
            this.logger.log(`📋 Proceeding with upload check - upload method will skip if files already exist`);
          }
          
          this.logger.log(`✅ Submission ${submissionId} completed - updating status and uploading to OneDrive`);
          
          // Sync status first
          await this.syncSubmissionStatusFromDocuSeal(submissionId);
          
          // Get document URLs from webhook payload (more reliable than API call)
          const signedDocumentUrl = payload.data?.documents?.[0]?.url || 
                                   payload.data?.submitters?.[0]?.documents?.[0]?.url ||
                                   null;
          const auditLogUrl = payload.data?.audit_log_url || null;
          
          // Only upload if status is actually completed and we have document URL
          if (payload.data?.status === 'completed' || payload.data?.status === 'approved') {
            if (!signedDocumentUrl) {
              this.logger.warn(`⚠️ No signed document URL in webhook payload for submission ${submissionId}, trying API...`);
              // Fallback to API if URL not in payload
              try {
                const completedSubmission = await this.getSubmissionStatus(submissionId);
                if (completedSubmission.status === 'completed' || completedSubmission.status === 'approved') {
                  // Try to get from API
                  const signedDocumentBuffer = await this.getSignedDocument(submissionId);
                  this.logger.log(`✅ Downloaded signed document via API for submission ${submissionId}`);
                  
                  // Download audit log if available
                  let auditLogBuffer: Buffer | null = null;
                  if (completedSubmission.audit_log_url) {
                    try {
                      auditLogBuffer = await this.getAuditLog(completedSubmission.audit_log_url);
                      this.logger.log(`✅ Downloaded audit log for submission ${submissionId}`);
                    } catch (auditError) {
                      this.logger.warn(`Failed to download audit log: ${auditError.message}`);
                    }
                  }

                  // Get customer name
                  const customerName = dbSubmission.customerName || 
                                       completedSubmission.submitters?.[0]?.name || 
                                       'Customer';

                  // Upload to OneDrive
                  const uploadResult = await this.oneDriveFileManagerService.uploadSignedDocumentAndAuditLogToOrders(
                    dbSubmission.opportunityId,
                    customerName,
                    signedDocumentBuffer,
                    auditLogBuffer,
                    dbSubmission.documentType,
                    submissionId
                  );

                  if (uploadResult.success) {
                    this.logger.log(`✅ Successfully uploaded signed document and audit log to OneDrive for submission ${submissionId}`);
                  } else {
                    this.logger.error(`❌ Failed to upload to OneDrive for submission ${submissionId}: ${uploadResult.error}`);
                  }
                }
              } catch (apiError) {
                this.logger.error(`Failed to download via API: ${apiError.message}`);
              }
            } else {
              // Use URLs from webhook payload (preferred method)
              try {
                this.logger.log(`📥 Downloading signed document and audit log from webhook URLs for submission ${submissionId}`);
                
                // Download signed document from URL
                const signedDocumentBuffer = await this.downloadDocumentFromUrl(signedDocumentUrl);
                this.logger.log(`✅ Downloaded signed document for submission ${submissionId}`);

                // Download audit log if available
                let auditLogBuffer: Buffer | null = null;
                if (auditLogUrl) {
                  try {
                    auditLogBuffer = await this.getAuditLog(auditLogUrl);
                    this.logger.log(`✅ Downloaded audit log for submission ${submissionId}`);
                  } catch (auditError) {
                    this.logger.warn(`Failed to download audit log for submission ${submissionId}: ${auditError.message}`);
                  }
                } else {
                  this.logger.warn(`No audit log URL available for submission ${submissionId}`);
                }

                // Get customer name from database submission or payload
                const customerName = dbSubmission.customerName || 
                                     payload.data?.submitters?.[0]?.name || 
                                     'Customer';

                // Upload to OneDrive
                const uploadResult = await this.oneDriveFileManagerService.uploadSignedDocumentAndAuditLogToOrders(
                  dbSubmission.opportunityId,
                  customerName,
                  signedDocumentBuffer,
                  auditLogBuffer,
                  dbSubmission.documentType,
                  submissionId
                );

                if (uploadResult.success) {
                  this.logger.log(`✅ Successfully uploaded signed document and audit log to OneDrive for submission ${submissionId}`);
                } else {
                  this.logger.error(`❌ Failed to upload to OneDrive for submission ${submissionId}: ${uploadResult.error}`);
                }
              } catch (uploadError) {
                this.logger.error(`Error uploading signed document to OneDrive for submission ${submissionId}: ${uploadError.message}`);
                // Don't throw - we don't want to fail the webhook if OneDrive upload fails
              }
            }
          }
          break;

        case 'form.declined':
        case 'submission.declined':
        case 'submission.rejected':
          this.logger.log(`❌ Submission ${submissionId} declined - updating status`);
          await this.updateSubmissionStatus(
            dbSubmission.opportunityId,
            dbSubmission.documentType as 'CONTRACT' | 'DISCLAIMER' | 'BOOKING_CONFIRMATION',
            'declined'
          );
          break;

        case 'submission.expired':
          this.logger.log(`⏰ Submission ${submissionId} expired - updating status`);
          await this.updateSubmissionStatus(
            dbSubmission.opportunityId,
            dbSubmission.documentType as 'CONTRACT' | 'DISCLAIMER' | 'BOOKING_CONFIRMATION',
            'expired'
          );
          break;

        case 'form.viewed':
        case 'submission.viewed':
          this.logger.log(`👁️ Submission ${submissionId} viewed`);
          // Optionally track view events
          break;

        case 'form.started':
        case 'submission.started':
        case 'submission.created':
          this.logger.log(`🚀 Submission ${submissionId} started/created`);
          // Optionally track start events
          break;

        case 'submission.archived':
          this.logger.log(`📦 Submission ${submissionId} archived`);
          // Optionally handle archived events
          break;

        case 'template.created':
        case 'template.updated':
          this.logger.log(`📄 Template event: ${eventType}`);
          // Optionally handle template events
          break;

        default:
          this.logger.log(`ℹ️ Unhandled webhook event type: ${eventType}`);
          // For unknown events, try to sync status anyway
          await this.syncSubmissionStatusFromDocuSeal(submissionId);
      }

      this.logger.log(`✅ Successfully processed webhook event: ${eventType} for submission: ${submissionId}`);
    } catch (error) {
      this.logger.error(`❌ Error handling webhook event ${eventType}: ${error.message}`);
      throw error;
    }
  }
}
