import { Controller, Post, Get, Body, Param, HttpException, HttpStatus } from '@nestjs/common';
import * as fs from 'fs';
import * as path from 'path';

@Controller('adobe-sign')
export class AdobeSignController {
  private readonly integrationKey = '3AAABLblqZhBa7sdLD0vFx66nBVEx4b1ValKW0TGwkEWwWbCRfu6B4bjuLxUREFeEZWU4pod_f-pqJR86bD0VpRS74mbIqVba';
  private readonly baseUrl = 'https://api.eu1.adobesign.com/api/rest/v6';

  /**
   * Helper method to truncate base64 data for logging
   */
  private truncateBase64ForLogging(data: any): any {
    if (typeof data === 'string' && data.length > 100 && data.match(/^[A-Za-z0-9+/=]+$/)) {
      // Likely base64 data, truncate it
      return `${data.substring(0, 50)}... (truncated, ${data.length} chars)`;
    }
    
    if (Array.isArray(data)) {
      return data.map(item => this.truncateBase64ForLogging(item));
    }
    
    if (data && typeof data === 'object') {
      const truncated: any = {};
      for (const [key, value] of Object.entries(data)) {
        if (key === 'base64' || key === 'base64Data' || key === 'fileContent' || key === 'document') {
          truncated[key] = typeof value === 'string' && value.length > 100 
            ? `${value.substring(0, 50)}... (truncated, ${value.length} chars)`
            : value;
        } else {
          truncated[key] = this.truncateBase64ForLogging(value);
        }
      }
      return truncated;
    }
    
    return data;
  }

  @Get('test')
  async testConnection() {
    return {
      success: true,
      message: 'Adobe Sign controller is working',
      integrationKey: this.integrationKey.substring(0, 20) + '...',
      baseUrl: this.baseUrl,
      timestamp: new Date().toISOString()
    };
  }

  @Post('send-for-signing/:agreementId')
  async sendForSigning(@Param('agreementId') agreementId: string) {
    try {
      console.log('📤 Sending agreement for signing:', agreementId);
      
      const response = await fetch(`${this.baseUrl}/agreements/${agreementId}/state`, {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          state: "SEND_FOR_SIGNATURE"
        }),
      });

      if (!response.ok) {
        const errorData = await response.text();
        console.error('❌ Failed to send agreement for signing:', errorData);
        throw new HttpException(
          `Failed to send agreement for signing: ${response.status} - ${errorData}`,
          response.status
        );
      }

      const result = await response.json();
      console.log('✅ Agreement sent for signing:', result);
      
      return {
        success: true,
        message: 'Agreement sent for signing successfully',
        data: result,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('❌ Error sending agreement for signing:', error);
      throw new HttpException(
        `Error sending agreement for signing: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('agreement-documents/:agreementId')
  async getAgreementDocuments(@Param('agreementId') agreementId: string) {
    try {
      console.log('🔍 Getting documents for agreement:', agreementId);
      
      const response = await fetch(`${this.baseUrl}/agreements/${agreementId}/documents`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        const errorData = await response.text();
        console.error('❌ Failed to get agreement documents:', errorData);
        throw new HttpException(
          `Failed to get agreement documents: ${response.status} - ${errorData}`,
          response.status
        );
      }

      const documents = await response.json();
      console.log('✅ Agreement documents retrieved:', documents);
      
      return {
        success: true,
        message: 'Agreement documents retrieved successfully',
        data: documents,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('❌ Error getting agreement documents:', error);
      throw new HttpException(
        `Error getting agreement documents: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('document-info/:agreementId/:documentId')
  async getDocumentInfo(@Param('agreementId') agreementId: string, @Param('documentId') documentId: string) {
    try {
      console.log('🔍 Getting document info for:', agreementId, documentId);
      
      const response = await fetch(`${this.baseUrl}/agreements/${agreementId}/documents/${documentId}`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        const errorData = await response.text();
        console.error('❌ Failed to get document info:', errorData);
        throw new HttpException(
          `Failed to get document info: ${response.status} - ${errorData}`,
          response.status
        );
      }

      const documentInfo = await response.json();
      console.log('✅ Document info retrieved:', documentInfo);
      
      return {
        success: true,
        message: 'Document info retrieved successfully',
        data: documentInfo,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('❌ Error getting document info:', error);
      throw new HttpException(
        `Error getting document info: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('document-pdf/:agreementId/:documentId')
  async getDocumentPdf(@Param('agreementId') agreementId: string, @Param('documentId') documentId: string) {
    try {
      console.log('🔍 Getting PDF document for:', agreementId, documentId);
      
      const response = await fetch(`${this.baseUrl}/agreements/${agreementId}/documents/${documentId}`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Accept': 'application/pdf',
        },
      });

      if (!response.ok) {
        const errorData = await response.text();
        console.error('❌ Failed to get PDF document:', errorData);
        throw new HttpException(
          `Failed to get PDF document: ${response.status} - ${errorData}`,
          response.status
        );
      }

      const pdfBuffer = await response.arrayBuffer();
      console.log('✅ PDF document retrieved:', pdfBuffer.byteLength, 'bytes');
      
      return {
        success: true,
        message: 'PDF document retrieved successfully',
        data: {
          size: pdfBuffer.byteLength,
          buffer: Buffer.from(pdfBuffer).toString('base64')
        },
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('❌ Error getting PDF document:', error);
      throw new HttpException(
        `Error getting PDF document: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('form-fields/:agreementId')
  async getFormFields(@Param('agreementId') agreementId: string) {
    try {
      console.log('🔍 Getting form fields for agreement:', agreementId);
      
      const response = await fetch(`${this.baseUrl}/agreements/${agreementId}/formFields`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        const errorData = await response.text();
        console.error('❌ Failed to get form fields:', errorData);
        throw new HttpException(
          `Failed to get form fields: ${response.status} - ${errorData}`,
          response.status
        );
      }

      const formFields = await response.json();
      console.log('✅ Form fields retrieved:', formFields);
      
      return {
        success: true,
        message: 'Form fields retrieved successfully',
        data: formFields,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('❌ Error getting form fields:', error);
      throw new HttpException(
        `Error getting form fields: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('basic-test')
  async basicTest() {
    try {
      console.log('🔍 Testing Adobe Sign API connection...');
      
      // Test 1: Try baseUris endpoint
      const baseUrisResponse = await fetch(`${this.baseUrl}/baseUris`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
      });

      console.log('🔍 BaseUris response status:', baseUrisResponse.status);

      if (baseUrisResponse.ok) {
        const baseUrisData = await baseUrisResponse.json();
        console.log('✅ BaseUris test successful:', baseUrisData);
        return {
          success: true,
          message: 'Adobe Sign API connection successful',
          baseUris: baseUrisData,
          timestamp: new Date().toISOString()
        };
      } else {
        const errorData = await baseUrisResponse.text();
        console.error('❌ BaseUris test failed:', errorData);
        return {
          success: false,
          message: 'Adobe Sign API connection failed',
          error: errorData,
          status: baseUrisResponse.status,
          timestamp: new Date().toISOString()
        };
      }
    } catch (error) {
      console.error('❌ Basic test error:', error);
      return {
        success: false,
        message: 'Adobe Sign API connection error',
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }

  @Post('upload-and-sign')
  async uploadAndSign(@Body() body: any) {
    try {
      console.log('🔍 Starting complete PDF upload and signing workflow...');
      console.log('🔍 Request body structure:', Object.keys(this.truncateBase64ForLogging(body)));
      console.log('🔍 Request body type:', typeof body);
      console.log('🔍 Request body keys:', body ? Object.keys(body) : 'body is null/undefined');
      console.log('🔍 File path:', body?.filePath);
      console.log('🔍 Signer email:', body?.signerEmail);
      console.log('🔍 Agreement name:', body?.agreementName);

      // Step 1: Read PDF file from filesystem
      console.log('📤 Step 1: Reading PDF file from filesystem...');
      
      if (!body) {
        throw new HttpException('Request body is missing or invalid', HttpStatus.BAD_REQUEST);
      }
      
      if (!body?.filePath) {
        throw new HttpException('filePath is required in request body', HttpStatus.BAD_REQUEST);
      }
      
      if (!body?.signerEmail) {
        throw new HttpException('signerEmail is required in request body', HttpStatus.BAD_REQUEST);
      }
      
      if (!body?.agreementName) {
        throw new HttpException('agreementName is required in request body', HttpStatus.BAD_REQUEST);
      }
      
      const filePath = body.filePath.replace(/\\/g, '/');
      console.log('🔍 Reading file from:', filePath);
      
      if (!fs.existsSync(filePath)) {
        throw new HttpException(`File not found: ${filePath}`, HttpStatus.NOT_FOUND);
      }

      const fileBuffer = fs.readFileSync(filePath);
      const fileName = path.basename(filePath);
      console.log('✅ File read successfully:', fileName, fileBuffer.length, 'bytes');

      // Step 2: Upload PDF to Adobe Sign
      console.log('📤 Step 2: Uploading PDF to Adobe Sign...');
      const formData = new FormData();
      formData.append('File', new Blob([new Uint8Array(fileBuffer)]), fileName);

      const uploadResponse = await fetch(`${this.baseUrl}/transientDocuments`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
        },
        body: formData,
      });

      console.log('🔍 Upload response status:', uploadResponse.status);

      if (!uploadResponse.ok) {
        const errorData = await uploadResponse.text();
        console.error('❌ PDF upload failed:', errorData);
        throw new HttpException(
          `PDF upload failed: ${uploadResponse.status} - ${errorData}`,
          uploadResponse.status
        );
      }

      const uploadData = await uploadResponse.json();
      console.log('✅ PDF uploaded successfully:', uploadData);
      const transientDocumentId = uploadData.transientDocumentId;

      // Step 3: Create Agreement for Signing
      console.log('📝 Step 3: Creating agreement for signing...');
      const agreementPayload = {
        fileInfos: [
          {
            transientDocumentId: transientDocumentId
          }
        ],
        name: body.agreementName,
        participantSetsInfo: [
          {
            name: "Primary Signature",
            memberInfos: [
              {
                email: body.signerEmail
              }
            ],
            order: 1,
            role: "SIGNER"
          }
        ],
        signatureType: "ESIGN",
        state: "IN_PROCESS",
        message: "Please review and sign this solar installation contract. This document contains important terms and conditions for your solar energy system installation.",
        // Using text tags instead of formFieldGenerators for cleaner UI
        // Text tags should be embedded in the PDF: {{Sig_es_:signer1:signature}}
        formFieldGenerators: [
          {
            formFieldDescription: {
              contentType: "SIGNATURE",
              inputType: "SIGNATURE",
              backgroundColor: "0xFFFFFF",
              borderColor: "0x000000",
              borderWidth: "1",
              required: true
            },
             anchorTextInfo: {
               anchorText: "Page 19 of 23",
               anchoredFormFieldLocation: {
                 offsetX: -100,
                 offsetY: 300,
                 height: 23,
                 width: 223
               }
             },
            formFieldNamePrefix: "SignatureField1_",
            participantSetName: "Primary Signature",
            generatorType: "ANCHOR_TEXT"
          },
          {
            formFieldDescription: {
              contentType: "SIGNATURE",
              inputType: "SIGNATURE",
              backgroundColor: "0xFFFFFF",
              borderColor: "0x000000",
              borderWidth: "1",
              required: true
            },
            anchorTextInfo: {
              anchorText: "Page 21 of 23",
              anchoredFormFieldLocation: {
                offsetX:  -20,
                offsetY: 200,
                height: 43,
                width: 137
              }
            },
            formFieldNamePrefix: "SignatureField2_",
            participantSetName: "Primary Signature",
            generatorType: "ANCHOR_TEXT"
          },
          {
            formFieldDescription: {
              contentType: "SIGNATURE",
              inputType: "SIGNATURE",
              backgroundColor: "0xFFFFFF",
              borderColor: "0x000000",
              borderWidth: "1",
              required: true
            },
             anchorTextInfo: {
               anchorText: "Page 23 of 23",
               anchoredFormFieldLocation: {
                 offsetX: 0,
                 offsetY: -30,
                 height: 24,
                 width: 260
               }
             },
            formFieldNamePrefix: "SignatureField3_",
            participantSetName: "Primary Signature",
            generatorType: "ANCHOR_TEXT"
          }
        ]
      };

      console.log('🔍 Agreement payload structure:', Object.keys(this.truncateBase64ForLogging(agreementPayload)));
      
      const agreementResponse = await fetch(`${this.baseUrl}/agreements`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(agreementPayload),
      });

      console.log('🔍 Agreement creation response status:', agreementResponse.status);

      if (!agreementResponse.ok) {
        const errorData = await agreementResponse.text();
        console.error('❌ Agreement creation failed:', errorData);
        throw new HttpException(
          `Agreement creation failed: ${agreementResponse.status} - ${errorData}`,
          agreementResponse.status
        );
      }

      const agreementData = await agreementResponse.json();
      console.log('✅ Agreement created successfully:', agreementData);

      // Step 4: Wait for document processing
      console.log('⏳ Step 4: Waiting for document processing...');
      await new Promise(resolve => setTimeout(resolve, 10000));

      // Step 4.1: Form fields are already included in the agreement creation
      console.log('📝 Step 4.1: Form fields included in agreement creation via formFieldGenerators');
      
      // Step 4.5: Send agreement for signing (only if not in AUTHORING state)
      console.log('📤 Step 4.5: Checking agreement state...');
      
      if (agreementPayload.state === "IN_PROCESS") {
        console.log('📝 Agreement is in AUTHORING state - ready for field placement');
        console.log('🔗 Use the signing URL to place signature fields manually');
        console.log('📤 After placing fields, call the send-for-signing endpoint');
      } else {
        console.log('📤 Sending agreement for signing...');
        
        const sendAgreementResponse = await fetch(`${this.baseUrl}/agreements/${agreementData.id}/state`, {
          method: 'PUT',
          headers: {
            'Authorization': `Bearer ${this.integrationKey}`,
            'x-api-user': 'email:ramzi@ramzidaher.com',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            state: "SEND_FOR_SIGNATURE"
          }),
        });

        if (!sendAgreementResponse.ok) {
          const errorData = await sendAgreementResponse.text();
          console.error('❌ Failed to send agreement for signing:', errorData);
          console.log('⚠️ Continuing - agreement may already be in correct state');
        } else {
          const sendAgreementData = await sendAgreementResponse.json();
          console.log('✅ Agreement sent for signing:', sendAgreementData);
        }
      }

      // Step 5: Get URL (Signing URL or Authoring URL)
      let signingUrl: string | null = null;
      
      if (agreementPayload.state === "IN_PROCESS") {
        console.log('🔗 Step 5: Getting authoring URL...');
        const authoringUrlResponse = await fetch(`${this.baseUrl}/agreements/${agreementData.id}/views`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${this.integrationKey}`,
            'x-api-user': 'email:ramzi@ramzidaher.com',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            name: "AUTHORING"
          }),
        });

        if (authoringUrlResponse.ok) {
          const authoringUrlData = await authoringUrlResponse.json();
          console.log('✅ Authoring URL retrieved:', authoringUrlData);
          signingUrl = authoringUrlData.agreementViewList[0].url;
          console.log('🔗 Authoring URL:', signingUrl);
        } else {
          const errorData = await authoringUrlResponse.text();
          console.log('⚠️ Could not get authoring URL:', authoringUrlResponse.status, errorData);
        }
      } else {
        console.log('🔗 Step 5: Getting signing URL...');
        const signingUrlResponse = await fetch(`${this.baseUrl}/agreements/${agreementData.id}/signingUrls`, {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${this.integrationKey}`,
            'x-api-user': 'email:ramzi@ramzidaher.com',
            'Content-Type': 'application/json',
          },
        });

        if (signingUrlResponse.ok) {
          const signingUrlData = await signingUrlResponse.json();
          console.log('✅ Signing URL retrieved:', signingUrlData);
          if (signingUrlData.signingUrlSetInfos && signingUrlData.signingUrlSetInfos.length > 0) {
            const firstSet = signingUrlData.signingUrlSetInfos[0];
            if (firstSet.signingUrls && firstSet.signingUrls.length > 0) {
              signingUrl = firstSet.signingUrls[0].esignUrl;
              console.log('🔗 Signing URL:', signingUrl);
            }
          }
        } else {
          console.log('⚠️ Could not get signing URL:', signingUrlResponse.status, await signingUrlResponse.text());
        }
      }

      // Step 6: Get Agreement Details
      console.log('📋 Step 6: Getting agreement details...');
      const agreementDetailsResponse = await fetch(`${this.baseUrl}/agreements/${agreementData.id}`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
      });

      let agreementDetails: any = null;
      if (agreementDetailsResponse.ok) {
        agreementDetails = await agreementDetailsResponse.json();
        console.log('✅ Agreement details retrieved:', agreementDetails);
      } else {
        console.log('⚠️ Could not get agreement details:', agreementDetailsResponse.status);
      }

      return {
        success: true,
        message: 'PDF uploaded and agreement created successfully with multiple signature fields',
        data: {
          agreementId: agreementData.id,
          transientDocumentId: transientDocumentId,
          signerEmail: body.signerEmail,
          agreementName: body.agreementName,
          status: agreementDetails?.status || 'UNKNOWN',
          signingUrl: signingUrl,
          agreementDetails: agreementDetails
        },
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('❌ Upload and sign workflow failed:', error);
      throw new HttpException(
        `Upload and sign workflow failed: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('agreement/:agreementId/signing-url')
  async getAgreementSigningUrl(@Body() body: { agreementId: string }) {
    try {
      console.log('🔗 Getting signing URL for agreement:', body.agreementId);
      
      const signingUrlResponse = await fetch(`${this.baseUrl}/agreements/${body.agreementId}/signingUrls`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.integrationKey}`,
          'x-api-user': 'email:ramzi@ramzidaher.com',
          'Content-Type': 'application/json',
        },
      });

      if (!signingUrlResponse.ok) {
        const errorData = await signingUrlResponse.text();
        console.error('❌ Failed to get signing URL:', errorData);
        throw new HttpException(
          `Failed to get signing URL: ${signingUrlResponse.status} - ${errorData}`,
          signingUrlResponse.status
        );
      }

      const signingUrlData = await signingUrlResponse.json();
      console.log('✅ Signing URL retrieved:', signingUrlData);

      let signingUrl = null;
      if (signingUrlData.signingUrlSetInfos && signingUrlData.signingUrlSetInfos.length > 0) {
        const firstSet = signingUrlData.signingUrlSetInfos[0];
        if (firstSet.signingUrls && firstSet.signingUrls.length > 0) {
          signingUrl = firstSet.signingUrls[0].esignUrl;
        }
      }

      return {
        success: true,
        message: 'Signing URL retrieved successfully',
        data: {
          agreementId: body.agreementId,
          signingUrl: signingUrl,
          signingUrlData: signingUrlData
        },
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('❌ Get signing URL failed:', error);
      throw new HttpException(
        `Get signing URL failed: ${error.message}`,
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}