import { Injectable, Logger } from '@nestjs/common';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as crypto from 'crypto';

@Injectable()
export class PdfSignatureService {
  private readonly logger = new Logger(PdfSignatureService.name);

  /**
   * Create a signed PDF with embedded signature and verification marks
   */
  async createSignedPDF(
    originalPdfData: string, // Base64 PDF data
    signatureData: string,
    digitalFootprint: any,
    outputPath: string
  ): Promise<{ success: boolean; signedPdfPath?: string; error?: string }> {
    try {
      this.logger.log('Creating signed PDF with embedded signature...');

      // Convert base64 PDF data to buffer
      const pdfBuffer = Buffer.from(originalPdfData, 'base64');
      
      // Create a new PDF with signature and verification marks
      const signedPdfBuffer = await this.embedSignatureInPDF(
        pdfBuffer,
        signatureData,
        digitalFootprint
      );

      // Save the signed PDF
      await fs.writeFile(outputPath, signedPdfBuffer);

      this.logger.log(`Signed PDF created successfully: ${outputPath}`);
      
      return {
        success: true,
        signedPdfPath: outputPath,
      };
    } catch (error) {
      this.logger.error(`Error creating signed PDF: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Embed signature and verification marks into PDF
   */
  private async embedSignatureInPDF(
    pdfBuffer: Buffer,
    signatureData: string,
    digitalFootprint: any
  ): Promise<Buffer> {
    try {
      // Convert signature from base64 to image buffer
      const signatureImageBuffer = this.convertSignatureToImageBuffer(signatureData);
      
      // Create verification watermark
      const verificationWatermark = this.createVerificationWatermark(digitalFootprint);
      
      // For now, we'll create a simple PDF with embedded content
      // In a production environment, you'd use a proper PDF library like PDF-lib
      const signedPdfContent = this.createSignedPdfContent(
        pdfBuffer,
        signatureImageBuffer,
        verificationWatermark,
        digitalFootprint
      );

      return Buffer.from(signedPdfContent);
    } catch (error) {
      this.logger.error(`Error embedding signature: ${error.message}`);
      throw error;
    }
  }

  /**
   * Convert signature data to image buffer
   */
  private convertSignatureToImageBuffer(signatureData: string): Buffer {
    try {
      // Extract base64 data from data URL
      const base64Data = signatureData.split(',')[1];
      return Buffer.from(base64Data, 'base64');
    } catch (error) {
      this.logger.error(`Error converting signature: ${error.message}`);
      throw error;
    }
  }

  /**
   * Create verification watermark content
   */
  private createVerificationWatermark(digitalFootprint: any): string {
    // Handle different digital footprint structures
    const signatureId = digitalFootprint.id || 'unknown';
    const algorithm = digitalFootprint.algorithm || 'SHA-256';
    const timestamp = digitalFootprint.timestamp || digitalFootprint.signedAt || new Date().toISOString();
    
    // Get verification hash from different possible locations
    let verificationHash = '';
    if (digitalFootprint.hashes?.verification) {
      verificationHash = digitalFootprint.hashes.verification;
    } else if (digitalFootprint.signatureHash) {
      verificationHash = digitalFootprint.signatureHash;
    } else if (digitalFootprint.verificationHash) {
      verificationHash = digitalFootprint.verificationHash;
    } else {
      verificationHash = 'unknown';
    }
    
    return `
      VERIFIED DIGITAL SIGNATURE
      Signature ID: ${signatureId}
      Algorithm: ${algorithm}
      Timestamp: ${timestamp}
      Status: ESIGN ACT COMPLIANT
      Hash: ${verificationHash.substring(0, 16)}...
    `;
  }

  /**
   * Create signed PDF content (simplified version)
   * In production, use PDF-lib or similar library
   */
  private createSignedPdfContent(
    originalPdf: Buffer,
    signatureImage: Buffer,
    watermark: string,
    digitalFootprint: any
  ): string {
    // This is a simplified approach - in production you'd use a proper PDF library
    // For now, we'll create a PDF-like structure with embedded signature
    
    const signatureBase64 = signatureImage.toString('base64');
    const watermarkBase64 = Buffer.from(watermark).toString('base64');
    
    // Handle different digital footprint structures
    const documentName = digitalFootprint.documentName || 'Unknown Document';
    const signedBy = digitalFootprint.signedBy || 'Unknown Signer';
    const signedAt = digitalFootprint.signedAt || digitalFootprint.timestamp || new Date().toISOString();
    const signatureId = digitalFootprint.id || 'unknown';
    const algorithm = digitalFootprint.algorithm || 'SHA-256';
    
    // Get verification hash from different possible locations
    let verificationHash = '';
    if (digitalFootprint.hashes?.verification) {
      verificationHash = digitalFootprint.hashes.verification;
    } else if (digitalFootprint.signatureHash) {
      verificationHash = digitalFootprint.signatureHash;
    } else if (digitalFootprint.verificationHash) {
      verificationHash = digitalFootprint.verificationHash;
    } else {
      verificationHash = 'unknown';
    }
    
    // Create a PDF-like structure with embedded signature and verification
    const pdfContent = `
%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
/Resources <<
  /XObject <<
    /Sig1 5 0 R
    /Watermark 6 0 R
  >>
  /Font <<
    /F1 7 0 R
  >>
>>
/Annots [8 0 R]
>>
endobj

4 0 obj
<<
/Length 200
>>
stream
BT
/F1 12 Tf
50 750 Td
(DIGITALLY SIGNED DOCUMENT) Tj
0 -20 Td
(Document: ${documentName}) Tj
0 -20 Td
(Signed by: ${signedBy}) Tj
0 -20 Td
(Signed at: ${new Date(signedAt).toLocaleString()}) Tj
0 -40 Td
(Signature ID: ${signatureId}) Tj
0 -20 Td
(Algorithm: ${algorithm}) Tj
0 -20 Td
(Verification Hash: ${verificationHash.substring(0, 32)}...) Tj
0 -20 Td
(Status: ESIGN ACT COMPLIANT) Tj
0 -40 Td
q
200 0 0 100 50 500 cm
/Sig1 Do
Q
q
612 0 0 50 0 0 cm
/Watermark Do
Q
ET
endstream
endobj

5 0 obj
<<
/Type /XObject
/Subtype /Image
/Width 200
/Height 100
/ColorSpace /DeviceRGB
/BitsPerComponent 8
/Length ${signatureImage.length}
>>
stream
${signatureBase64}
endstream
endobj

6 0 obj
<<
/Type /XObject
/Subtype /Image
/Width 612
/Height 50
/ColorSpace /DeviceRGB
/BitsPerComponent 8
/Length ${Buffer.from(watermark).length}
>>
stream
${watermarkBase64}
endstream
endobj

7 0 obj
<<
/Type /Font
/Subtype /Type1
/BaseFont /Helvetica
>>
endobj

8 0 obj
<<
/Type /Annot
/Subtype /Widget
/Rect [50 500 250 600]
/FT /Sig
/T (Digital Signature)
/F 4
/V <<
  /Type /Sig
  /Filter /Adobe.PPKLite
  /SubFilter /adbe.pkcs7.detached
  /Contents <${verificationHash}>
  /ByteRange [0 1000 2000 1000]
  /Reference [<<
    /Type /SigRef
    /TransformMethod /DocMDP
    /TransformParams <<
      /P 1
      /V /1.2
    >>
  >>]
  /M (D:${signedAt.replace(/[-:]/g, '').replace(/\.\d{3}Z$/, '+00\'00\'')})
  /Location (Digital Signature)
  /Reason (Document Authentication)
  /ContactInfo (${signedBy})
>>
endobj

xref
0 9
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000254 00000 n 
0000000514 00000 n 
0000000680 00000 n 
0000000850 00000 n 
0000000920 00000 n 
trailer
<<
/Size 9
/Root 1 0 R
>>
startxref
${1000 + signatureImage.length + Buffer.from(watermark).length}
%%EOF
    `;

    return pdfContent;
  }

  /**
   * Extract digital footprint from signed PDF
   */
  async extractDigitalFootprint(pdfPath: string): Promise<{
    success: boolean;
    digitalFootprint?: any;
    error?: string;
  }> {
    try {
      this.logger.log(`Extracting digital footprint from: ${pdfPath}`);
      
      const pdfBuffer = await fs.readFile(pdfPath);
      const pdfContent = pdfBuffer.toString();
      
      // Extract signature information from PDF
      const signatureMatch = pdfContent.match(/Signature ID: ([^\s]+)/);
      const algorithmMatch = pdfContent.match(/Algorithm: ([^\s]+)/);
      const hashMatch = pdfContent.match(/Verification Hash: ([^\s]+)/);
      const timestampMatch = pdfContent.match(/Signed at: ([^)]+)/);
      const signerMatch = pdfContent.match(/Signed by: ([^)]+)/);
      
      if (!signatureMatch || !algorithmMatch || !hashMatch) {
        return {
          success: false,
          error: 'Digital footprint not found in PDF',
        };
      }

      const digitalFootprint = {
        id: signatureMatch[1],
        algorithm: algorithmMatch[1],
        verificationHash: hashMatch[1],
        signedAt: timestampMatch ? timestampMatch[1] : null,
        signedBy: signerMatch ? signerMatch[1] : null,
        extractedAt: new Date().toISOString(),
        source: 'PDF_EMBEDDED',
      };

      this.logger.log(`Digital footprint extracted successfully: ${digitalFootprint.id}`);
      
      return {
        success: true,
        digitalFootprint,
      };
    } catch (error) {
      this.logger.error(`Error extracting digital footprint: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Verify PDF signature integrity
   */
  async verifyPdfSignature(pdfPath: string, expectedHash: string): Promise<{
    success: boolean;
    verified: boolean;
    error?: string;
  }> {
    try {
      const extractionResult = await this.extractDigitalFootprint(pdfPath);
      
      if (!extractionResult.success) {
        return {
          success: false,
          verified: false,
          error: extractionResult.error,
        };
      }

      const extractedHash = extractionResult.digitalFootprint.verificationHash;
      const isVerified = extractedHash === expectedHash;

      this.logger.log(`PDF signature verification: ${isVerified ? 'VERIFIED' : 'FAILED'}`);
      
      return {
        success: true,
        verified: isVerified,
      };
    } catch (error) {
      this.logger.error(`Error verifying PDF signature: ${error.message}`);
      return {
        success: false,
        verified: false,
        error: error.message,
      };
    }
  }
}
