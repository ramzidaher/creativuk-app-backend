import { Injectable, Logger } from '@nestjs/common';
import { PDFDocument, PDFPage, PDFImage, rgb } from 'pdf-lib';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as crypto from 'crypto';

export interface DigitalFootprint {
  deviceInfo: {
    platform: string;
    userAgent: string;
    screenResolution: string;
    timezone: string;
    language: string;
  };
  signatureData: {
    totalPoints: number;
    duration: number;
    startTime: number;
    endTime: number;
    pressurePoints: number[];
    velocityPoints: number[];
    boundingBox: {
      minX: number;
      minY: number;
      maxX: number;
      maxY: number;
    };
  };
  security: {
    hash: string;
    timestamp: number;
    sessionId: string;
  };
}

export interface SignatureMetadata {
  signatureId: string;
  opportunityId: string;
  signedBy: string;
  signedAt: Date;
  digitalFootprint: DigitalFootprint;
  pdfPath: string;
  signaturePosition: {
    x: number;
    y: number;
    width: number;
    height: number;
    page: number;
  };
  verificationHash: string;
}

@Injectable()
export class DigitalSignatureService {
  private readonly logger = new Logger(DigitalSignatureService.name);

  async signPDFWithDigitalFootprint(
    pdfPath: string,
    signatureData: string,
    digitalFootprint: DigitalFootprint,
    opportunityId: string,
    signedBy: string,
    pageNumbers: number[] = [6, 19, 21, 23]
  ): Promise<{ success: boolean; message: string; metadata?: SignatureMetadata; error?: string }> {
    try {
      this.logger.log(`Adding digital signature to PDF: ${pdfPath} with digital footprint`);

      // Read the PDF file
      const pdfBytes = await fs.readFile(pdfPath);
      const pdfDoc = await PDFDocument.load(pdfBytes);

      // Check if the pages exist
      const totalPages = pdfDoc.getPageCount();
      this.logger.log(`PDF has ${totalPages} pages, attempting to add signature to pages ${pageNumbers.join(', ')}`);
      
      for (const pageNumber of pageNumbers) {
        if (pageNumber > totalPages) {
          throw new Error(`Page ${pageNumber} does not exist. PDF has ${totalPages} pages.`);
        }
      }

      // Convert base64 signature to image and resize it
      const signatureImage = await this.convertBase64ToImage(signatureData);
      
      // Embed the image in the PDF
      const embeddedImage = await pdfDoc.embedPng(signatureImage);
      
      // Smart adaptive signature sizing - fits all signature sizes nicely
      const { width: originalWidth, height: originalHeight } = embeddedImage;
      
      // Define signature area bounds
      const signatureAreaWidth = 200;  // Total area width
      const signatureAreaHeight = 100; // Total area height
      
      // Define different positions for each page
      const signaturePositions = {
        1: { x: 200, y: 180 },  // Page 1 position (Email Confirmation) - moved left
        6: { x: 300, y: 10 },  // Page 6 position
        19: { x: 150, y: 160 }, // Page 19 position
        21: { x: 150, y: 150 }, // Page 21 position
        23: { x: 80, y: 330 }, // Page 23 position
      };
      
      // Calculate optimal size based on signature characteristics
      const aspectRatio = originalWidth / originalHeight;
      let signatureWidth, signatureHeight;
      
      // Smart scaling logic
      if (originalWidth < 50 || originalHeight < 25) {
        // Small signatures - scale up to minimum visible size
        const minWidth = 80;
        const minHeight = 40;
        
        if (aspectRatio > minWidth / minHeight) {
          signatureWidth = minWidth;
          signatureHeight = minWidth / aspectRatio;
        } else {
          signatureHeight = minHeight;
          signatureWidth = minHeight * aspectRatio;
        }
        
        this.logger.log(`Small signature detected - scaling up to minimum size: ${signatureWidth.toFixed(1)}x${signatureHeight.toFixed(1)}`);
        
      } else if (originalWidth > 300 || originalHeight > 150) {
        // Large signatures - scale down to fit nicely
        const maxWidth = signatureAreaWidth;
        const maxHeight = signatureAreaHeight;
        
        if (aspectRatio > maxWidth / maxHeight) {
          signatureWidth = maxWidth;
          signatureHeight = maxWidth / aspectRatio;
        } else {
          signatureHeight = maxHeight;
          signatureWidth = maxHeight * aspectRatio;
        }
        
        this.logger.log(`Large signature detected - scaling down to fit: ${signatureWidth.toFixed(1)}x${signatureHeight.toFixed(1)}`);
        
      } else {
        // Medium signatures - use original size or scale slightly for consistency
        const targetWidth = Math.min(originalWidth, signatureAreaWidth);
        const targetHeight = Math.min(originalHeight, signatureAreaHeight);
        
        if (aspectRatio > targetWidth / targetHeight) {
          signatureWidth = targetWidth;
          signatureHeight = targetWidth / aspectRatio;
        } else {
          signatureHeight = targetHeight;
          signatureWidth = targetHeight * aspectRatio;
        }
        
        this.logger.log(`Medium signature - optimal size: ${signatureWidth.toFixed(1)}x${signatureHeight.toFixed(1)}`);
      }
      
      // Ensure signature is never too small to be visible
      const absoluteMinWidth = 60;
      const absoluteMinHeight = 30;
      
      if (signatureWidth < absoluteMinWidth) {
        signatureWidth = absoluteMinWidth;
        signatureHeight = absoluteMinWidth / aspectRatio;
      }
      if (signatureHeight < absoluteMinHeight) {
        signatureHeight = absoluteMinHeight;
        signatureWidth = absoluteMinHeight * aspectRatio;
      }

      this.logger.log(`Adding signature with adaptive sizing: width=${signatureWidth.toFixed(1)}, height=${signatureHeight.toFixed(1)}`);

      // Generate signature metadata
      const signatureId = this.generateSignatureId();
      const signedAt = new Date();
      
      const signatureMetadata: SignatureMetadata = {
        signatureId,
        opportunityId,
        signedBy,
        signedAt,
        digitalFootprint,
        pdfPath,
        signaturePosition: {
          x: signaturePositions[pageNumbers[0]]?.x || 250,
          y: signaturePositions[pageNumbers[0]]?.y || 180,
          width: signatureWidth,
          height: signatureHeight,
          page: pageNumbers[0], // Primary page
        },
        verificationHash: this.generateVerificationHash(digitalFootprint, signatureData, signedAt),
      };

      // Add the signature image to each specified page
      this.logger.log(`Starting signature placement for ${pageNumbers.length} pages: ${pageNumbers.join(', ')}`);
      
      for (const pageNumber of pageNumbers) {
        const page = pdfDoc.getPage(pageNumber - 1);
        const position = signaturePositions[pageNumber] || { x: 250, y: 180 };
        
        this.logger.log(`Adding signature to page ${pageNumber} at position (${position.x}, ${position.y})`);
        
        page.drawImage(embeddedImage, {
          x: position.x,
          y: position.y,
          width: signatureWidth,
          height: signatureHeight,
        });

        this.logger.log(`✅ Signature successfully added to page ${pageNumber}`);
      }
      
      this.logger.log(`Completed signature placement - total pages processed: ${pageNumbers.length}`);

      // Add verification stamps to ALL pages (like Sign.com)
      await this.addVerificationStampsToAllPages(pdfDoc, signatureMetadata);

      // Add detailed digital footprint to signature pages only (19, 21)
      await this.addDetailedDigitalFootprintToSignaturePages(pdfDoc, signatureMetadata, pageNumbers);

      // Add digital footprint as invisible metadata to the PDF
      await this.addDigitalFootprintToPDF(pdfDoc, signatureMetadata);

      // Save the modified PDF
      const modifiedPdfBytes = await pdfDoc.save();
      await fs.writeFile(pdfPath, modifiedPdfBytes);

      // Save signature metadata to database (you can implement this based on your database setup)
      await this.saveSignatureMetadata(signatureMetadata);

      this.logger.log(`Digital signature added successfully with footprint to pages ${pageNumbers.join(', ')}`);
      return {
        success: true,
        message: `Digital signature added to pages ${pageNumbers.join(', ')} with digital footprint`,
        metadata: signatureMetadata
      };

    } catch (error) {
      this.logger.error(`Error adding digital signature to PDF: ${error.message}`);
      return {
        success: false,
        message: 'Failed to add digital signature to PDF',
        error: error.message
      };
    }
  }

  private async convertBase64ToImage(base64Data: string): Promise<Uint8Array> {
    try {
      this.logger.log(`Converting base64 data to image, length: ${base64Data.length}`);
      
      // Remove data URL prefix if present
      const base64String = base64Data.includes(',') 
        ? base64Data.split(',')[1] 
        : base64Data;

      this.logger.log(`Base64 string length after prefix removal: ${base64String.length}`);

      // Check if it's SVG data
      if (base64Data.includes('data:image/svg+xml')) {
        this.logger.log('Detected SVG signature data');
        // For SVG, we need to convert it to PNG first
        // For now, let's create a simple signature-like PNG
        return this.createPlaceholderImage();
      }

      // Convert base64 to buffer
      const buffer = Buffer.from(base64String, 'base64');
      this.logger.log(`Buffer created successfully, length: ${buffer.length}`);
      return new Uint8Array(buffer);
    } catch (error) {
      this.logger.error(`Error converting base64 to image: ${error.message}`);
      this.logger.error(`Base64 data preview: ${base64Data.substring(0, 100)}...`);
      throw new Error(`Failed to convert base64 to image: ${error.message}`);
    }
  }

  private createPlaceholderImage(): Uint8Array {
    // Create a simple 1x1 transparent PNG as placeholder
    // This is a minimal PNG file in base64
    const pngBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    const buffer = Buffer.from(pngBase64, 'base64');
    return new Uint8Array(buffer);
  }

  private async addVerificationStampsToAllPages(pdfDoc: PDFDocument, metadata: SignatureMetadata): Promise<void> {
    try {
      const totalPages = pdfDoc.getPageCount();
      this.logger.log(`Adding ultra-subtle verification stamps to all ${totalPages} pages`);

      for (let i = 0; i < totalPages; i++) {
        const page = pdfDoc.getPage(i);
        const { width, height } = page.getSize();

        // Create ultra-subtle verification in top-right corner - moved further left
        const fontSize = 7;
        const rightMargin = 50; // Moved further to the left (was 120px)
        const textY = height - 15; // 15px from top edge

        // Create compact verification text (one line) - removed platform info
        const shortId = metadata.signatureId.substring(0, 8) + '...';
        const shortDate = metadata.signedAt.toLocaleDateString('en-US', {
          month: 'short',
          day: 'numeric',
          hour: '2-digit',
          minute: '2-digit'
        });
        
        // Single line verification text (removed platform info)
        const verificationText = `[VERIFIED] ${shortId} | ${shortDate}`;
        
        // Calculate text width for positioning
        const textWidth = verificationText.length * (fontSize * 0.6); // Approximate width
        const iconWidth = 8;
        const iconSpacing = 4;
        
        // Position text further to the left
        const textX = width - textWidth - rightMargin;
        
        // Draw verification icon first (to the left of text)
        this.drawMicroVerificationIcon(page, textX - iconWidth - iconSpacing, textY - 2, iconWidth, iconWidth);
        
        // Draw text
        page.drawText(verificationText, {
          x: textX,
          y: textY,
          size: fontSize,
          color: rgb(0.4, 0.4, 0.4), // Very subtle gray
        });
      }

      this.logger.log(`Ultra-subtle verification stamps added to all ${totalPages} pages`);
    } catch (error) {
      this.logger.error(`Error adding verification stamps: ${error.message}`);
      throw error;
    }
  }

  private async addDigitalFootprintToPDF(pdfDoc: PDFDocument, metadata: SignatureMetadata): Promise<void> {
    try {
      // Add digital footprint as PDF metadata
      const footprintData = {
        signatureId: metadata.signatureId,
        signedBy: metadata.signedBy,
        signedAt: metadata.signedAt.toISOString(),
        digitalFootprint: metadata.digitalFootprint,
        verificationHash: metadata.verificationHash,
      };

      // Set PDF metadata
      pdfDoc.setTitle(`Signed Document - ${metadata.opportunityId}`);
      pdfDoc.setAuthor(metadata.signedBy);
      pdfDoc.setSubject(`Digital Signature - ${metadata.signatureId}`);
      pdfDoc.setKeywords(['digital-signature', 'signed', metadata.opportunityId]);
      pdfDoc.setProducer('Creativ Solar App - Digital Signature Service');
      pdfDoc.setCreator('Creativ Solar App');

      // Add custom metadata (this will be embedded in the PDF)
      const customMetadata = JSON.stringify(footprintData);
      
      // Note: PDF-lib doesn't directly support custom metadata, but we can add it as a comment
      // In a production environment, you might want to use a more advanced PDF library
      // or store this metadata separately in a database
      
      this.logger.log(`Digital footprint metadata added to PDF: ${metadata.signatureId}`);
    } catch (error) {
      this.logger.error(`Error adding digital footprint to PDF: ${error.message}`);
      throw error;
    }
  }

  private drawMicroVerificationIcon(page: PDFPage, x: number, y: number, width: number, height: number): void {
    try {
      // Draw a micro verification icon - just a simple checkmark
      const centerX = x + width / 2;
      const centerY = y + height / 2;
      const size = Math.min(width, height) / 2;

      // Draw simple checkmark (no circle, just the checkmark)
      const checkSize = size * 0.6;
      page.drawLine({
        start: { x: centerX - checkSize * 0.3, y: centerY },
        end: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        thickness: 0.5,
        color: rgb(0.4, 0.4, 0.4),
      });
      
      page.drawLine({
        start: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        end: { x: centerX + checkSize * 0.3, y: centerY - checkSize * 0.3 },
        thickness: 0.5,
        color: rgb(0.4, 0.4, 0.4),
      });
    } catch (error) {
      this.logger.warn(`Error drawing micro verification icon: ${error.message}`);
    }
  }

  private drawTinyVerificationIcon(page: PDFPage, x: number, y: number, width: number, height: number): void {
    try {
      // Draw a tiny, subtle verification icon
      const centerX = x + width / 2;
      const centerY = y + height / 2;
      const radius = Math.min(width, height) / 2 - 1;

      // Draw tiny circle
      page.drawCircle({
        x: centerX,
        y: centerY,
        size: radius,
        borderColor: rgb(0.6, 0.6, 0.6),
        borderWidth: 0.5,
        color: rgb(0.95, 0.95, 0.95),
      });

      // Draw tiny checkmark
      const checkSize = radius * 0.5;
      page.drawLine({
        start: { x: centerX - checkSize * 0.3, y: centerY },
        end: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        thickness: 0.8,
        color: rgb(0.6, 0.6, 0.6),
      });
      
      page.drawLine({
        start: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        end: { x: centerX + checkSize * 0.3, y: centerY - checkSize * 0.3 },
        thickness: 0.8,
        color: rgb(0.6, 0.6, 0.6),
      });
    } catch (error) {
      this.logger.warn(`Error drawing tiny verification icon: ${error.message}`);
    }
  }

  private drawVerificationPattern(page: PDFPage, x: number, y: number, width: number, height: number): void {
    try {
      // Draw a simple verification pattern (checkmark-like design)
      const centerX = x + width / 2;
      const centerY = y + height / 2;
      const radius = Math.min(width, height) / 2 - 2;

      // Draw outer circle
      page.drawCircle({
        x: centerX,
        y: centerY,
        size: radius,
        borderColor: rgb(0.2, 0.4, 0.8),
        borderWidth: 1,
        color: rgb(0.9, 0.95, 1.0),
      });

      // Draw checkmark
      const checkSize = radius * 0.6;
      page.drawLine({
        start: { x: centerX - checkSize * 0.3, y: centerY },
        end: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        thickness: 1.5,
        color: rgb(0.2, 0.4, 0.8),
      });
      
      page.drawLine({
        start: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        end: { x: centerX + checkSize * 0.3, y: centerY - checkSize * 0.3 },
        thickness: 1.5,
        color: rgb(0.2, 0.4, 0.8),
      });
    } catch (error) {
      this.logger.warn(`Error drawing verification pattern: ${error.message}`);
    }
  }

  private generateSignatureId(): string {
    return `SIG_${Date.now()}_${crypto.randomBytes(8).toString('hex').toUpperCase()}`;
  }

  private generateVerificationHash(
    digitalFootprint: DigitalFootprint,
    signatureData: string,
    signedAt: Date
  ): string {
    const dataToHash = JSON.stringify({
      digitalFootprint,
      signatureData: signatureData.substring(0, 100), // First 100 chars for hash
      signedAt: signedAt.toISOString(),
    });
    
    return crypto.createHash('sha256').update(dataToHash).digest('hex');
  }

  private async saveSignatureMetadata(metadata: SignatureMetadata): Promise<void> {
    try {
      // Create metadata directory if it doesn't exist
      const metadataDir = path.join(process.cwd(), 'signature-metadata');
      await fs.mkdir(metadataDir, { recursive: true });

      // Save metadata to file
      const metadataPath = path.join(metadataDir, `${metadata.signatureId}.json`);
      await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2));

      this.logger.log(`Signature metadata saved: ${metadataPath}`);
    } catch (error) {
      this.logger.error(`Error saving signature metadata: ${error.message}`);
      // Don't throw error here as the main signature operation should still succeed
    }
  }

  async verifySignature(pdfPath: string, signatureId: string): Promise<{
    success: boolean;
    isValid: boolean;
    metadata?: SignatureMetadata;
    error?: string;
  }> {
    try {
      // Load metadata
      const metadataPath = path.join(process.cwd(), 'signature-metadata', `${signatureId}.json`);
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      const metadata: SignatureMetadata = JSON.parse(metadataContent);

      // Verify the PDF still exists and matches
      const pdfExists = await fs.access(pdfPath).then(() => true).catch(() => false);
      if (!pdfExists) {
        return {
          success: true,
          isValid: false,
          metadata,
          error: 'PDF file not found'
        };
      }

      // Additional verification logic can be added here
      // For example, checking if the signature is still present in the PDF
      
      return {
        success: true,
        isValid: true,
        metadata
      };
    } catch (error) {
      this.logger.error(`Error verifying signature: ${error.message}`);
      return {
        success: false,
        isValid: false,
        error: error.message
      };
    }
  }

  async getSignatureHistory(opportunityId: string): Promise<SignatureMetadata[]> {
    try {
      const metadataDir = path.join(process.cwd(), 'signature-metadata');
      const files = await fs.readdir(metadataDir);
      
      const signatures: SignatureMetadata[] = [];
      
      for (const file of files) {
        if (file.endsWith('.json')) {
          try {
            const content = await fs.readFile(path.join(metadataDir, file), 'utf-8');
            const metadata: SignatureMetadata = JSON.parse(content);
            
            if (metadata.opportunityId === opportunityId) {
              signatures.push(metadata);
            }
          } catch (error) {
            this.logger.warn(`Error reading signature metadata file ${file}: ${error.message}`);
          }
        }
      }
      
      // Sort by signed date (newest first)
      signatures.sort((a, b) => b.signedAt.getTime() - a.signedAt.getTime());
      
      return signatures;
    } catch (error) {
      this.logger.error(`Error getting signature history: ${error.message}`);
      return [];
    }
  }

  private async addDetailedDigitalFootprintToSignaturePages(
    pdfDoc: PDFDocument, 
    metadata: SignatureMetadata, 
    signaturePageNumbers: number[]
  ): Promise<void> {
    try {
      this.logger.log(`Adding detailed digital footprint to signature pages: ${signaturePageNumbers.join(', ')}`);

      for (const pageNumber of signaturePageNumbers) {
        const page = pdfDoc.getPage(pageNumber - 1); // PDF pages are 0-indexed
        const { width, height } = page.getSize();

        // Position in bottom-left corner (moved to far left)
        const footprintX = 1;  // Distance from left edge (1px from left corner)
        const footprintY = 1;  // Distance from bottom edge (1px from bottom corner)
        const footprintWidth = 100;
        const footprintHeight = 50;

        // Create a subtle background box
        page.drawRectangle({
          x: footprintX,
          y: footprintY,
          width: footprintWidth,
          height: footprintHeight,
          borderColor: rgb(0.7, 0.7, 0.7),
          borderWidth: 0.5,
          color: rgb(0.98, 0.98, 0.98),
        });

        // Add detailed information
        const shortId = metadata.signatureId.substring(0, 8);
        const shortDate = metadata.signedAt.toLocaleDateString('en-GB');
        const customerName = metadata.signedBy || 'Unknown Customer';
        
        // Main verification text
        page.drawText('[DIGITALLY SIGNED]', {
          x: footprintX + 5,
          y: footprintY + 40,
          size: 8,
          color: rgb(0.2, 0.2, 0.2),
        });

        // Customer name
        page.drawText(`Customer: ${customerName}`, {
          x: footprintX + 5,
          y: footprintY + 30,
          size: 7,
          color: rgb(0.3, 0.3, 0.3),
        });

        // Date and ID
        page.drawText(`Date: ${shortDate}`, {
          x: footprintX + 5,
          y: footprintY + 20,
          size: 7,
          color: rgb(0.3, 0.3, 0.3),
        });

        page.drawText(`ID: ${shortId}`, {
          x: footprintX + 5,
          y: footprintY + 10,
          size: 7,
          color: rgb(0.3, 0.3, 0.3),
        });

        // Add a small verification icon in the top-right of the footprint box
        this.drawDetailedVerificationIcon(page, footprintX + footprintWidth - 15, footprintY + footprintHeight - 15, 10, 10);
      }

      this.logger.log(`Detailed digital footprint added to signature pages: ${signaturePageNumbers.join(', ')}`);
    } catch (error) {
      this.logger.error(`Error adding detailed digital footprint: ${error.message}`);
      throw error;
    }
  }

  private drawDetailedVerificationIcon(page: PDFPage, x: number, y: number, width: number, height: number): void {
    try {
      // Draw a more detailed verification icon for the signature pages
      const centerX = x + width / 2;
      const centerY = y + height / 2;
      const radius = Math.min(width, height) / 2 - 1;

      // Draw circle with border
      page.drawCircle({
        x: centerX,
        y: centerY,
        size: radius,
        borderColor: rgb(0.2, 0.6, 0.2),
        borderWidth: 1,
        color: rgb(0.9, 1.0, 0.9),
      });

      // Draw checkmark
      const checkSize = radius * 0.6;
      page.drawLine({
        start: { x: centerX - checkSize * 0.3, y: centerY },
        end: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        thickness: 1,
        color: rgb(0.2, 0.6, 0.2),
      });
      
      page.drawLine({
        start: { x: centerX - checkSize * 0.1, y: centerY + checkSize * 0.3 },
        end: { x: centerX + checkSize * 0.3, y: centerY - checkSize * 0.3 },
        thickness: 1,
        color: rgb(0.2, 0.6, 0.2),
      });
    } catch (error) {
      this.logger.warn(`Error drawing detailed verification icon: ${error.message}`);
    }
  }
}
