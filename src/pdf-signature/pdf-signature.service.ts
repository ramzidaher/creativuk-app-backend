import { Injectable, Logger } from '@nestjs/common';
import { PDFDocument, PDFPage, PDFImage } from 'pdf-lib';
import * as fs from 'fs/promises';
import * as path from 'path';

@Injectable()
export class PdfSignatureService {
  private readonly logger = new Logger(PdfSignatureService.name);

    async addSignatureToPDF(
    pdfPath: string,
    signatureData: string,
    pageNumbers: number[] = [6, 19, 21, 23]
  ): Promise<{ success: boolean; message: string; error?: string }> {
    try {
      this.logger.log(`Adding signature to PDF: ${pdfPath} on pages ${pageNumbers.join(', ')}`);

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
      
      // Calculate signature position and size (resized to fit properly)
      const maxSignatureWidth = 150;
      const maxSignatureHeight = 75;
      const x = 250;  // Fixed x position
      const y = 180;  // Fixed y position

      // Resize signature to fit within bounds while maintaining aspect ratio
      const { width: originalWidth, height: originalHeight } = embeddedImage;
      let signatureWidth = maxSignatureWidth;
      let signatureHeight = maxSignatureHeight;

      // Calculate aspect ratio and resize if needed
      const aspectRatio = originalWidth / originalHeight;
      if (aspectRatio > maxSignatureWidth / maxSignatureHeight) {
        // Width is the limiting factor
        signatureWidth = maxSignatureWidth;
        signatureHeight = maxSignatureWidth / aspectRatio;
      } else {
        // Height is the limiting factor
        signatureHeight = maxSignatureHeight;
        signatureWidth = maxSignatureHeight * aspectRatio;
      }

      this.logger.log(`Adding signature at position: x=${x}, y=${y}, width=${signatureWidth.toFixed(1)}, height=${signatureHeight.toFixed(1)}`);

      // Add the signature image to each specified page
      for (const pageNumber of pageNumbers) {
        const page = pdfDoc.getPage(pageNumber - 1);
        
        page.drawImage(embeddedImage, {
          x,
          y,
          width: signatureWidth,
          height: signatureHeight,
        });

        this.logger.log(`Signature added to page ${pageNumber}`);
      }

      // Save the modified PDF
      const modifiedPdfBytes = await pdfDoc.save();
      await fs.writeFile(pdfPath, modifiedPdfBytes);

      this.logger.log(`Signature added successfully to pages ${pageNumbers.join(', ')}`);
      return {
        success: true,
        message: `Signature added to pages ${pageNumbers.join(', ')}`
      };

    } catch (error) {
      this.logger.error(`Error adding signature to PDF: ${error.message}`);
      return {
        success: false,
        message: 'Failed to add signature to PDF',
        error: error.message
      };
    }
  }

  private async convertBase64ToImage(base64Data: string): Promise<Uint8Array> {
    try {
      // Remove data URL prefix if present
      const base64String = base64Data.includes(',') 
        ? base64Data.split(',')[1] 
        : base64Data;

      // Convert base64 to buffer
      const buffer = Buffer.from(base64String, 'base64');
      return new Uint8Array(buffer);
    } catch (error) {
      throw new Error(`Failed to convert base64 to image: ${error.message}`);
    }
  }

  async getPDFPageCount(pdfPath: string): Promise<number> {
    try {
      const pdfBytes = await fs.readFile(pdfPath);
      const pdfDoc = await PDFDocument.load(pdfBytes);
      return pdfDoc.getPageCount();
    } catch (error) {
      this.logger.error(`Error getting PDF page count: ${error.message}`);
      throw error;
    }
  }
}
