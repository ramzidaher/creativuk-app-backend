import { Injectable, Logger } from '@nestjs/common';
import * as fs from 'fs';
import * as path from 'path';

export interface ContractData {
  opportunityId: string;
  customerName: string;
  customerEmail: string;
  date: string;
  postcode: string;
  contractType: string;
  solarData?: {
    systemSize: string;
    estimatedSavings: string;
    paybackPeriod: string;
  };
}

@Injectable()
export class PdfGeneratorService {
  private readonly logger = new Logger(PdfGeneratorService.name);
  private readonly outputDir = path.join(__dirname, 'output');

  constructor() {
    this.ensureOutputDirectory();
  }

  private ensureOutputDirectory() {
    if (!fs.existsSync(this.outputDir)) {
      fs.mkdirSync(this.outputDir, { recursive: true });
      this.logger.log(`Created output directory: ${this.outputDir}`);
    }
  }

  async generateContractPdf(contractData: ContractData): Promise<string> {
    try {
      this.logger.log(`Generating contract PDF for opportunity: ${contractData.opportunityId}`);
      
      const fileName = `contract_${contractData.opportunityId}.pdf`;
      const filePath = path.join(this.outputDir, fileName);
      
      // Create a simple PDF content (in a real implementation, you'd use a PDF library like PDFKit)
      const pdfContent = this.generateSimplePdfContent(contractData);
      
      // For now, create a simple text file that represents the contract
      // In production, you'd use a proper PDF generation library
      const contractText = this.generateContractText(contractData);
      
      // Write the contract content to a file
      fs.writeFileSync(filePath.replace('.pdf', '.txt'), contractText);
      
      // Create a simple PDF-like file (this is a placeholder - in production use PDFKit or similar)
      const pdfBuffer = Buffer.from(contractText);
      fs.writeFileSync(filePath, pdfBuffer);
      
      this.logger.log(`Contract PDF generated: ${filePath}`);
      return filePath;
    } catch (error) {
      this.logger.error(`Failed to generate contract PDF: ${error.message}`);
      throw error;
    }
  }

  private generateContractText(contractData: ContractData): string {
    return `
SOLAR INSTALLATION CONTRACT

Contract ID: ${contractData.opportunityId}
Date: ${contractData.date}
Customer: ${contractData.customerName}
Email: ${contractData.customerEmail}
Postcode: ${contractData.postcode}

SOLAR SYSTEM DETAILS:
${contractData.solarData ? `
System Size: ${contractData.solarData.systemSize}
Estimated Annual Savings: ${contractData.solarData.estimatedSavings}
Payback Period: ${contractData.solarData.paybackPeriod}
` : ''}

TERMS AND CONDITIONS:
1. The customer agrees to the installation of solar panels as specified above.
2. All work will be completed by certified installers.
3. The system comes with a 25-year warranty on panels and 10-year warranty on inverters.
4. Payment terms: 50% deposit, 50% on completion.

CUSTOMER SIGNATURE: _________________________ Date: _________

INSTALLER SIGNATURE: _________________________ Date: _________

This contract is legally binding and subject to local regulations.
    `.trim();
  }

  private generateSimplePdfContent(contractData: ContractData): string {
    // This is a placeholder - in production you'd use PDFKit or similar
    return `%PDF-1.4
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
>>
endobj

4 0 obj
<<
/Length 200
>>
stream
BT
/F1 12 Tf
72 720 Td
(SOLAR INSTALLATION CONTRACT) Tj
0 -20 Td
(Contract ID: ${contractData.opportunityId}) Tj
0 -20 Td
(Date: ${contractData.date}) Tj
0 -20 Td
(Customer: ${contractData.customerName}) Tj
0 -20 Td
(Email: ${contractData.customerEmail}) Tj
0 -20 Td
(Postcode: ${contractData.postcode}) Tj
0 -40 Td
(Customer Signature: _________________________) Tj
0 -20 Td
(Date: _________) Tj
ET
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000204 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
454
%%EOF`;
  }

  async cleanupOldContracts(maxAgeHours: number = 24): Promise<void> {
    try {
      const files = fs.readdirSync(this.outputDir);
      const now = Date.now();
      const maxAge = maxAgeHours * 60 * 60 * 1000; // Convert hours to milliseconds

      for (const file of files) {
        const filePath = path.join(this.outputDir, file);
        const stats = fs.statSync(filePath);
        
        if (now - stats.mtime.getTime() > maxAge) {
          fs.unlinkSync(filePath);
          this.logger.log(`Cleaned up old contract file: ${file}`);
        }
      }
    } catch (error) {
      this.logger.error(`Failed to cleanup old contracts: ${error.message}`);
    }
  }
}
