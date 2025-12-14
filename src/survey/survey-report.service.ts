import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { SurveyImageService } from './survey-image.service';
import * as fs from 'fs';
import * as path from 'path';
import { promisify } from 'util';
import * as puppeteer from 'puppeteer';

const writeFile = promisify(fs.writeFile);
const mkdir = promisify(fs.mkdir);

@Injectable()
export class SurveyReportService {
  private readonly logger = new Logger(SurveyReportService.name);
  private readonly uploadsDir = path.join(process.cwd(), 'uploads', 'survey-reports');

  constructor(
    private readonly prisma: PrismaService,
    private readonly surveyImageService: SurveyImageService
  ) {
    this.ensureUploadsDirectory();
  }

  private async ensureUploadsDirectory() {
    try {
      if (!fs.existsSync(this.uploadsDir)) {
        await mkdir(this.uploadsDir, { recursive: true });
        this.logger.log(`Created survey reports directory: ${this.uploadsDir}`);
      }
    } catch (error) {
      this.logger.error(`Failed to create survey reports directory: ${error.message}`);
    }
  }

  async generateHtmlReport(opportunityId: string, surveyData: any, surveyId?: string): Promise<string> {
    try {
      const reportDir = path.join(this.uploadsDir, opportunityId);
      if (!fs.existsSync(reportDir)) {
        await mkdir(reportDir, { recursive: true });
      }

      // Get all images for this survey
      let images: any[] = [];
      if (surveyId) {
        images = await this.surveyImageService.getSurveyImages(surveyId);
      } else {
        // Fallback to opportunity-based lookup (for backward compatibility)
        images = await this.surveyImageService.getSurveyImagesByOpportunity(opportunityId);
      }
      
      // Group images by field name
      const imagesByField: { [key: string]: any[] } = {};
      images.forEach(img => {
        if (!imagesByField[img.fieldName]) {
          imagesByField[img.fieldName] = [];
        }
        imagesByField[img.fieldName].push(img);
      });

      this.logger.log(`Found ${images.length} images for opportunity ${opportunityId}`);

      // Generate HTML content
      const htmlContent = this.generateHtmlContent(opportunityId, surveyData, imagesByField);
      
      const htmlFilePath = path.join(reportDir, `Survey_Report_${opportunityId}.html`);
      await writeFile(htmlFilePath, htmlContent, 'utf8');
      
      this.logger.log(`Generated HTML survey report: ${htmlFilePath}`);
      return htmlFilePath;
    } catch (error) {
      this.logger.error(`Failed to generate HTML survey report: ${error.message}`);
      throw error;
    }
  }

  private generateHtmlContent(opportunityId: string, surveyData: any, imagesByField: { [key: string]: any[] }): string {
    const currentDate = new Date().toLocaleString();
    
    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Solar Survey Report - ${opportunityId}</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f8f9fa;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: white;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }
        
        .header {
            text-align: center;
            padding: 30px 0;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            margin-bottom: 30px;
            border-radius: 10px;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header .subtitle {
            font-size: 1.2em;
            opacity: 0.9;
        }
        
        .report-info {
            background-color: #e9ecef;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            border-left: 4px solid #667eea;
        }
        
        .report-info h2 {
            color: #495057;
            margin-bottom: 10px;
        }
        
        .page-section {
            margin-bottom: 40px;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            overflow: hidden;
        }
        
        .page-header {
            background-color: #f8f9fa;
            padding: 15px 20px;
            border-bottom: 1px solid #dee2e6;
        }
        
        .page-header h3 {
            color: #495057;
            font-size: 1.4em;
            margin: 0;
        }
        
        .page-content {
            padding: 20px;
        }
        
        .field-group {
            margin-bottom: 20px;
        }
        
        .field-label {
            font-weight: 600;
            color: #495057;
            margin-bottom: 5px;
            display: block;
        }
        
        .field-value {
            background-color: #f8f9fa;
            padding: 10px;
            border-radius: 4px;
            border-left: 3px solid #667eea;
            margin-bottom: 10px;
        }
        
        .images-container {
            margin-top: 15px;
        }
        
        .image-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 10px;
        }
        
        .image-item {
            text-align: center;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 10px;
            background-color: white;
        }
        
        .image-item img {
            max-width: 100%;
            height: auto;
            max-height: 200px;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .image-item iframe {
            width: 100%;
            height: 300px;
            border: 1px solid #dee2e6;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .pdf-link {
            display: inline-block;
            padding: 10px 20px;
            background-color: #667eea;
            color: white;
            text-decoration: none;
            border-radius: 4px;
            margin-top: 10px;
            transition: background-color 0.3s;
        }
        
        .pdf-link:hover {
            background-color: #5568d3;
        }
        
        .image-caption {
            font-size: 0.9em;
            color: #6c757d;
            margin-top: 8px;
        }
        
        .no-data {
            color: #6c757d;
            font-style: italic;
            text-align: center;
            padding: 20px;
        }
        
        .footer {
            text-align: center;
            padding: 30px 0;
            color: #6c757d;
            border-top: 1px solid #dee2e6;
            margin-top: 40px;
        }
        
        @media print {
            body { background-color: white; }
            .container { box-shadow: none; }
            .page-section { break-inside: avoid; }
        }
        
        @media (max-width: 768px) {
            .container { padding: 10px; }
            .header h1 { font-size: 2em; }
            .image-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🌞 Solar Survey Report</h1>
            <div class="subtitle">Comprehensive Site Assessment & Customer Information</div>
        </div>
        
        <div class="report-info">
            <h2>📋 Report Information</h2>
            <p><strong>Opportunity ID:</strong> ${opportunityId}</p>
            <p><strong>Generated:</strong> ${currentDate}</p>
            <p><strong>Report Type:</strong> Complete Site Survey Assessment</p>
            ${this.getCustomerAndSurveyorInfo(surveyData)}
        </div>

        ${this.generatePageSections(surveyData, imagesByField)}
        
        <div class="footer">
            <p>This report was automatically generated by the Creative Solar Survey System</p>
            <p>For questions or support, please contact your solar consultant</p>
        </div>
    </div>
</body>
</html>`;
  }

  private generatePageSections(surveyData: any, imagesByField: { [key: string]: any[] }): string {
    const pageConfigs = [
      {
        key: 'page1',
        title: '📅 Appointment & Customer Details',
        icon: '👤',
        fields: [
          { key: 'date', label: 'Survey Date' },
          { key: 'postcode', label: 'Postcode' },
          { key: 'addressLine1', label: 'Address' },
          { key: 'customerFirstName', label: 'Customer First Name' },
          { key: 'customerLastName', label: 'Customer Last Name' },
          { key: 'appointmentDateTime', label: 'Appointment Date & Time' },
          { key: 'homeOwnersAvailable', label: 'Homeowners Available' },
          { key: 'renewableExecutiveFirstName', label: 'Executive First Name' },
          { key: 'renewableExecutiveLastName', label: 'Executive Last Name' }
        ]
      },
      {
        key: 'page2',
        title: '🎯 Customer Motivation',
        icon: '💡',
        fields: [
          { key: 'selectedReasons', label: 'Reasons for Solar Interest' }
        ]
      },
      {
        key: 'page3',
        title: '🏠 Property Information',
        icon: '🏡',
        fields: [
          { key: 'bedrooms', label: 'Number of Bedrooms' },
          { key: 'property', label: 'Property Type' },
          { key: 'movingPlans', label: 'Moving Plans' },
          { key: 'propertyType', label: 'Property Style' }
        ]
      },
      {
        key: 'page4',
        title: '⚡ Energy & Heating Details',
        icon: '🔌',
        fields: [
          { key: 'phaseMeter', label: 'Phase Meter' },
          { key: 'heatingType', label: 'Heating Type' },
          { key: 'energyCompany', label: 'Energy Company' },
          { key: 'additionalFeatures', label: 'Additional Features' },
          { key: 'annualElectricUsage', label: 'Annual Electric Usage' },
          { key: 'electricPricePerUnit', label: 'Electric Price Per Unit' },
          { key: 'monthlyElectricSpend', label: 'Monthly Electric Spend' }
        ],
        imageFields: ['energyBill']
      },
      {
        key: 'page5',
        title: '📄 EPC Certificate',
        icon: '📋',
        imageFields: ['epcCertificate']
      },
      {
        key: 'page6',
        title: '🔋 Battery & Storage',
        icon: '🔋',
        fields: [
          { key: 'batteryLocation', label: 'Battery Location' },
          { key: 'batterySize', label: 'Battery Size' }
        ],
        imageFields: ['batteryInverterLocation']
      },
      {
        key: 'page7',
        title: '📸 Property Photos',
        icon: '📷',
        fields: [
          { key: 'roofTileType', label: 'Roof Tile Type' },
          { key: 'roofType', label: 'Roof Type' },
          { key: 'hasSolarBattery', label: 'Solar/Battery Storage' }
        ],
        imageFields: ['frontDoor', 'frontProperty', 'targetRoofs', 'roofAngle', 'roofTileCloseup', 'electricMeter', 'fuseBoard']
      },
      {
        key: 'page8',
        title: '🔧 Installation Details',
        icon: '⚙️',
        fields: [
          { key: 'optimiserDetails', label: 'Optimiser Details' },
          { key: 'furtherInformation', label: 'Further Information' },
          { key: 'scaffoldingRequired', label: 'Scaffolding Required' }
        ],
        imageFields: ['scaffolding', 'shadingIssues', 'evLocation', 'evCharger']
      },
      {
        key: 'page9',
        title: '📝 Signatures & Additional Photos',
        icon: '✍️',
        imageFields: ['customerSignature', 'renewableExecutiveSignature', 'otherRoofPictures', 'otherBuildings', 'garage', 'propertySides', 'batteryInverterLocation']
      }
    ];

    let html = '';
    
    for (const pageConfig of pageConfigs) {
      const pageData = surveyData[pageConfig.key];
      if (!pageData) continue;

      html += `
        <div class="page-section">
          <div class="page-header">
            <h3>${pageConfig.icon} ${pageConfig.title}</h3>
          </div>
          <div class="page-content">
      `;

      // Add text fields
      if (pageConfig.fields) {
        for (const field of pageConfig.fields) {
          const value = pageData[field.key];
          if (value !== null && value !== undefined && value !== '') {
            html += `
              <div class="field-group">
                <span class="field-label">${field.label}:</span>
                <div class="field-value">${this.formatFieldValue(value)}</div>
              </div>
            `;
          }
        }
      }

      // Add image fields
      if (pageConfig.imageFields) {
        for (const imageField of pageConfig.imageFields) {
          const images = imagesByField[imageField] || [];
          if (images.length > 0) {
            html += `
              <div class="field-group">
                <span class="field-label">${this.formatFieldLabel(imageField)}:</span>
                <div class="images-container">
                  <div class="image-grid">
            `;
            
            for (const image of images) {
              // Use Cloudinary URL directly since images are stored there
              const imageUrl = image.filePath; // This contains the Cloudinary URL
              const isPdf = image.mimeType === 'application/pdf';
              
              if (isPdf) {
                // For PDFs, display as embedded PDF or download link
                html += `
                  <div class="image-item">
                    <iframe src="${imageUrl}" type="application/pdf" style="width: 100%; height: 300px; border: 1px solid #dee2e6; border-radius: 4px;"></iframe>
                    <div class="image-caption">
                      <a href="${imageUrl}" target="_blank" class="pdf-link">📄 View/Download PDF: ${image.originalName}</a>
                    </div>
                  </div>
                `;
              } else {
                // For images, display as image
                html += `
                  <div class="image-item">
                    <img src="${imageUrl}" alt="${image.originalName}" onerror="this.style.display='none'; this.nextElementSibling.style.display='block';" />
                    <div class="image-caption" style="display:none;">Image: ${image.originalName}</div>
                    <div class="image-caption">${image.originalName}</div>
                  </div>
                `;
              }
            }
            
            html += `
                  </div>
                </div>
              </div>
            `;
          }
        }
      }

      // Check if page has no data
      const hasTextData = pageConfig.fields?.some(field => {
        const value = pageData[field.key];
        return value !== null && value !== undefined && value !== '';
      });
      
      const hasImageData = pageConfig.imageFields?.some(field => {
        return (imagesByField[field] || []).length > 0;
      });

      if (!hasTextData && !hasImageData) {
        html += `<div class="no-data">No data provided for this section</div>`;
      }

      html += `
          </div>
        </div>
      `;
    }

    return html;
  }

  private formatFieldValue(value: any): string {
    if (Array.isArray(value)) {
      return value.join(', ');
    }
    if (typeof value === 'object' && value !== null) {
      return JSON.stringify(value, null, 2);
    }
    return String(value);
  }

  private formatFieldLabel(fieldName: string): string {
    return fieldName
      .replace(/([A-Z])/g, ' $1')
      .replace(/^./, str => str.toUpperCase())
      .replace(/Files$/, ' Photos')
      .replace(/Files$/, ' Images');
  }

  private getCustomerAndSurveyorInfo(surveyData: any): string {
    const page1 = surveyData.page1 || {};
    
    const customerFirstName = page1.customerFirstName || '';
    const customerLastName = page1.customerLastName || '';
    const customerName = customerFirstName && customerLastName 
      ? `${customerFirstName} ${customerLastName}`.trim()
      : customerFirstName || customerLastName || 'Not provided';
    
    const surveyorFirstName = page1.renewableExecutiveFirstName || '';
    const surveyorLastName = page1.renewableExecutiveLastName || '';
    const surveyorName = surveyorFirstName && surveyorLastName
      ? `${surveyorFirstName} ${surveyorLastName}`.trim()
      : surveyorFirstName || surveyorLastName || 'Not provided';
    
    return `
            <p><strong>Customer Name:</strong> ${customerName}</p>
            <p><strong>Surveyor (Rep):</strong> ${surveyorName}</p>
    `;
  }

  async generatePdfReport(opportunityId: string, surveyData: any, surveyId?: string): Promise<string> {
    try {
      // First generate HTML report
      const htmlFilePath = await this.generateHtmlReport(opportunityId, surveyData, surveyId);
      
      // Convert HTML to PDF using Puppeteer
      const reportDir = path.join(this.uploadsDir, opportunityId);
      const pdfFilePath = path.join(reportDir, `Survey_Report_${opportunityId}.pdf`);
      
      this.logger.log(`Converting HTML to PDF for opportunity ${opportunityId}`);
      
      const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
      });
      
      try {
        const page = await browser.newPage();
        
        // Load the HTML file
        await page.goto(`file://${htmlFilePath}`, { 
          waitUntil: 'networkidle0',
          timeout: 30000 
        });
        
        // Wait for images to load - increased timeout for Cloudinary images
        await new Promise(resolve => setTimeout(resolve, 5000));
        
        // Generate PDF
        await page.pdf({
          path: pdfFilePath,
          format: 'A4',
          printBackground: true,
          margin: {
            top: '20mm',
            right: '20mm',
            bottom: '20mm',
            left: '20mm'
          },
          displayHeaderFooter: true,
          headerTemplate: '<div style="font-size: 10px; text-align: center; width: 100%; color: #666;">Solar Survey Report - ${opportunityId}</div>',
          footerTemplate: '<div style="font-size: 10px; text-align: center; width: 100%; color: #666;">Page <span class="pageNumber"></span> of <span class="totalPages"></span></div>'
        });
        
        this.logger.log(`Successfully generated PDF report: ${pdfFilePath}`);
        
      } finally {
        await browser.close();
      }
      
      return pdfFilePath;
    } catch (error) {
      this.logger.error(`Failed to generate PDF report: ${error.message}`);
      throw error;
    }
  }

  async getReportPath(opportunityId: string, format: 'html' | 'pdf' = 'html'): Promise<string | null> {
    try {
      const reportDir = path.join(this.uploadsDir, opportunityId);
      const fileName = `Survey_Report_${opportunityId}.${format}`;
      const filePath = path.join(reportDir, fileName);
      
      if (fs.existsSync(filePath)) {
        return filePath;
      }
      
      return null;
    } catch (error) {
      this.logger.error(`Failed to get report path: ${error.message}`);
      return null;
    }
  }
}