import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';

@Injectable()
export class UltraSimpleVideoService {
  private readonly logger = new Logger(UltraSimpleVideoService.name);
  private readonly outputDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'output');
  private readonly publicDir = path.join(process.cwd(), 'public', 'videos');

  constructor() {
    // Ensure public directory exists
    if (!fs.existsSync(this.publicDir)) {
      fs.mkdirSync(this.publicDir, { recursive: true });
    }
  }

  /**
   * Ultra simple approach: Just serve the PDF as a "video" placeholder
   * This allows the mobile app to display the PDF in a video player interface
   * while we work on proper video conversion
   */
  async generateVideoFromPdf(pdfPath: string, opportunityId: string, customerName: string): Promise<{
    success: boolean;
    videoPath?: string;
    publicUrl?: string;
    error?: string;
  }> {
    try {
      this.logger.log(`Creating video placeholder for PDF: ${pdfPath}`);

      // Check if PDF exists
      if (!fs.existsSync(pdfPath)) {
        throw new Error(`PDF file not found: ${pdfPath}`);
      }

      const timestamp = Date.now();
      const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      
      // For now, we'll create a simple HTML file that can be served as a "video"
      // This is a temporary solution until we implement proper PDF to video conversion
      const htmlFilename = `proposal_${safeCustomerName}_${opportunityId}_${timestamp}.html`;
      const htmlPath = path.join(this.outputDir, htmlFilename);
      const publicHtmlPath = path.join(this.publicDir, htmlFilename);

      // Create a simple HTML viewer for the PDF
      const htmlContent = this.createPdfViewerHtml(pdfPath, customerName, opportunityId);
      
      fs.writeFileSync(htmlPath, htmlContent);
      fs.writeFileSync(publicHtmlPath, htmlContent);

      const publicUrl = `/videos/${htmlFilename}`;
      
      this.logger.log(`PDF viewer created successfully: ${htmlPath}`);
      this.logger.log(`Public URL: ${publicUrl}`);

      return {
        success: true,
        videoPath: htmlPath,
        publicUrl: publicUrl,
      };

    } catch (error) {
      this.logger.error(`PDF viewer creation failed: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Create a simple HTML viewer for the PDF
   */
  private createPdfViewerHtml(pdfPath: string, customerName: string, opportunityId: string): string {
    const pdfFilename = path.basename(pdfPath);
    const pdfUrl = `/presentation/download/${pdfFilename}`;
    
    return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Solar Proposal - ${customerName}</title>
    <style>
        body {
            margin: 0;
            padding: 20px;
            font-family: Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: white;
            border-radius: 20px;
            padding: 40px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            max-width: 800px;
            width: 100%;
            text-align: center;
        }
        h1 {
            color: #333;
            margin-bottom: 20px;
            font-size: 2.5em;
        }
        .subtitle {
            color: #666;
            margin-bottom: 30px;
            font-size: 1.2em;
        }
        .pdf-container {
            margin: 30px 0;
            border: 2px solid #ddd;
            border-radius: 10px;
            overflow: hidden;
        }
        iframe {
            width: 100%;
            height: 600px;
            border: none;
        }
        .download-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 30px;
            border: none;
            border-radius: 10px;
            font-size: 1.1em;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            margin: 20px 10px;
            transition: transform 0.2s;
        }
        .download-btn:hover {
            transform: translateY(-2px);
        }
        .info {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            color: #666;
        }
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                padding: 20px;
            }
            h1 {
                font-size: 2em;
            }
            iframe {
                height: 400px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🌞 Solar Proposal</h1>
        <div class="subtitle">For ${customerName}</div>
        
        <div class="info">
            <strong>Opportunity ID:</strong> ${opportunityId}<br>
            <strong>Generated:</strong> ${new Date().toLocaleDateString()}
        </div>
        
        <div class="pdf-container">
            <iframe src="${pdfUrl}" type="application/pdf">
                <p>Your browser doesn't support PDF viewing. 
                <a href="${pdfUrl}" class="download-btn">Download PDF</a></p>
            </iframe>
        </div>
        
        <div>
            <a href="${pdfUrl}" class="download-btn" download>📥 Download PDF</a>
            <a href="javascript:window.print()" class="download-btn">🖨️ Print</a>
        </div>
        
        <div class="info">
            <small>Generated by Creativ Solar - Professional Solar Solutions</small>
        </div>
    </div>
    
    <script>
        // Auto-refresh if PDF fails to load
        setTimeout(() => {
            const iframe = document.querySelector('iframe');
            if (iframe && iframe.contentDocument && iframe.contentDocument.body.innerHTML === '') {
                console.log('PDF failed to load, showing download option');
            }
        }, 3000);
    </script>
</body>
</html>`;
  }

  /**
   * Alternative: Create a simple video placeholder that shows PDF pages as images
   * This is a middle-ground approach that doesn't require external tools
   */
  async createSimpleVideoPlaceholder(pdfPath: string, opportunityId: string, customerName: string): Promise<{
    success: boolean;
    videoPath?: string;
    publicUrl?: string;
    error?: string;
  }> {
    // For now, we'll use the HTML viewer approach
    // In the future, this could be enhanced to:
    // 1. Use a PDF.js library to render PDF pages as images
    // 2. Create a simple slideshow video from those images
    // 3. Use a web-based PDF to video conversion service
    
    return this.generateVideoFromPdf(pdfPath, opportunityId, customerName);
  }
}
