import { Injectable, Logger } from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import * as fsPromises from 'fs/promises';
import { promisify } from 'util';
import * as XLSX from 'xlsx';
import * as PizZip from 'pizzip';
import { exec, spawn } from 'child_process';
import { SessionManagementService } from '../session-management/session-management.service';
import { ComProcessManagerService } from '../session-management/com-process-manager.service';
import { PowerpointMp4Service } from '../video-generation/powerpoint-mp4.service';
import { FfmpegVideoService } from '../video-generation/ffmpeg-video.service';
import { ImageGenerationService } from '../video-generation/image-generation.service';
import { ExcelAutomationService } from '../excel-automation/excel-automation.service';

@Injectable()
export class PresentationService {
  private readonly logger = new Logger(PresentationService.name);
  private readonly presentationTemplatePath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'presnetation', 'Proposal.pptx');
  private readonly outputDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'output');
  private readonly opportunitiesDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
  private readonly epvsOpportunitiesDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');
  private readonly execAsync = promisify(exec);

  constructor(
    private readonly sessionManagementService: SessionManagementService,
    private readonly comProcessManagerService: ComProcessManagerService,
    private readonly powerpointMp4Service: PowerpointMp4Service,
    private readonly ffmpegVideoService: FfmpegVideoService,
    private readonly imageGenerationService: ImageGenerationService,
    private readonly excelAutomationService: ExcelAutomationService
  ) {
    // Ensure output directory exists
    if (!fs.existsSync(this.outputDir)) {
  
      fs.mkdirSync(this.outputDir, { recursive: true });
    }
  }

  /**
   * Generate presentation with user session isolation
   */
  async generatePresentationWithSession(
    userId: string,
    data: {
      opportunityId: string;
      calculatorType?: 'flux' | 'off-peak' | 'epvs';
      customerName?: string;
      date?: string;
      postcode?: string;
      solarData?: any;
    }
  ) {
    this.logger.log(`Generating presentation with session isolation for user: ${userId}, opportunity: ${data.opportunityId}`);

    try {
      // Queue the request through session management
      const result = await this.sessionManagementService.queueRequest(
        userId,
        'powerpoint_generation',
        'com',
        data,
        1 // High priority
      );

      return result;
    } catch (error) {
      this.logger.error(`Session-based presentation generation failed for user ${userId}:`, error);
      throw error;
    }
  }

  /**
   * Generate MP4 video presentation using FFmpeg pipeline (more efficient than COM automation)
   */
  async generateVideoPresentation(data: {
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      this.logger.log(`Generating MP4 video presentation for opportunity: ${data.opportunityId}`);

      // Generate PowerPoint presentation without PDF conversion
      const presentationResult = await this.generatePowerPointOnly(data);
      
      // Check if PowerPoint was generated successfully
      if (!presentationResult.pptxFile) {
        return {
          success: false,
          message: 'Failed to generate PowerPoint presentation',
          error: 'PowerPoint generation failed',
        };
      }

      // Extract customer data for video generation
      const customerData = await this.extractCustomerDataFromExcel(data.opportunityId, data.calculatorType);
      const finalCustomerName = data.customerName || customerData.customerName || 'Customer';

      // Get the full PowerPoint path
      const pptxPath = path.join(this.outputDir, presentationResult.pptxFile);

      // Use the generated PowerPoint directly for video processing (don't overwrite the template)
      // The template should remain untouched with variable placeholders

      // Generate video using FFmpeg pipeline (replaces the old PowerPoint COM automation)
      this.logger.log(`🎬 Converting PowerPoint to MP4 using FFmpeg: ${pptxPath}`);
      const videoResult = await this.ffmpegVideoService.generateProposalVideo({
        opportunityId: data.opportunityId,
        customerName: finalCustomerName,
        date: data.date || new Date().toISOString().split('T')[0],
        postcode: data.postcode || '',
        solarData: data.solarData || {},
        pptxPath: pptxPath, // Pass the generated PowerPoint path
      });

      if (videoResult.success) {
        return {
          success: true,
          message: 'MP4 video presentation generated successfully',
          data: {
            videoPath: videoResult.videoPath,
            publicUrl: videoResult.publicUrl,
            filename: path.basename(videoResult.videoPath || ''),
            opportunityId: data.opportunityId,
            customerName: finalCustomerName,
            calculatorType: data.calculatorType || 'off-peak',
            pptxPath: pptxPath, // Keep PowerPoint path as well
            pptxFile: presentationResult.pptxFile,
          },
        };
      } else {
        return {
          success: false,
          message: 'MP4 video conversion failed',
          error: videoResult.error,
        };
      }
    } catch (error) {
      this.logger.error(`MP4 video presentation generation error: ${error.message}`);
      return {
        success: false,
        message: 'MP4 video presentation generation error',
        error: error.message,
      };
    }
  }

  /**
   * Generate image presentation by exporting all slides as PNG images
   */
  async generateImagePresentation(data: {
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      this.logger.log(`Generating image presentation for opportunity: ${data.opportunityId}`);

      // Generate PowerPoint presentation without PDF conversion
      const presentationResult = await this.generatePowerPointOnly(data);
      
      // Check if PowerPoint was generated successfully
      if (!presentationResult.pptxFile) {
        return {
          success: false,
          message: 'Failed to generate PowerPoint presentation',
          error: 'PowerPoint generation failed',
        };
      }

      // Extract customer data for image generation
      const customerData = await this.extractCustomerDataFromExcel(data.opportunityId, data.calculatorType);
      const finalCustomerName = data.customerName || customerData.customerName || 'Customer';

      // Get the full PowerPoint path
      const pptxPath = path.join(this.outputDir, presentationResult.pptxFile);

      // Generate images using image generation service
      this.logger.log(`🖼️ Converting PowerPoint to images: ${pptxPath}`);
      const imageResult = await this.imageGenerationService.generateProposalImages({
        opportunityId: data.opportunityId,
        customerName: finalCustomerName,
        date: data.date || new Date().toISOString().split('T')[0],
        postcode: data.postcode || '',
        solarData: data.solarData || {},
        pptxPath: pptxPath,
      });

      if (imageResult.success && imageResult.publicUrls) {
        return {
          success: true,
          message: 'Image presentation generated successfully',
          data: {
            images: imageResult.images,
            publicUrls: imageResult.publicUrls,
            opportunityId: data.opportunityId,
            customerName: finalCustomerName,
            calculatorType: data.calculatorType || 'off-peak',
            pptxPath: pptxPath,
            pptxFile: presentationResult.pptxFile,
          },
        };
      } else {
        return {
          success: false,
          message: 'Image conversion failed',
          error: imageResult.error,
        };
      }
    } catch (error) {
      this.logger.error(`Image presentation generation error: ${error.message}`);
      return {
        success: false,
        message: 'Image presentation generation error',
        error: error.message,
      };
    }
  }

  /**
   * Generate PowerPoint presentation only (without PDF conversion)
   */
  async generatePowerPointOnly(data: {
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      // Create a unique filename
      const timestamp = Date.now();
      const outputFilename = `presentation_${data.opportunityId}_${timestamp}.pptx`;
      
      const outputPath = path.join(this.outputDir, outputFilename);

      // Extract customer data from Excel file based on calculator type
      const customerData = await this.extractCustomerDataFromExcel(data.opportunityId, data.calculatorType);
      
      // Merge with provided data (provided data takes precedence)
      const finalData = {
        ...customerData,
        customerName: data.customerName || customerData.customerName,
        date: data.date || customerData.date,
        postcode: data.postcode || customerData.postcode,
        ...data.solarData
      };

      // Generate presentation with replaced variables
      await this.replaceVariablesInPresentation(this.presentationTemplatePath, outputPath, finalData, data.opportunityId, data.calculatorType || 'off-peak');

      return {
        pptxFile: outputFilename,
        downloadUrl: `/presentation/download/${outputFilename}`,
        pptxViewUrl: `/presentation/view/${outputFilename}`,
        message: 'PowerPoint presentation generated successfully.',
        variables: finalData
      };

    } catch (error) {
      this.logger.error(`PowerPoint generation error: ${error.message}`);
      throw error;
    }
  }

  async generatePresentation(data: {
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      // Create a unique filename
      const timestamp = Date.now();
      const outputFilename = `presentation_${data.opportunityId}_${timestamp}.pptx`;
      const pdfFilename = `presentation_${data.opportunityId}_${timestamp}.pdf`;
      
      const outputPath = path.join(this.outputDir, outputFilename);
      const pdfPath = path.join(this.outputDir, pdfFilename);

      // Extract customer data from Excel file based on calculator type
      const customerData = await this.extractCustomerDataFromExcel(data.opportunityId, data.calculatorType);
      
      // Merge with provided data (provided data takes precedence)
      const finalData = {
        ...customerData,
        customerName: data.customerName || customerData.customerName,
        date: data.date || customerData.date,
        postcode: data.postcode || customerData.postcode,
        ...data.solarData
      };

      // Generate presentation with replaced variables
      await this.replaceVariablesInPresentation(this.presentationTemplatePath, outputPath, finalData, data.opportunityId, data.calculatorType || 'off-peak');
      
      // Convert to PDF with improved error handling
      let pdfGenerated = false;
      let pdfMessage = '';
      
      try {
        pdfGenerated = await this.convertToPdf(outputPath, pdfPath);
        pdfMessage = pdfGenerated ? 'PDF generated successfully.' : 'PDF generation failed, but PowerPoint file is available.';
      } catch (pdfError) {
        console.error('PDF conversion error:', pdfError);
        pdfMessage = `PDF generation failed: ${pdfError.message}. PowerPoint file is still available for download.`;
      }

      return {
        pptxFile: outputFilename,
        pdfFile: pdfGenerated ? pdfFilename : null,
        downloadUrl: `/presentation/download/${outputFilename}`,
        pdfDownloadUrl: pdfGenerated ? `/presentation/download/${pdfFilename}` : null,
        pdfViewUrl: pdfGenerated ? `/presentation/view/${pdfFilename}` : null,
        pptxViewUrl: `/presentation/view/${outputFilename}`, // Add PowerPoint view URL
        pdfGenerated,
        message: `Presentation generated successfully with data from Excel file. ${pdfMessage}`,
        variables: finalData
      };
    } catch (error) {
      console.error('Presentation generation error:', error);
      throw error;
    }
  }

  private async extractCustomerDataFromExcel(opportunityId: string, calculatorType: 'flux' | 'off-peak' | 'epvs' = 'off-peak') {
    try {
      console.log(`🔍 extractCustomerDataFromExcel called with calculatorType: ${calculatorType}`);
      
      // Use robust finder to locate the correct/latest Excel file for this opportunity
      const excelFilePath = await this.findOpportunityExcelFile(opportunityId, calculatorType);
      if (!excelFilePath) {
        throw new Error(`${calculatorType === 'flux' || calculatorType === 'epvs' ? 'EPVS' : 'Off Peak'} opportunity Excel file not found for ID: ${opportunityId}`);
      }
      console.log(`Reading ${calculatorType === 'flux' || calculatorType === 'epvs' ? 'EPVS/Flux' : 'Off Peak'} Excel file: ${excelFilePath}`);

      // Read the Excel file
      const workbook = XLSX.readFile(excelFilePath);
      const inputsSheet = workbook.Sheets['Inputs'];

      if (!inputsSheet) {
        throw new Error('Inputs sheet not found in Excel file');
      }

      // Extract year values from Solar Projections sheet
      const yearValues = this.extractYearValues(workbook, calculatorType);
      
      // Extract customer data from specific cells based on calculator type
      let customerData: any;
      
      if (calculatorType === 'flux' || calculatorType === 'epvs') {
        // EPVS/Flux calculator cell references
        customerData = {
          customerName: this.getCellValue(inputsSheet, 'H11') || '', // Customer Name
          address: this.getCellValue(inputsSheet, 'H12') || '', // Address
          postcode: this.getCellValue(inputsSheet, 'H13') || '', // Postcode
          date: new Date().toLocaleDateString('en-GB'),
          // Solar system data for EPVS/Flux
          p_w: this.extractPanelWattage(this.getCellValue(inputsSheet, 'H43')), // Panel wattage
          p_q: this.countPanels(inputsSheet, 'C70', 'C77'), // Panel quantity from C70...C77
          i_s: this.extractInverterSize(this.getCellValue(inputsSheet, 'H54')), // Inverter size
          b_s: this.extractBatterySize(this.getCellValue(inputsSheet, 'H47')), // Battery size
          t_y_s_o: this.formatNumber(this.getCellValue(inputsSheet, 'E78')), // Total year first solar output
          t_y_s_g: this.getCellValue(workbook.Sheets['Solar Projections'] || {}, 'G3'), // Total year first solar generation
          // New proposal variables from Solar Projections sheet (formatted to 2 decimal places)
          term_var: this.getCellValue(workbook.Sheets['Solar Projections'] || {}, 'B10') || '', // Term variable from B10
          payment_type: (() => {
            const b7Value = this.getCellValue(workbook.Sheets['Solar Projections'] || {}, 'B7');
            console.log(`🔍 DEBUG payment_type from B7 (EPVS/Flux - Solar Projections): "${b7Value}"`);
            return b7Value || 'Cash';
          })(), // Payment type from B7 in Solar Projections sheet
          yearly_plan_cost: this.formatCurrency(this.getCellValue(workbook.Sheets['Solar Projections'] || {}, 'B14')), // Yearly plan cost from B14
          yearly_saving: this.formatCurrency(this.getCellValue(workbook.Sheets['Solar Projections'] || {}, 'B20')), // Yearly saving from B20
          yearly_con: (() => {
            const rawValue = this.getCellValue(workbook.Sheets['Solar Projections'] || {}, 'D20');
            console.log(`🔍 DEBUG yearly_con from D20: raw="${rawValue}", formatted="${this.formatCurrency(rawValue)}"`);
            
            // Always return £0.00 for yearly contribution (no contribution required)
            console.log(`🔍 DEBUG yearly_con set to £0.00 (no contribution required)`);
            return '£0.00';
          })(), // Yearly con always £0.00 (no contribution required)
          your_LT_profit: this.formatCurrency(this.getCellValue(workbook.Sheets['Solar Projections'] || {}, 'B25')), // Your LT profit from B25
          // Year values for PowerPoint
          ...yearValues
        };
      } else {
        // Off Peak calculator cell references
        customerData = {
          customerName: this.getCellValue(inputsSheet, 'H12') || '', // Customer Name
          address: this.getCellValue(inputsSheet, 'H13') || '', // Address
          postcode: this.getCellValue(inputsSheet, 'H14') || '', // Postcode
          date: new Date().toLocaleDateString('en-GB'),
          // Solar system data for Off Peak
          p_w: this.extractPanelWattage(this.getCellValue(inputsSheet, 'H42')), // Panel wattage
          p_q: this.countPanels(inputsSheet, 'C69', 'C76'), // Panel quantity from C69...C76
          i_s: this.extractInverterSize(this.getCellValue(inputsSheet, 'H53')), // Inverter size
          b_s: this.extractBatterySize(this.getCellValue(inputsSheet, 'H46')), // Battery size
          t_y_s_o: (() => {
            const rawValue = this.getCellValue(inputsSheet, 'E77');
            console.log(`🔍 t_y_s_o raw value from E77: "${rawValue}" (type: ${typeof rawValue})`);
            // Convert kW to W by multiplying by 1000
            const numValue = parseFloat(rawValue);
            const wattsValue = numValue * 1000;
            console.log(`🔍 t_y_s_o converted from ${numValue} kW to ${wattsValue} W`);
            const formatted = Math.round(wattsValue).toString();
            console.log(`🔍 t_y_s_o formatted value: "${formatted}"`);
            return formatted;
          })(), // Total year first solar output
          t_y_s_g: this.getCellValue(workbook.Sheets['Solar Projection'] || workbook.Sheets['Solar Projections'] || {}, 'G3'), // Total year first solar generation
          // New proposal variables from Solar Projections sheet (formatted to 2 decimal places)
          term_var: this.getCellValue(workbook.Sheets['Solar Projection'] || workbook.Sheets['Solar Projections'] || {}, 'B10') || '', // Term variable from B10
          payment_type: (() => {
            const b7Value = this.getCellValue(workbook.Sheets['Solar Projection'] || workbook.Sheets['Solar Projections'] || {}, 'B7');
            console.log(`🔍 DEBUG payment_type from B7 (Off Peak - Solar Projections): "${b7Value}"`);
            return b7Value || 'Cash';
          })(), // Payment type from B7 in Solar Projections sheet
          yearly_plan_cost: this.formatCurrency(this.getCellValue(workbook.Sheets['Solar Projection'] || workbook.Sheets['Solar Projections'] || {}, 'B14')), // Yearly plan cost from B14
          yearly_saving: this.formatCurrency(this.getCellValue(workbook.Sheets['Solar Projection'] || workbook.Sheets['Solar Projections'] || {}, 'B20')), // Yearly saving from B20
          yearly_con: (() => {
            const rawValue = this.getCellValue(workbook.Sheets['Solar Projection'] || workbook.Sheets['Solar Projections'] || {}, 'D20');
            console.log(`🔍 DEBUG yearly_con from D20 (Off Peak): raw="${rawValue}", formatted="${this.formatCurrency(rawValue)}"`);
            
            // Always return £0.00 for yearly contribution (no contribution required)
            console.log(`🔍 DEBUG yearly_con set to £0.00 (no contribution required)`);
            return '£0.00';
          })(), // Yearly con always £0.00 (no contribution required)
          your_LT_profit: this.formatCurrency(this.getCellValue(workbook.Sheets['Solar Projection'] || workbook.Sheets['Solar Projections'] || {}, 'B25')), // Your LT profit from B25
          // Year values for PowerPoint
          ...yearValues
        };
      }

      console.log(`Extracted ${calculatorType === 'flux' || calculatorType === 'epvs' ? 'EPVS/Flux' : 'Off Peak'} customer data:`, customerData);
      return customerData;

    } catch (error) {
      console.error('Error extracting customer data from Excel:', error);
      // Return default data if Excel reading fails
      return {
        customerName: '',
        address: '',
        postcode: '',
        date: new Date().toLocaleDateString('en-GB'),
        p_w: '',
        p_q: 0,
        i_s: '',
        b_s: '',
        t_y_s_o: '',
        t_y_s_g: '',
        year_1: '',
        year_10: '',
        year_25: ''
      };
    }
  }

  private getCellValue(sheet: any, cellAddress: string): string {
    const cell = sheet[cellAddress];
    return cell ? String(cell.v || '') : '';
  }

  private extractYearValues(workbook: any, calculatorType: 'flux' | 'off-peak' | 'epvs'): any {
    try {
      console.log(`🔍 Extracting year values for ${calculatorType} calculator...`);
      
      // Find the Solar Projections sheet
      const solarProjectionsSheet = workbook.SheetNames.find((name: string) => 
        name.toLowerCase().includes('solar') && name.toLowerCase().includes('projection')
      );
      
      if (!solarProjectionsSheet) {
        console.log('⚠️ Solar Projections sheet not found');
        return {
          year_1: '',
          year_10: '',
          year_25: ''
        };
      }
      
      console.log(`✅ Found sheet: "${solarProjectionsSheet}"`);
      
      // Get the worksheet
      const worksheet = workbook.Sheets[solarProjectionsSheet];
      
      // Extract year 1 value (stays the same)
      const year1Value = worksheet['N3'] ? worksheet['N3'].v : null;
      
      // Calculate year 10: sum of N3 to N12 (years 1-10)
      let year10Sum = 0;
      for (let row = 3; row <= 12; row++) {
        const cellValue = worksheet[`N${row}`] ? worksheet[`N${row}`].v : null;
        if (cellValue && !isNaN(parseFloat(cellValue))) {
          year10Sum += parseFloat(cellValue);
          console.log(`  - N${row}: ${cellValue} (running sum: ${year10Sum})`);
        }
      }
      
      // Calculate year 25: sum of N3 to N27 (years 1-25)
      let year25Sum = 0;
      for (let row = 3; row <= 27; row++) {
        const cellValue = worksheet[`N${row}`] ? worksheet[`N${row}`].v : null;
        if (cellValue && !isNaN(parseFloat(cellValue))) {
          year25Sum += parseFloat(cellValue);
          console.log(`  - N${row}: ${cellValue} (running sum: ${year25Sum})`);
        }
      }
      
      console.log(`📊 Calculated year values:`);
      console.log(`  - Year 1 (N3): ${year1Value} (type: ${typeof year1Value})`);
      console.log(`  - Year 10 (N3-N12 sum): ${year10Sum}`);
      console.log(`  - Year 25 (N3-N27 sum): ${year25Sum}`);
      
      // Format the values based on calculator type
      const formattedValues = {
        year_1: this.formatYearValue(year1Value, calculatorType),
        year_10: this.formatYearValue(year10Sum, calculatorType),
        year_25: this.formatYearValue(year25Sum, calculatorType)
      };
      
      console.log(`🎯 Formatted year values:`, formattedValues);
      
      // Additional validation logging
      console.log(`🔍 Year value validation:`);
      console.log(`  - year_1: "${formattedValues.year_1}" (length: ${formattedValues.year_1.length})`);
      console.log(`  - year_10: "${formattedValues.year_10}" (length: ${formattedValues.year_10.length})`);
      console.log(`  - year_25: "${formattedValues.year_25}" (length: ${formattedValues.year_25.length})`);
      
      return formattedValues;
      
    } catch (error) {
      console.error('❌ Error extracting year values:', error);
      return {
        year_1: '',
        year_10: '',
        year_25: ''
      };
    }
  }

  private formatYearValue(value: any, calculatorType: 'flux' | 'off-peak' | 'epvs'): string {
    if (!value || value === null || value === undefined) {
      return '';
    }
    
    try {
      const numValue = typeof value === 'number' ? value : parseFloat(value);
      
      if (isNaN(numValue)) {
        return '';
      }
      
      // Format to 2 decimal places without rounding - use exact number
      // Convert to string and ensure exactly 2 decimal places
      const strValue = numValue.toString();
      const parts = strValue.split('.');
      
      if (parts.length === 1) {
        // No decimal point, add .00
        return `${parts[0]}.00`;
      } else {
        // Has decimal point, ensure exactly 2 decimal places
        const integerPart = parts[0];
        const decimalPart = parts[1];
        
        if (decimalPart.length === 1) {
          // Only 1 decimal place, add one more
          return `${integerPart}.${decimalPart}0`;
        } else if (decimalPart.length >= 2) {
          // 2 or more decimal places, take first 2 without rounding
          return `${integerPart}.${decimalPart.substring(0, 2)}`;
        } else {
          // Fallback to 2 decimal places
          return `${integerPart}.00`;
        }
      }
    } catch (error) {
      console.error('Error formatting year value:', error);
      return String(value);
    }
  }

  private extractPanelWattage(cellValue: string): string {
    // Extract wattage from dropdown text (e.g., "400W Panel" -> "400W", "CHSM54RN 450" -> "450W")
    if (!cellValue || cellValue.trim() === '') {
      this.logger.warn('Panel wattage cell is empty');
      return '';
    }
    
    this.logger.log(`Extracting panel wattage from: "${cellValue}"`);
    
    // Try multiple patterns to match different formats
    const patterns = [
      /(\d+(?:\.\d+)?)\s*W/i,           // "400W" or "400 W"
      /(\d+(?:\.\d+)?)\s*Wp/i,          // "400Wp" or "400 Wp"
      /(\d+(?:\.\d+)?)\s*Watt/i,        // "400Watt" or "400 Watt"
      /(\d+(?:\.\d+)?)\s*Watts/i,       // "400Watts" or "400 Watts"
      /(\d+(?:\.\d+)?)\s*W\s*Panel/i,   // "400W Panel"
      /(\d+(?:\.\d+)?)\s*Wp\s*Panel/i,  // "400Wp Panel"
      /(\d+(?:\.\d+)?)\s*$/i,           // "CHSM54RN 450" -> "450" (number at end)
      /\s(\d+(?:\.\d+)?)\s*$/i,         // "Model Name 450" -> "450" (number at end with space)
    ];
    
    for (const pattern of patterns) {
      const match = cellValue.match(pattern);
      if (match) {
        const wattage = `${match[1]}W`;
        this.logger.log(`✅ Extracted panel wattage: ${wattage}`);
        return wattage;
      }
    }
    
    this.logger.warn(`❌ Could not extract panel wattage from: "${cellValue}"`);
    return '';
  }

  private extractInverterSize(cellValue: string): string {
    // Extract size from inverter name (e.g., "PowerOcean HD-P1-3K-S1" -> "3K")
    const sizeMatch = cellValue.match(/HD-P1-([^-]+)-S1/);
    return sizeMatch ? sizeMatch[1] : '';
  }

  private extractBatterySize(cellValue: string): string {
    // Extract kWh from battery text (e.g., "10kWh Battery" -> "10kWh")
    const batteryMatch = cellValue.match(/(\d+(?:\.\d+)?)\s*kWh/i);
    return batteryMatch ? `${batteryMatch[1]}kWh` : '';
  }

  private countPanels(sheet: any, startCell: string, endCell: string): number {
    // Sum up all panel quantities from startCell to endCell (each array has a panel count)
    let totalPanelCount = 0;
    const startCol = startCell.charAt(0);
    const startRow = parseInt(startCell.substring(1));
    const endRow = parseInt(endCell.substring(1));
    
    // Sum up all numbers in the specified range (C69-C76 for off-peak, C70-C77 for EPVS)
    for (let row = startRow; row <= endRow; row++) {
      const cellAddress = `${startCol}${row}`;
      const cellValue = this.getCellValue(sheet, cellAddress);
      if (cellValue && cellValue.trim() !== '') {
        const numValue = parseInt(cellValue);
        if (!isNaN(numValue) && numValue > 0) {
          totalPanelCount += numValue;
        }
      }
    }
    
    return totalPanelCount;
  }

  private formatAddressWithPostcode(address: string, postcode: string): string {
    // Format address and postcode with comma separator
    const cleanAddress = (address || '').trim();
    const cleanPostcode = (postcode || '').trim();
    
    if (cleanAddress && cleanPostcode) {
      return `${cleanAddress}, ${cleanPostcode}`;
    } else if (cleanAddress) {
      return cleanAddress;
    } else if (cleanPostcode) {
      return cleanPostcode;
    } else {
      return '';
    }
  }

  private formatCurrency(value: any): string {
    console.log(`🔍 formatCurrency input: "${value}" (type: ${typeof value})`);
    
    if (!value || value === '') {
      console.log(`🔍 formatCurrency: Empty value, returning £0.00`);
      return '£0.00';
    }
    
    // Ensure we can handle values like "£1,025.40" or "1,025.40" without truncation
    const numericInput = typeof value === 'number' ? value.toString() : String(value);
    const cleaned = numericInput.replace(/[^0-9.\-]/g, ''); // remove £, commas, spaces, and any non-numeric chars
    const numValue = parseFloat(cleaned);
    console.log(`🔍 formatCurrency: Parsed number: ${numValue}`);
    
    if (isNaN(numValue)) {
      console.log(`🔍 formatCurrency: Invalid number, returning £0.00`);
      return '£0.00';
    }
    
    const formatted = `£${numValue.toFixed(2)}`;
    console.log(`🔍 formatCurrency: Final formatted value: "${formatted}"`);
    return formatted;
  }

  private formatNumber(value: any): string {
    if (!value || value === '') {
      return '0.00';
    }
    
    const numericInput = typeof value === 'number' ? value.toString() : String(value);
    const cleaned = numericInput.replace(/[^0-9.\-]/g, '');
    const numValue = parseFloat(cleaned);
    if (isNaN(numValue)) {
      return '0.00';
    }
    
    return numValue.toFixed(2);
  }

  private formatInteger(value: any): string {
    console.log(`🔍 formatInteger input: "${value}" (type: ${typeof value})`);
    
    if (!value || value === '') {
      console.log(`🔍 formatInteger: Empty value, returning 0`);
      return '0';
    }
    
    const numericInput = typeof value === 'number' ? value.toString() : String(value);
    const cleaned = numericInput.replace(/[^0-9.\-]/g, '');
    const numValue = parseFloat(cleaned);
    console.log(`🔍 formatInteger: Parsed number: ${numValue}`);
    
    if (isNaN(numValue)) {
      console.log(`🔍 formatInteger: Invalid number, returning 0`);
      return '0';
    }
    
    const rounded = Math.round(numValue);
    const formatted = rounded.toString();
    console.log(`🔍 formatInteger: Rounded to integer: ${rounded}, final formatted: "${formatted}"`);
    return formatted;
  }

  /**
   * Get display name for calculator type
   */
  private getCalculatorDisplayName(calculatorType: 'flux' | 'off-peak' | 'epvs'): string {
    if (calculatorType === 'flux' || calculatorType === 'epvs') {
      return 'Flux';
    }
    return 'Off Peak';
  }


  private async replaceVariablesInPresentation(templatePath: string, outputPath: string, variables: any, opportunityId: string, calculatorType: 'flux' | 'off-peak' | 'epvs'): Promise<void> {
    try {
      console.log('🔄 Replacing variables in PowerPoint...');
      
      // Read the template file
      const content = fs.readFileSync(templatePath, 'binary');
      const zip = new PizZip(content);
      
      // Create a mapping of variables to replace using the new variable names
      const addressPostcodeFormatted = this.formatAddressWithPostcode(variables.address, variables.postcode);
      const variableMapping = {
        'customer_name': variables.customerName || '',
        'date': variables.date || '',
        'postcode': addressPostcodeFormatted, // Replace old postcode with formatted address,postcode
        'address_postcode': addressPostcodeFormatted, // New variable for formatted address,postcode
        'p_w': variables.p_w || '',
        'p_q': variables.p_q || '',
        'i_s': variables.i_s || '',
        'b_s': variables.b_s || '',
        't_y_s_o': variables.t_y_s_o || '',
        't_y_s_g': variables.t_y_s_g || '',
        'year_1': variables.year_1 || '',
        'year_10': variables.year_10 || '',
        'year_25': variables.year_25 || '',
        // New proposal variables
        'term_var': variables.term_var || '',
        'yearly_plan_cost': variables.yearly_plan_cost || '',
        'yearly_saving': variables.yearly_saving || '',
        'yearly_con': variables.yearly_con || '',
        'yearly_consumption': variables.yearly_con || '',
        'consumption': variables.yearly_con || '',
        'yearly_con_cost': variables.yearly_con || '',
        'your_LT_profit': variables.your_LT_profit || '',
        // Template placeholders to replace
        'payment_type': variables.payment_type || 'Cash', // Replace Payment_type with B7 value
        'your_LT_Profit': variables.your_LT_profit || '' // Replace your_LT_Profit with B25 value
      };
      
      console.log('Variable mapping:', variableMapping);
      console.log('🔍 DEBUG yearly_con in variables:', variables.yearly_con);
      console.log('🔍 DEBUG yearly_con in mapping:', variableMapping['yearly_con']);
      
      // Replace variables in slide XML files only (to preserve layout)
      Object.keys(zip.files).forEach(filename => {
        if (filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml')) {
          try {
            let fileContent = zip.files[filename].asText();
            let contentChanged = false;
            
            // Use a more comprehensive replacement approach
            Object.keys(variableMapping).forEach(variable => {
              const value = String(variableMapping[variable] || '');
              if (value && value.trim() !== '') {
                // Try multiple replacement patterns to catch different XML structures
                const patterns = [
                  // Pattern 1: Within <a:t> tags
                  new RegExp(`(<a:t[^>]*>)(${variable})(</a:t>)`, 'g'),
                  // Pattern 2: Direct text replacement (fallback)
                  new RegExp(`\\b${variable}\\b`, 'g'),
                  // Pattern 3: Within any text content
                  new RegExp(`>${variable}<`, 'g')
                ];
                
                let replaced = false;
                patterns.forEach((pattern, index) => {
                  if (pattern.test(fileContent)) {
                    if (index === 0) {
                      // For <a:t> pattern, preserve the XML structure
                      fileContent = fileContent.replace(pattern, `$1${value}$3`);
                    } else if (index === 2) {
                      // For >variable< pattern, preserve the angle brackets
                      fileContent = fileContent.replace(pattern, `>${value}<`);
                    } else {
                      // For direct replacement
                      fileContent = fileContent.replace(pattern, value);
                    }
                    replaced = true;
                  }
                });
                
                if (replaced) {
                  contentChanged = true;
                  console.log(`  ✅ Replaced "${variable}" with "${value}" in ${filename}`);
                } else {
                  console.log(`  ⚠️ Variable "${variable}" not found in ${filename}`);
                }
              } else {
                console.log(`  ⚠️ Skipping empty value for "${variable}"`);
              }
            });
            
            // Update the file if content changed
            if (contentChanged) {
              zip.file(filename, fileContent);
            }
          } catch (e) {
            console.log(`  ⚠️ Could not process ${filename}: ${e.message}`);
          }
        }
      });
      
      // Generate the modified file
      const buffer = zip.generate({ type: 'nodebuffer' });
      fs.writeFileSync(outputPath, buffer);
      
        // Note: Slide 16 table insertion and £ symbol encoding fixes have been removedw
      
      // Create a metadata file with the variables that were replaced
      const metadataPath = outputPath.replace('.pptx', '_variables.json');
      fs.writeFileSync(metadataPath, JSON.stringify(variables, null, 2));
      
      console.log('✅ Variables replaced in PowerPoint successfully!');
      console.log('Variables replaced:', variables);
      
    } catch (error) {
      console.error('❌ Error replacing variables in presentation:', error);
      
      // Fallback: just copy the template if variable replacement fails
      console.log('🔄 Falling back to template copy...');
      fs.copyFileSync(templatePath, outputPath);
      
      // Create a metadata file with the variables that should be replaced
      const metadataPath = outputPath.replace('.pptx', '_variables.json');
      fs.writeFileSync(metadataPath, JSON.stringify(variables, null, 2));
      
      console.log('📋 Variables metadata created (manual replacement needed):', variables);
    }
  }

  // Note: fixPoundSymbolEncoding method removed - no longer fixing £ symbols in slide 16

  // Note: generatePoundSymbolFixScript method removed - no longer generating £ symbol fix scripts

  // Note: fixPoundSymbolEncodingWithIsolation method removed - no longer fixing £ symbols in slide 16

  // Note: generatePoundSymbolFixScriptWithIsolation method removed - no longer generating £ symbol fix scripts

  private async convertToPdf(pptxPath: string, pdfPath: string): Promise<boolean> {
    try {
      console.log('🔄 Converting PowerPoint to PDF using COM automation...');
      console.log('📁 PowerPoint file:', pptxPath);
      console.log('📄 PDF output:', pdfPath);
      
      // Create PowerShell script for COM-based conversion
      const scriptContent = this.generatePowerShellScript(pptxPath, pdfPath);
      const scriptPath = path.join(this.outputDir, `convert_${Date.now()}.ps1`);
      
      // Write the PowerShell script
      fs.writeFileSync(scriptPath, scriptContent);
      console.log('📝 PowerShell script created:', scriptPath);
      
      // Execute the PowerShell script
      const execAsync = promisify(exec);
      const command = `powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`;
      
      console.log('🚀 Executing PowerShell script...');
      const { stdout, stderr } = await execAsync(command);
      
      if (stdout) {
        console.log('PowerShell output:', stdout);
      }
      if (stderr) {
        console.log('PowerShell errors:', stderr);
      }
      
      // Check if PDF was created successfully
      if (fs.existsSync(pdfPath)) {
        console.log('✅ PDF conversion successful!');
        
        // Clean up the PowerShell script
        try {
          fs.unlinkSync(scriptPath);
          console.log('🧹 PowerShell script cleaned up');
        } catch (cleanupError) {
          console.log('⚠️ Could not clean up PowerShell script:', cleanupError);
        }
        
        return true;
      } else {
        throw new Error('PDF file was not created');
      }
    } catch (error) {
      console.error('❌ PDF conversion failed:', error);
      return false;
    }
  }

  private generatePowerShellScript(pptxPath: string, pdfPath: string): string {
    return `# Convert PowerPoint to PDF using COM automation
$ErrorActionPreference = "Stop"

# Configuration
$pptxFilePath = "${pptxPath.replace(/\\/g, '\\\\')}"
$pdfPath = "${pdfPath.replace(/\\/g, '\\\\')}"

Write-Host "Converting PowerPoint to PDF..." -ForegroundColor Green
Write-Host "PowerPoint file: $pptxFilePath" -ForegroundColor Yellow
Write-Host "PDF output: $pdfPath" -ForegroundColor Yellow

# Validate paths
if (!(Test-Path $pptxFilePath)) {
    throw "PowerPoint file not found: $pptxFilePath"
}

# Clean and validate PDF path
$pdfPath = [System.IO.Path]::GetFullPath($pdfPath)
$pdfDir = Split-Path $pdfPath -Parent

if (!(Test-Path $pdfDir)) {
    New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
    Write-Host "Created PDF directory: $pdfDir" -ForegroundColor Green
}

# Verify write permissions
try {
    [System.IO.File]::WriteAllText("$pdfDir\\test_write.tmp", "test")
    Remove-Item "$pdfDir\\test_write.tmp" -Force
    Write-Host "Write permissions verified for: $pdfDir" -ForegroundColor Green
} catch {
    throw "No write permissions for PDF directory: $pdfDir - $_"
}

# Kill any existing PowerPoint processes to prevent conflicts
try {
    Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Host "Cleared any existing PowerPoint processes" -ForegroundColor Green
} catch {
    Write-Host "No existing PowerPoint processes to clear" -ForegroundColor Yellow
}

$ppt = $null
$presentation = $null

try {
    # Create PowerPoint application
    Write-Host "Creating PowerPoint application..." -ForegroundColor Yellow
    $ppt = New-Object -ComObject PowerPoint.Application
    
    # Configure PowerPoint settings using correct enum values
    $ppt.Visible = -1  # -1 = msoTrue
    $ppt.DisplayAlerts = "ppAlertsNone"
    
    # Minimize the PowerPoint window to reduce visual impact
    try {
        $ppt.WindowState = "ppWindowMinimized"
    } catch {
        Write-Host "Warning: Could not minimize PowerPoint window" -ForegroundColor Yellow
    }
    
    Write-Host "PowerPoint application created successfully" -ForegroundColor Green
    
    # Open presentation
    Write-Host "Opening PowerPoint presentation..." -ForegroundColor Yellow
    try {
        $presentation = $ppt.Presentations.Open($pptxFilePath, $true, $true, $false)  # ReadOnly, Untitled, WithWindow
        Write-Host "Presentation opened successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open presentation: $_" -ForegroundColor Red
        throw "Could not open PowerPoint presentation: $_"
    }
    
    # Export to PDF using SaveAs method (simpler approach)
    Write-Host "Exporting to PDF..." -ForegroundColor Yellow
    try {
        # Use SaveAs method with PDF format - this is more reliable than ExportAsFixedFormat
        $presentation.SaveAs($pdfPath, 32)  # 32 = ppSaveAsPDF
        Write-Host "PDF export completed successfully using SaveAs method" -ForegroundColor Green
    } catch {
        Write-Host "SaveAs method failed, trying ExportAsFixedFormat: $_" -ForegroundColor Yellow
        try {
            # Fallback to ExportAsFixedFormat with numeric values instead of enums
            $presentation.ExportAsFixedFormat($pdfPath, 2)  # 2 = ppFixedFormatTypePDF
            Write-Host "PDF export completed successfully using ExportAsFixedFormat with numeric values" -ForegroundColor Green
        } catch {
            Write-Host "All PDF export methods failed: $_" -ForegroundColor Red
            throw "PDF export failed: $_"
        }
    }
    
    # Verify PDF was created
    if (Test-Path $pdfPath) {
        $pdfSize = (Get-Item $pdfPath).Length
        Write-Host "PDF created successfully: $pdfPath (Size: $pdfSize bytes)" -ForegroundColor Green
    } else {
        throw "PDF file was not created at expected location: $pdfPath"
    }
    
} catch {
    Write-Host "Critical error in PDF conversion: $_" -ForegroundColor Red
    throw $_
} finally {
    Write-Host "Starting cleanup process..." -ForegroundColor Yellow
    
    # Always try to close PowerPoint properly
    if ($presentation) {
        try {
            Write-Host "Closing presentation..." -ForegroundColor Yellow
            $presentation.Close()
        } catch {
            Write-Host "Warning: Error closing presentation: $_" -ForegroundColor Yellow
        }
    }
    
    if ($ppt) {
        try {
            Write-Host "Quitting PowerPoint..." -ForegroundColor Yellow
            $ppt.Quit()
        } catch {
            Write-Host "Warning: Error quitting PowerPoint: $_" -ForegroundColor Yellow
        }
    }
    
    # Force garbage collection and release COM objects
    try {
        Write-Host "Releasing COM objects..." -ForegroundColor Yellow
        if ($presentation) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
        }
        if ($ppt) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "COM objects released successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error releasing COM objects: $_" -ForegroundColor Yellow
    }
    
    # Final cleanup - kill any remaining PowerPoint processes
    try {
        Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-Host "Final PowerPoint process cleanup completed" -ForegroundColor Green
    } catch {
        Write-Host "No remaining PowerPoint processes to clean up" -ForegroundColor Yellow
    }
    
    Write-Host "Cleanup completed" -ForegroundColor Green
}

Write-Host "PowerPoint to PDF conversion completed successfully!" -ForegroundColor Green`;
  }

  async getPresentationPath(filename: string): Promise<string> {
    const filePath = path.join(this.outputDir, filename);
    if (!fs.existsSync(filePath)) {
      throw new Error('Presentation file not found');
    }
    return filePath;
  }

  async extractVariablesFromExcel(opportunityId: string, calculatorType: 'flux' | 'off-peak' | 'epvs' = 'off-peak') {
    try {
      return await this.extractCustomerDataFromExcel(opportunityId, calculatorType);
    } catch (error) {
      console.error('Variable extraction error:', error);
      throw error;
    }
  }

  /**
   * Extract Solar Projection data from Excel file for frontend display
   */
  async extractSolarProjectionData(opportunityId: string, calculatorType: 'flux' | 'off-peak' | 'epvs' = 'off-peak', fileName?: string) {
    try {
      console.log(`🔍 extractSolarProjectionData called with calculatorType: ${calculatorType}, fileName: ${fileName}`);
      
      // Only auto-detect calculator type if no specific type was requested (default parameter)
      // If user explicitly requests a calculator type, respect their choice
      const wasExplicitlyRequested = calculatorType !== 'off-peak' || arguments.length > 1;
      
      if (!wasExplicitlyRequested) {
        // Auto-detect calculator type only when no explicit choice was made
        const detectedCalculatorType = await this.detectCorrectCalculatorType(opportunityId);
        console.log(`🔍 Auto-detected calculator type: ${detectedCalculatorType}`);
        
        if (detectedCalculatorType !== calculatorType) {
          console.log(`🔧 Using auto-detected calculator type: ${detectedCalculatorType}`);
          calculatorType = detectedCalculatorType;
        }
      } else {
        console.log(`🔧 Using explicitly requested calculator type: ${calculatorType}`);
      }
      
      // For off-peak calculator, we need to get the payment type (but not set it yet)
      let paymentType: string | undefined = undefined;
      if (calculatorType === 'off-peak') {
        const extractedPaymentType = await this.getPaymentTypeFromWorkflow(opportunityId);
        paymentType = extractedPaymentType || undefined;
        console.log(`🔍 Extracted payment type: ${paymentType}`);
      }
      
      // Use findOpportunityExcelFile to find the specific file (respects fileName if provided)
      const excelFilePath = await this.findOpportunityExcelFile(opportunityId, calculatorType, fileName);
      
      if (!excelFilePath) {
        throw new Error(`${calculatorType === 'flux' || calculatorType === 'epvs' ? 'EPVS' : 'Off Peak'} opportunity Excel file not found for ID: ${opportunityId}${fileName ? ` with fileName: ${fileName}` : ''}`);
      }
      
      console.log(`Reading ${calculatorType === 'flux' || calculatorType === 'epvs' ? 'EPVS/Flux' : 'Off Peak'} Excel file: ${excelFilePath}`);

      // For off-peak calculator, handle payment type selection and data extraction in one script
      if (calculatorType === 'off-peak') {
        const solarProjectionData = await this.handleOffPeakPaymentTypeSelection(opportunityId, fileName);
        if (solarProjectionData) {
          return solarProjectionData;
        }
      }

      // For other calculator types, extract solar projection data using the existing method
      const solarProjectionData = await this.extractSolarProjectionsForPowerPoint(excelFilePath, calculatorType, paymentType);
      
      if (!solarProjectionData) {
        throw new Error('Failed to extract solar projection data from Excel file');
      }

      // Update the calculator type display name in the result
      if (solarProjectionData.summary) {
        solarProjectionData.summary.calculatorType = this.getCalculatorDisplayName(calculatorType);
      }
      if (solarProjectionData.title) {
        solarProjectionData.title = `Lifetime Savings Projections - ${this.getCalculatorDisplayName(calculatorType)}`;
      }

      return solarProjectionData;
    } catch (error) {
      console.error('Solar projection data extraction error:', error);
      throw error;
    }
  }

  /**
   * Handle off-peak payment type selection and extract solar projection data in one script
   */
  private async handleOffPeakPaymentTypeSelection(opportunityId: string, fileName?: string): Promise<any> {
    try {
      console.log(`🔧 Handling off-peak payment type selection for opportunity: ${opportunityId}, fileName: ${fileName}`);
      
      // Get the payment type from step 3 (calculator step) data
      let paymentType = await this.getPaymentTypeFromWorkflow(opportunityId);
      
      // If no payment type found in workflow, use a default
      if (!paymentType) {
        console.log('⚠️ No payment type found in workflow, using default "Hometree"');
        paymentType = 'Hometree';
      }
      
      console.log(`🔧 Payment type to use: ${paymentType}`);
      
      // Find the Excel file for this opportunity (respects fileName if provided)
      const excelFilePath = await this.findOpportunityExcelFile(opportunityId, 'off-peak', fileName);
      if (!excelFilePath) {
        throw new Error(`No Excel file found for opportunity ${opportunityId}${fileName ? ` with fileName: ${fileName}` : ''}`);
      }
      
      console.log(`📁 Found Excel file: ${excelFilePath}`);
      
      // Set the payment type in the Inputs tab and extract data in one operation
      // (This now handles both payment selection and data extraction in a single Excel session)
      const result = await this.setPaymentTypeInInputsTab(excelFilePath, paymentType);
      
      if (result && result.success) {
        console.log(`✅ Successfully set payment type ${paymentType} and extracted solar projection data`);
        console.log('PowerShell output length:', result.output.length);
        console.log('PowerShell output preview:', result.output.substring(0, 500));
        
        // Parse the JSON result from the PowerShell output
        let jsonLine: string | null = null;
        
        try {
          const output = result.output;
          const resultIndex = output.indexOf('RESULT:');
          if (resultIndex >= 0) {
            const firstBraceIndex = output.indexOf('{', resultIndex);
            if (firstBraceIndex >= 0) {
              let depth = 0;
              let endIndex = -1;
              for (let i = firstBraceIndex; i < output.length; i++) {
                const ch = output[i];
                if (ch === '{') {
                  depth++;
                } else if (ch === '}') {
                  depth--;
                  if (depth === 0) {
                    endIndex = i;
                    break;
                  }
                }
              }

              if (endIndex > firstBraceIndex) {
                const jsonString = output.substring(firstBraceIndex, endIndex + 1);
                console.log('Extracted JSON string:', jsonString.substring(0, 200) + (jsonString.length > 200 ? '...' : ''));
            const jsonData = JSON.parse(jsonString);
            console.log('✅ Successfully parsed JSON result');

            // Normalize numeric fields so the frontend receives numbers instead of formatted strings
            const coerceNumber = (value: any) => {
              if (value === null || value === undefined || value === '') {
                return null;
              }
              if (typeof value === 'number') {
                return value;
              }
              const cleaned = Number(String(value).replace(/[^0-9.-]/g, ''));
              return Number.isNaN(cleaned) ? null : cleaned;
            };

            if (jsonData?.summary) {
              jsonData.summary.yearlySavingYear1 = coerceNumber(jsonData.summary.yearlySavingYear1);
              jsonData.summary.yearlyPlanCost = coerceNumber(jsonData.summary.yearlyPlanCost);
              jsonData.summary.monthlyPlanCost = coerceNumber(jsonData.summary.monthlyPlanCost);
              jsonData.summary.yearlyContributionYear1 = coerceNumber(jsonData.summary.yearlyContributionYear1);
              jsonData.summary.lifetimeProfit = coerceNumber(jsonData.summary.lifetimeProfit);
              jsonData.summary.totalSavings = coerceNumber(jsonData.summary.totalSavings);
            }

            if (jsonData?.table?.rows) {
              jsonData.table.rows = jsonData.table.rows.map((row: any[]) =>
                row.map((cell: any) => (cell === null || cell === undefined ? '' : cell.toString()))
              );
              jsonData.table.totalRows = Array.isArray(jsonData.table.rows)
                ? jsonData.table.rows.length
                : 0;
            }

            console.log('🔍 Parsed table headers count:', jsonData?.table?.headers?.length ?? 0);
            console.log('🔍 Parsed table rows:', jsonData?.table?.rows?.length ?? 0);
            return jsonData;
              }
            }
          }

          console.error('❌ Could not locate complete JSON block in PowerShell output');
          console.error('PowerShell output preview:', output.substring(resultIndex >= 0 ? resultIndex : 0, Math.min(output.length, (resultIndex >= 0 ? resultIndex : 0) + 500)));
          return null;
        } catch (parseError) {
          console.error('❌ Failed to parse JSON result:', parseError);
          console.error('PowerShell output causing parse error:', result.output.substring(result.output.indexOf('RESULT:')));
          return null;
        }
      } else {
        console.error('❌ Failed to set payment type and extract data');
        return null;
      }
      
    } catch (error) {
      console.error('❌ Error handling off-peak payment type selection:', error);
      // Don't throw error - continue with solar projection extraction
      console.log('⚠️ Continuing with solar projection extraction despite payment type selection error');
    }
  }


  /**
   * Get payment type from workflow step 3 (calculator step) data
   */
  private async getPaymentTypeFromWorkflow(opportunityId: string): Promise<string | null> {
    try {
      // Import PrismaService to access the database
      const { PrismaService } = await import('../prisma/prisma.service');
      const prismaService = new PrismaService();
      
      // Get the opportunity progress with steps
      const progress = await prismaService.opportunityProgress.findUnique({
        where: { ghlOpportunityId: opportunityId },
        include: { steps: true }
      });
      
      if (!progress) {
        console.log(`⚠️ No progress found for opportunity: ${opportunityId}`);
        return null;
      }
      
      // Find step 3 (calculator step) or step 4 (payment step)
      const calculatorStep = progress.steps.find(step => step.stepNumber === 3);
      const paymentStep = progress.steps.find(step => step.stepNumber === 4);
      
      if (!calculatorStep || !calculatorStep.data) {
        console.log(`⚠️ No calculator step data found for opportunity: ${opportunityId}`);
        return null;
      }
      
      // Extract payment type from step data - try both calculator and payment steps
      const calculatorStepData = calculatorStep.data as any;
      const paymentStepData = paymentStep?.data as any;
      
      // Extract payment type from step data
      const paymentType = calculatorStepData.paymentType || 
                         calculatorStepData.selectedPaymentType || 
                         calculatorStepData.paymentMethod ||
                         calculatorStepData.savedInputs?.payment_method ||
                         calculatorStepData.savedInputs?.paymentType ||
                         calculatorStepData.savedInputs?.selectedPaymentType ||
                         paymentStepData?.paymentType ||
                         paymentStepData?.selectedPaymentType ||
                         paymentStepData?.paymentMethod ||
                         paymentStepData?.savedInputs?.payment_method ||
                         paymentStepData?.savedInputs?.paymentType ||
                         paymentStepData?.savedInputs?.selectedPaymentType;
      
      console.log(`🔍 Step 3 data:`, calculatorStepData);
      if (paymentStepData) {
        console.log(`🔍 Step 4 data:`, paymentStepData);
      }
      console.log(`🔍 Extracted payment type: ${paymentType}`);
      
      return paymentType;
    } catch (error) {
      console.error('❌ Error getting payment type from workflow:', error);
      return null;
    }
  }

  /**
   * Detect the correct calculator type based on existing files
   */
  private async detectCorrectCalculatorType(opportunityId: string): Promise<'flux' | 'off-peak' | 'epvs'> {
    try {
      console.log(`🔍 Detecting correct calculator type for opportunity: ${opportunityId}`);
      
      // Check for Off Peak file first
      const offPeakFiles = fs.readdirSync(this.opportunitiesDir).filter(file => 
        file.includes(opportunityId) && file.endsWith('.xlsm')
      );
      
      if (offPeakFiles.length > 0) {
        console.log(`✅ Found Off Peak file: ${offPeakFiles[0]}`);
        return 'off-peak';
      }
      
      // Check for EPVS/Flux file
      const epvsFiles = fs.readdirSync(this.epvsOpportunitiesDir).filter(file => 
        file.includes(opportunityId) && file.endsWith('.xlsm')
      );
      
      if (epvsFiles.length > 0) {
        console.log(`✅ Found EPVS/Flux file: ${epvsFiles[0]}`);
        return 'flux';
      }
      
      console.log(`⚠️ No calculator files found for opportunity: ${opportunityId}, defaulting to off-peak`);
      return 'off-peak';
    } catch (error) {
      console.error('❌ Error detecting calculator type:', error);
      return 'off-peak';
    }
  }


  /**
   * Set payment type in the Inputs tab to refresh calculations and extract solar projection data
   */
  private async setPaymentTypeInInputsTab(excelFilePath: string, paymentType: string): Promise<any> {
    try {
      console.log(`🔧 Setting payment type ${paymentType} in Inputs tab...`);
      
      // Create PowerShell script to set payment type in Inputs tab
      const scriptContent = `
# Set Payment Type in Inputs Tab
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
$password = "99"
$paymentType = "${paymentType}"

Write-Host "Setting payment type '$paymentType' in Inputs tab: $filePath" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    # Disable error dialogs to prevent VBA popups
    $excel.DisplayAlerts = $false
    $excel.AlertBeforeOverwriting = $false
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Check if file exists
    if (-not (Test-Path $filePath)) {
        throw "File does not exist: $filePath"
    }
    
    # Open workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        # Try to open with password first
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
        # Enable calculations and ensure workbook can be modified
        $excel.Calculation = -4105  # xlCalculationAutomatic
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $true
        Write-Host "Excel calculation mode set to automatic" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($filePath)
            Write-Host "Workbook opened without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook without password: $_" -ForegroundColor Red
            throw "Could not open workbook: $_"
        }
    }
    
    # Find the Inputs sheet
    $inputsSheet = $null
    try {
        $sheetCount = $workbook.Worksheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Inputs") {
                $inputsSheet = $workbook.Worksheets.Item($i)
                Write-Host "Found Inputs sheet: $sheetName" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "Error finding Inputs sheet: $_" -ForegroundColor Red
    }
    
    if (-not $inputsSheet) {
        Write-Host "Available sheet names:" -ForegroundColor Red
        try {
            $sheetCount = $workbook.Worksheets.Count
            for ($i = 1; $i -le $sheetCount; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                Write-Host "  '$sheetName'" -ForegroundColor Red
            }
        } catch {
            Write-Host "Error listing sheets in error message: $_" -ForegroundColor Red
        }
        throw "Inputs sheet not found. Available sheets listed above."
    }
    
    # Use the working SetOptionHomeTree macro approach (same as our successful test)
    Write-Host "Using SetOptionHomeTree macro approach for payment selection..." -ForegroundColor Green
    
    # Ensure we're on the Inputs sheet
    $inputsSheet.Select()
    Start-Sleep -Seconds 1
    
    # Run the SetOptionHomeTree macro (this is the working method from our test)
    try {
        Write-Host "Running SetOptionHomeTree macro..." -ForegroundColor Yellow
        $excel.Run("SetOptionHomeTree")
        Write-Host "✅ SetOptionHomeTree macro executed successfully" -ForegroundColor Green
        Start-Sleep -Seconds 3
    } catch {
        Write-Host "⚠️ SetOptionHomeTree macro failed: $_" -ForegroundColor Yellow
        
        # Fallback: Set values directly in Lookups sheet
        Write-Host "Fallback: Setting values directly in Lookups sheet..." -ForegroundColor Yellow
        try {
            $lookupsSheet = $workbook.Worksheets.Item("Lookups")
            if ($lookupsSheet) {
                # Set PaymentMethodOptionValue to 2 (Hometree)
                $lookupsSheet.Range("PaymentMethodOptionValue").Value2 = 2
                Write-Host "Set Lookups PaymentMethodOptionValue to 2" -ForegroundColor Green
                
                # Set NewFinance to False (for Hometree)
                $lookupsSheet.Range("NewFinance").Value2 = "False"
                Write-Host "Set Lookups NewFinance to False" -ForegroundColor Green
                
                $excel.Calculate()
                Start-Sleep -Seconds 3
            }
        } catch {
            Write-Host "Fallback method also failed: $_" -ForegroundColor Red
        }
    }
    
    # Now extract the solar projection data from the same Excel session
    Write-Host "Extracting solar projection data from the same Excel session..." -ForegroundColor Green
    
    # Find the Solar Projections sheet
    $solarProjectionsSheet = $null
    try {
        $sheetCount = $workbook.Worksheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Solar Projections") {
                $solarProjectionsSheet = $workbook.Worksheets.Item($i)
                Write-Host "Found Solar Projections sheet: $sheetName" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "Error finding Solar Projections sheet: $_" -ForegroundColor Red
    }
    
    if ($solarProjectionsSheet) {
        Write-Host "Extracting summary data from sheet: Solar Projections" -ForegroundColor Green
        
        # Extract summary data
        $summary = @{}
        
        # Extract payment type from B7
        try {
            $paymentTypeValue = $solarProjectionsSheet.Range("B7").Value2
            $summary.paymentType = if ($paymentTypeValue) { $paymentTypeValue.ToString() } else { "Hometree" }
            Write-Host "Extracted paymentType from Excel B7: $($summary.paymentType)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract paymentType from B7: $_" -ForegroundColor Yellow
            $summary.paymentType = "Hometree"
        }
        
        # Extract yearly saving year 1 from B20
        try {
            $cellValue = $solarProjectionsSheet.Range("B20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlySavingYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlySavingYear1 = $null
            }
            Write-Host "Extracted yearlySavingYear1 from B20: $($summary.yearlySavingYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlySavingYear1 from B20: $_" -ForegroundColor Yellow
            $summary.yearlySavingYear1 = $null
        }
        
        # Extract yearly plan cost from summary cell B14 (Yearly Plan Cost)
        try {
            $cellValue = $solarProjectionsSheet.Range("B14").Value2
            if ($cellValue -ne $null) {
                Write-Host "Raw yearlyPlanCost cell (B14) value: $cellValue (type: $($cellValue.GetType().FullName))" -ForegroundColor Yellow
                $cellText = $cellValue.ToString()
                Write-Host "Cell text: $cellText" -ForegroundColor Yellow
                $cleaned = [System.Text.RegularExpressions.Regex]::Replace($cellText, "[^0-9.\-]", "")
                Write-Host "Cleaned numeric text: $cleaned" -ForegroundColor Yellow
                $numericValue = [double]::Parse($cleaned, [System.Globalization.CultureInfo]::InvariantCulture)
                Write-Host "Parsed numeric value (invariant): $numericValue" -ForegroundColor Yellow
                # Truncate to 2 decimal places (no rounding)
                $truncatedValue = [math]::Floor($numericValue * 100) / 100
                $summary.yearlyPlanCost = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyPlanCost = $null
            }
            Write-Host "Extracted yearlyPlanCost from B14: $($summary.yearlyPlanCost)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlyPlanCost from B14: $_" -ForegroundColor Yellow
            $summary.yearlyPlanCost = $null
        }
        
        # Extract lifetime profit from B25
        try {
            $cellValue = $solarProjectionsSheet.Range("B25").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.lifetimeProfit = $truncatedValue.ToString("F2")
            } else {
                $summary.lifetimeProfit = $null
            }
            Write-Host "Extracted lifetimeProfit from B25: $($summary.lifetimeProfit)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract lifetimeProfit from B25: $_" -ForegroundColor Yellow
            $summary.lifetimeProfit = $null
        }
        
        # Extract term from B10
        try {
            $termValue = $solarProjectionsSheet.Range("B10").Value2
            $summary.term = if ($termValue) { $termValue.ToString() } else { $null }
            Write-Host "Extracted term from B10: $($summary.term)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract term from B10: $_" -ForegroundColor Yellow
            $summary.term = $null
        }
        
        # Extract yearly contribution year 1 from D20
        try {
            $cellValue = $solarProjectionsSheet.Range("D20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlyContributionYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyContributionYear1 = $null
            }
            Write-Host "Extracted yearlyContributionYear1 from D20: $($summary.yearlyContributionYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlyContributionYear1 from D20: $_" -ForegroundColor Yellow
            $summary.yearlyContributionYear1 = $null
        }
        
        # Set calculator type
        $summary.calculatorType = "Off Peak"
        
        # Extract table headers (row 2, columns 6-14 - all data columns)
        Write-Host "Extracting table headers from row 2, columns 6-14..." -ForegroundColor Yellow
        $headers = @()
        for ($col = 6; $col -le 14; $col++) {
            try {
                $headerValue = $solarProjectionsSheet.Cells.Item(2, $col).Value2
                if ($headerValue -ne $null) {
                    $headers += $headerValue.ToString()
                    Write-Host "  Header $col : $($headerValue.ToString())" -ForegroundColor Green
                } else {
                    $headers += ""
                    Write-Host "  Header $col : (empty)" -ForegroundColor Gray
                }
            } catch {
                Write-Host "  Warning: Could not extract header from column $col : $_" -ForegroundColor Yellow
                $headers += ""
            }
        }

        # Extract table rows (rows 3-32, columns 6-14 - all data columns for 30 years)
        Write-Host "Extracting table rows from rows 3-32, columns 6-14..." -ForegroundColor Yellow
        $rows = @()
        for ($row = 3; $row -le 32; $row++) {
            $rowData = @()
            $hasData = $false
            for ($col = 6; $col -le 14; $col++) {
                try {
                    $cellValue = $solarProjectionsSheet.Cells.Item($row, $col).Value2
                    if ($cellValue -ne $null) {
                        if ($cellValue -is [double]) {
                            # Truncate to 2 decimal places (no rounding)
                        $truncatedValue = [math]::Floor([double]$cellValue * 100) / 100
                        $rowData += $truncatedValue.ToString("F2")
                        } else {
                            $rowData += $cellValue.ToString()
                        }
                        $hasData = $true
                    } else {
                        $rowData += ""
                    }
                } catch {
                    Write-Host "  Warning: Could not extract cell ($row,$col): $_" -ForegroundColor Yellow
                    $rowData += ""
                }
            }
            # Only add row if it has data
            if ($hasData) {
                # Create a proper array structure for each row
                $rows += ,@($rowData)
                Write-Host "  Row $row : Added with data" -ForegroundColor Green
            }
        }
        Write-Host "Extracted $($rows.Count) data rows" -ForegroundColor Green
        
        # Create the result object
        $result = @{
            summary = $summary
            title = "Lifetime Savings Projections - Off Peak"
            metadata = @{
                extractedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.ffffffK")
                sheetName = "Solar Projections"
                sourceFile = (Split-Path $filePath -Leaf)
            }
            table = @{
                headers = $headers
                rows = $rows
                totalRows = $rows.Count
            }
        }
        
        # Output the result as JSON
        $jsonResult = $result | ConvertTo-Json -Depth 10
        Write-Host "RESULT: $jsonResult" -ForegroundColor Green
        
        Write-Host "✅ Solar projection data extracted successfully from same Excel session" -ForegroundColor Green
    } else {
        Write-Host "❌ Solar Projections sheet not found" -ForegroundColor Red
    }
    
    # Close Excel properly to release file lock
    Write-Host "Closing Excel application to release file lock..." -ForegroundColor Yellow
    try {
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Payment type setting and data extraction completed - file closed" -ForegroundColor Green
    exit 0
    
} catch {
    Write-Host "Critical error in payment type setting: $_" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
`;

      // Write the script to a temporary file
      const scriptPath = path.join(process.cwd(), `temp-payment-type-inputs-${Date.now()}.ps1`);
      fs.writeFileSync(scriptPath, scriptContent);
      console.log(`Created payment type setting script: ${scriptPath}`);
      
      // Execute the PowerShell script
      const result = await this.runPowerShellScript(scriptPath);
      
      // Clean up temporary script
      try {
        fs.unlinkSync(scriptPath);
        console.log('Cleaned up temporary script');
      } catch (cleanupError) {
        console.warn('Warning: Could not clean up temporary script:', cleanupError);
      }
      
      if (result.success) {
        console.log('✅ Successfully set payment type in Inputs tab');
        console.log('Payment type setting PowerShell output:', result.output);
        return result;
      } else {
        console.error('❌ Failed to set payment type in Inputs tab:', result.error);
        console.error('PowerShell output:', result.output);
        return null;
      }
      
    } catch (error) {
      console.error('❌ Error setting payment type in Inputs tab:', error);
      return null;
    }
  }

  /**
   * Find opportunity Excel file
   */
  private async findOpportunityExcelFile(opportunityId: string, calculatorType: 'off-peak' | 'flux' | 'epvs' = 'off-peak', fileName?: string): Promise<string | null> {
    try {
      // Determine which directory to search based on calculator type
      let searchDir: string;
      if (calculatorType === 'flux' || calculatorType === 'epvs') {
        searchDir = this.epvsOpportunitiesDir;
      } else {
        searchDir = this.opportunitiesDir;
      }
      
      if (!fs.existsSync(searchDir)) {
        console.log(`${calculatorType} opportunities directory does not exist: ${searchDir}`);
        return null;
      }
      
      const files = fs.readdirSync(searchDir);
      console.log(`Found ${files.length} files in ${calculatorType} opportunities directory`);
      
      // If fileName is provided, try to find the exact file first
      if (fileName) {
        console.log(`🔍 Looking for specific file: ${fileName}`);
        const exactFile = files.find(file => file === fileName);
        if (exactFile) {
          const fullPath = path.join(searchDir, exactFile);
          console.log(`✅ Found exact file match: ${fullPath}`);
          return fullPath;
        } else {
          console.log(`⚠️ Exact file ${fileName} not found, falling back to opportunity ID search`);
        }
      }
      
      // Look for files that contain the opportunity ID
      const matchingFiles = files.filter(file => file.includes(opportunityId));
      
      if (matchingFiles.length === 0) {
        console.log(`No matching file found for opportunity ID: ${opportunityId} in ${calculatorType} directory`);
        return null;
      }
      
      // If only one file found, return it
      if (matchingFiles.length === 1) {
        const fullPath = path.join(searchDir, matchingFiles[0]);
        console.log(`Found single matching file: ${fullPath}`);
        return fullPath;
      }
      
      // Multiple files found - look for versioned files and find the latest one
      const baseFileName = matchingFiles[0].split('-v')[0]; // Get base name before version
      
      // Look for versioned files (v1, v2, v3, etc.) and find the latest one
      const basePattern = baseFileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const versionRegex = new RegExp(`^${basePattern}-v(\\d+)\\.xlsm$`);
      
      let latestFile: string | null = null;
      let maxVersion = 0;
      
      for (const file of matchingFiles) {
        const versionMatch = file.match(versionRegex);
        if (versionMatch) {
          const version = parseInt(versionMatch[1], 10);
          if (version > maxVersion) {
            maxVersion = version;
            latestFile = file;
          }
        }
      }
      
      // If no versioned files found, look for non-versioned files
      if (!latestFile) {
        const nonVersionedFile = `${baseFileName}.xlsm`;
        if (matchingFiles.includes(nonVersionedFile)) {
          latestFile = nonVersionedFile;
        }
      }
      
      if (latestFile) {
        const fullPath = path.join(searchDir, latestFile);
        console.log(`Found latest file: ${fullPath} (version: ${maxVersion || 'non-versioned'})`);
        return fullPath;
      }
      
      // Fallback: return the first matching file
      const fullPath = path.join(searchDir, matchingFiles[0]);
      console.log(`Fallback: using first matching file: ${fullPath}`);
      return fullPath;
    } catch (error) {
      console.error('Error finding opportunity Excel file:', error);
      return null;
    }
  }

  /**
   * Select radio button and keep Excel file open (custom implementation)
   */
  private async selectRadioButtonAndKeepOpen(excelFilePath: string, shapeName: string): Promise<void> {
    try {
      console.log(`🔧 Using custom COM automation to select radio button ${shapeName} and keep file open`);
      
      const scriptContent = `
# Custom Radio Button Automation - Keep File Open
$ErrorActionPreference = "Stop"

# Configuration
$shapeName = "${shapeName}"
$excelFilePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
$password = "99"

Write-Host "Starting custom radio button automation for shape: $shapeName" -ForegroundColor Green
Write-Host "Excel file: $excelFilePath" -ForegroundColor Yellow

try {
    # Remove read-only attribute if it exists
    if (Test-Path $excelFilePath) {
        $fileAttributes = Get-ItemProperty -Path $excelFilePath -Name Attributes
        if ($fileAttributes.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
            Write-Host "File is read-only, removing read-only attribute..." -ForegroundColor Yellow
            Set-ItemProperty -Path $excelFilePath -Name Attributes -Value ($fileAttributes.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly))
            Write-Host "Read-only attribute removed successfully" -ForegroundColor Green
        } else {
            Write-Host "File is not read-only" -ForegroundColor Green
        }
    }
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1  # Enable all macros
    
    # Disable error dialogs to prevent VBA popups
    $excel.DisplayAlerts = $false
    $excel.AlertBeforeOverwriting = $false
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Open the workbook
    Write-Host "Opening workbook: $excelFilePath" -ForegroundColor Green
    
    try {
        $workbook = $excel.Workbooks.Open($excelFilePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
        # Enable calculations and ensure workbook can be modified
        $excel.Calculation = -4105  # xlCalculationAutomatic
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $true
        Write-Host "Excel calculation mode set to automatic" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without password..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($excelFilePath)
            Write-Host "Workbook opened successfully without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook: $_" -ForegroundColor Red
            throw "Could not open workbook"
        }
    }
    
    # Get the Inputs worksheet
    $worksheet = $workbook.Worksheets.Item("Inputs")
    Write-Host "Found Inputs worksheet" -ForegroundColor Green
    
    # Unprotect all worksheets
    Write-Host "Unprotecting all worksheets..." -ForegroundColor Green
    foreach ($ws in $workbook.Worksheets) {
        if ($ws.ProtectContents) {
            try {
                $ws.Unprotect($password)
                Write-Host "Unprotected worksheet: $($ws.Name)" -ForegroundColor Green
            } catch {
                Write-Host "Could not unprotect worksheet: $($ws.Name) - $_" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Worksheet $($ws.Name) is not protected" -ForegroundColor Green
        }
    }
    
    # Use the working SetOptionHomeTree macro approach instead of direct shape manipulation
    Write-Host "Using SetOptionHomeTree macro approach for payment selection..." -ForegroundColor Green
    
    # Ensure we're on the Inputs sheet
    $worksheet.Select()
    Start-Sleep -Seconds 1
    
    # Run the SetOptionHomeTree macro (this is the working method from our test)
    try {
        Write-Host "Running SetOptionHomeTree macro..." -ForegroundColor Yellow
        $excel.Run("SetOptionHomeTree")
        Write-Host "✅ SetOptionHomeTree macro executed successfully" -ForegroundColor Green
        Start-Sleep -Seconds 3
    } catch {
        Write-Host "⚠️ SetOptionHomeTree macro failed: $_" -ForegroundColor Yellow
        
        # Fallback: Set values directly in Lookups sheet
        Write-Host "Fallback: Setting values directly in Lookups sheet..." -ForegroundColor Yellow
        try {
            $lookupsSheet = $workbook.Worksheets.Item("Lookups")
            if ($lookupsSheet) {
                # Set PaymentMethodOptionValue to 2 (Hometree)
                $lookupsSheet.Range("PaymentMethodOptionValue").Value2 = 2
                Write-Host "Set Lookups PaymentMethodOptionValue to 2" -ForegroundColor Green
                
                # Set NewFinance to False (for Hometree)
                $lookupsSheet.Range("NewFinance").Value2 = "False"
                Write-Host "Set Lookups NewFinance to False" -ForegroundColor Green
                
                $excel.Calculate()
                Start-Sleep -Seconds 3
            }
                        } catch {
            Write-Host "Fallback method also failed: $_" -ForegroundColor Red
        }
    }
    
    # Save the workbook
    Write-Host "Saving workbook..." -ForegroundColor Green
    try {
        $workbook.Save()
        Write-Host "Workbook saved successfully" -ForegroundColor Green
    } catch {
        Write-Host "Save failed, but continuing without save..." -ForegroundColor Yellow
        Write-Host "Warning: Changes may not be persisted" -ForegroundColor Yellow
    }
    
    # Close Excel properly to release file lock
    Write-Host "Closing Excel application to release file lock..." -ForegroundColor Yellow
    try {
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Custom radio button automation completed successfully (file closed)!" -ForegroundColor Green
    
} catch {
    Write-Error "Critical error in custom radio button automation: $($_.Exception.Message)"
    exit 1
}
      `;
      
      const scriptPath = path.join(process.cwd(), `temp-custom-radio-button-keep-open-${Date.now()}.ps1`);
      fs.writeFileSync(scriptPath, scriptContent);
      
      console.log(`Created custom PowerShell script: ${scriptPath}`);
      console.log('Executing custom PowerShell script for radio button automation...');
      
      const result = await this.runPowerShellScript(scriptPath);
      
      // Clean up temporary script
      try {
        fs.unlinkSync(scriptPath);
        console.log(`Cleaned up temporary script: ${scriptPath}`);
      } catch (cleanupError) {
        console.log(`Failed to cleanup temporary script: ${cleanupError.message}`);
      }
      
      if (result.success) {
        console.log(`✅ Successfully selected radio button ${shapeName} using custom COM automation (file kept open)`);
      } else {
        throw new Error(`Custom PowerShell script failed: ${result.error}`);
      }
      
    } catch (error) {
      console.error(`❌ Error in custom COM automation: ${error.message}`);
      throw new Error(`Failed to select radio button using custom COM: ${error.message}`);
    }
  }




  async downloadPresentation(filename: string): Promise<{ filePath: string; mimeType: string; size: number }> {
    const filePath = path.join(this.outputDir, filename);
    
    if (!fs.existsSync(filePath)) {
      // If PDF file doesn't exist, try to find the corresponding PowerPoint file
      if (filename.endsWith('.pdf')) {
        const pptxFilename = filename.replace('.pdf', '.pptx');
        const pptxPath = path.join(this.outputDir, pptxFilename);
        
        if (fs.existsSync(pptxPath)) {
          this.logger.warn(`PDF file ${filename} not found, but PowerPoint file ${pptxFilename} exists. Returning PowerPoint file instead.`);
          const stats = fs.statSync(pptxPath);
          const ext = path.extname(pptxFilename).toLowerCase();
          const mimeType = ext === '.pdf' ? 'application/pdf' : 
                          ext === '.pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation' : 
                          'application/octet-stream';
          
          return {
            filePath: pptxPath,
            mimeType,
            size: stats.size
          };
        }
      }
      
      throw new Error(`Presentation file not found: ${filename}`);
    }

    const stats = fs.statSync(filePath);
    const ext = path.extname(filename).toLowerCase();
    const mimeType = ext === '.pdf' ? 'application/pdf' : 
                    ext === '.pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation' : 
                    'application/octet-stream';

    return {
      filePath,
      mimeType,
      size: stats.size
    };
  }

  /**
   * Run PowerShell script and return result
   */
  private async runPowerShellScript(scriptPath: string): Promise<{ success: boolean; output: string; error: string }> {
    return new Promise((resolve) => {
      const { spawn } = require('child_process');
      const powershell = spawn('powershell.exe', ['-ExecutionPolicy', 'Bypass', '-File', scriptPath]);
      
      let output = '';
      let error = '';
      
      powershell.stdout.on('data', (data: Buffer) => {
        output += data.toString();
      });
      
      powershell.stderr.on('data', (data: Buffer) => {
        error += data.toString();
      });
      
      powershell.on('close', (code: number) => {
        if (code === 0) {
          resolve({ success: true, output, error });
        } else {
          resolve({ success: false, output, error });
        }
      });
      
      powershell.on('error', (err: Error) => {
        resolve({ success: false, output, error: err.message });
      });
    });
  }

  /**
   * Execute screenshot and insertion script
   */
  private async executeScreenshotAndInsertionScript(scriptPath: string): Promise<void> {
    return new Promise((resolve, reject) => {
      try {
        const { spawn } = require('child_process');
        const powershell = spawn('powershell.exe', ['-ExecutionPolicy', 'Bypass', '-File', scriptPath]);
        
        let output = '';
        let stderr = '';
        
        powershell.stdout.on('data', (data: Buffer) => {
          output += data.toString();
        });
        
        powershell.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
        });
        
        powershell.on('close', (code: number) => {
          if (code === 0) {
            if (output) {
              console.log('PowerShell output:', output);
            }
            if (stderr) {
              console.log('PowerShell stderr:', stderr);
            }
            
            console.log('✅ Screenshot and insertion completed successfully!');
            resolve();
          } else {
            console.error('❌ PowerShell script failed with code:', code);
            console.error('PowerShell stderr:', stderr);
            reject(new Error(`PowerShell script failed with code ${code}`));
          }
        });
        
        powershell.on('error', (error: Error) => {
          console.error('❌ Error executing PowerShell script:', error);
          reject(error);
        });
        
      } catch (error) {
        console.error('❌ Error executing screenshot and insertion script:', error);
        reject(error);
      }
    });
  }

  // Note: takeAndInsertScreenshot method removed - no longer inserting tables into slide 16

  /**
   * Extract solar projections data for PowerPoint presentation
   */
  private async extractSolarProjectionsForPowerPoint(excelFilePath: string, calculatorType: 'flux' | 'off-peak' | 'epvs', paymentType?: string): Promise<any> {
    console.log('Extracting solar projections data for PowerPoint...');
    
    try {
      // Create PowerShell script to extract solar projection data
      const scriptContent = this.createSolarProjectionExtractionScript(excelFilePath, calculatorType, paymentType);
      const scriptPath = path.join(process.cwd(), `temp-solar-projection-extraction-${Date.now()}.ps1`);
      
      fs.writeFileSync(scriptPath, scriptContent);
      console.log(`Created solar projection extraction script: ${scriptPath}`);
      
      const result = await this.runPowerShellScript(scriptPath);
      
      // Clean up temporary script
      try {
        fs.unlinkSync(scriptPath);
        console.log('Cleaned up temporary script');
      } catch (cleanupError) {
        console.warn('Warning: Could not clean up temporary script:', cleanupError);
      }
      
      if (result.success) {
        console.log('✅ Successfully extracted solar projection data');
        console.log('PowerShell output:', result.output);
        // Parse the JSON result from the PowerShell output
        try {
          const jsonMatch = result.output.match(/RESULT: (.+)/);
          if (jsonMatch) {
            const jsonData = JSON.parse(jsonMatch[1]);
            return jsonData;
          } else {
            console.error('❌ No JSON result found in PowerShell output');
            console.error('Full PowerShell output:', result.output);
            return null;
          }
        } catch (parseError) {
          console.error('❌ Failed to parse JSON result:', parseError);
          console.error('Raw output that failed to parse:', result.output);
          return null;
        }
      } else {
        console.error('❌ Failed to extract solar projection data');
        console.error('PowerShell error output:', result.error);
        console.error('PowerShell standard output:', result.output);
        return null;
      }
      
    } catch (error) {
      console.error('❌ Error in solar projection extraction:', error);
      return null;
    }
  }

  /**
   * Create PowerShell script to extract solar projection data from Excel
   */
  private createSolarProjectionExtractionScript(excelFilePath: string, calculatorType: 'flux' | 'off-peak' | 'epvs', paymentType?: string): string {
    const excelFilePathEscaped = excelFilePath.replace(/\\/g, '\\\\');
    
    return `
# Extract Solar Projection Data from Excel
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePathEscaped}"
$password = "99"

Write-Host "Extracting solar projection data from: $filePath" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    # Disable error dialogs to prevent VBA popups
    $excel.DisplayAlerts = $false
    $excel.AlertBeforeOverwriting = $false
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Check if file exists
    if (-not (Test-Path $filePath)) {
        throw "File does not exist: $filePath"
    }
    
    # Open workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        # Try to open with password first
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
        # Enable calculations and ensure workbook can be modified
        $excel.Calculation = -4105  # xlCalculationAutomatic
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $true
        Write-Host "Excel calculation mode set to automatic" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        try {
        $workbook = $excel.Workbooks.Open($filePath)
        Write-Host "Workbook opened without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook without password: $_" -ForegroundColor Red
            throw "Could not open workbook: $_"
        }
    }
    
    # Set payment type if provided (for off-peak calculator)
    ${paymentType ? `
    if ("${paymentType}" -ne "") {
        Write-Host "Setting payment type '${paymentType}' before data extraction..." -ForegroundColor Yellow
        
             # Convert payment type text to numeric value
             $paymentTypeNumeric = switch ("${paymentType}") {
                 "Cash" { "1" }
                 "Finance" { "2" }
                 "Hometree" { "2" }  # Hometree should map to Finance (2), not 3
                 default { "2" }  # Default to Finance
             }
        
        Write-Host "Converting payment type '${paymentType}' to numeric value '$paymentTypeNumeric'" -ForegroundColor Cyan
        
        # Method 1: Run SetOptionHomeTree macro from Inputs sheet
        try {
            Write-Host "Method 1: Running SetOptionHomeTree macro from Inputs sheet..." -ForegroundColor Yellow
            
            # Find and select the Inputs sheet
            $inputsSheet = $null
            for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                if ($sheetName -eq "Inputs") {
                    $inputsSheet = $workbook.Worksheets.Item($i)
                    break
                }
            }
            
            if ($inputsSheet) {
                Write-Host "Found Inputs sheet, selecting it..." -ForegroundColor Cyan
                $inputsSheet.Select()
                Start-Sleep -Seconds 1
                
                # Run the SetOptionHomeTree macro
                try {
                    $excel.Run("SetOptionHomeTree")
                    Write-Host "✅ SetOptionHomeTree macro executed successfully" -ForegroundColor Green
                    Start-Sleep -Seconds 3
                } catch {
                    Write-Host "⚠️ SetOptionHomeTree macro failed: $_" -ForegroundColor Yellow
                }
            } else {
                Write-Host "Inputs sheet not found" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Method 1 (SetOptionHomeTree macro) failed: $_" -ForegroundColor Yellow
        }
        
        # Method 2: Set values directly in Lookups sheet (safer than named ranges)
        try {
            Write-Host "Method 2: Setting values in Lookups sheet..." -ForegroundColor Yellow
            
            # Find Lookups sheet
            $lookupsSheet = $null
            for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                if ($sheetName -eq "Lookups") {
                    $lookupsSheet = $workbook.Worksheets.Item($i)
                    break
                }
            }
            
            if ($lookupsSheet) {
                # Set PaymentMethodOptionValue to 2 (Hometree)
                $lookupsSheet.Range("PaymentMethodOptionValue").Value2 = 2
                Write-Host "Set Lookups PaymentMethodOptionValue to 2" -ForegroundColor Green
                
                # Set NewFinance to False (for Hometree)
                $lookupsSheet.Range("NewFinance").Value2 = "False"
                Write-Host "Set Lookups NewFinance to False" -ForegroundColor Green
                
                $excel.Calculate()
                Start-Sleep -Seconds 3
            } else {
                Write-Host "Lookups sheet not found" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Method 2 (Lookups sheet) failed: $_" -ForegroundColor Yellow
        }
        
        
        # Method 4: Try to set payment type directly in Inputs sheet
        try {
            $inputsSheet = $null
            for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                if ($sheetName -eq "Inputs") {
                    $inputsSheet = $workbook.Worksheets.Item($i)
                    break
                }
            }
            
            if ($inputsSheet) {
                Write-Host "Found Inputs sheet, attempting to set payment type..." -ForegroundColor Cyan
                
                # Try to find and set payment method in Inputs sheet
                # Look for cells that might contain payment method options
                for ($row = 1; $row -le 50; $row++) {
                    for ($col = 1; $col -le 10; $col++) {
                        $cellValue = $inputsSheet.Cells.Item($row, $col).Value2
                        if ($cellValue -and $cellValue.ToString().ToLower().Contains("hometree")) {
                            Write-Host ("Found Hometree reference at row " + $row + ", col " + $col + ": " + $cellValue) -ForegroundColor Yellow
                            # Try to set the cell to trigger the selection
                            $inputsSheet.Cells.Item($row, $col).Value2 = $cellValue
                        }
                    }
                }
                
                $excel.Calculate()
                Start-Sleep -Seconds 3
            }
        } catch {
            Write-Host "Method 4 (Inputs sheet) failed: $_" -ForegroundColor Yellow
        }
        
        Write-Host "Payment type setting completed" -ForegroundColor Green
        
             # Simple recalculation to avoid VBA popups
             Write-Host "Performing simple recalculation..." -ForegroundColor Cyan
             
             # Just do a simple calculation without aggressive methods that might trigger VBA popups
             $excel.Calculate()
             Start-Sleep -Seconds 3
             
             Write-Host "Simple recalculation completed" -ForegroundColor Green
    }
    ` : ''}
    
    # List all available sheets for debugging
    Write-Host "Available sheets in workbook:" -ForegroundColor Cyan
    try {
        $sheetCount = $workbook.Worksheets.Count
        Write-Host "Total worksheets: $sheetCount" -ForegroundColor Cyan
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            Write-Host "  - $sheetName" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "Error listing sheets: $_" -ForegroundColor Red
    }
    
    # Find the Solar Projections sheet
    $solarProjectionsSheet = $null
    $possibleNames = @("Solar Projections", "*Solar Projection*", "*Solar*Projection*", "*Projection*", "*Solar*")
    
    try {
        $sheetCount = $workbook.Worksheets.Count
        # First try exact matches
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Solar Projections") {
                $solarProjectionsSheet = $workbook.Worksheets.Item($i)
                Write-Host "Found Solar Projections sheet: $sheetName (exact match)" -ForegroundColor Green
            break
        }
    }
    
        # If no exact match, try pattern matching
    if (-not $solarProjectionsSheet) {
            for ($i = 1; $i -le $sheetCount; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                foreach ($pattern in $possibleNames) {
                    if ($sheetName -like $pattern) {
                        $solarProjectionsSheet = $workbook.Worksheets.Item($i)
                        Write-Host "Found Solar Projections sheet: $sheetName (matched pattern: $pattern)" -ForegroundColor Green
                        break
                    }
                }
                if ($solarProjectionsSheet) { break }
            }
        }
    } catch {
        Write-Host "Error finding sheets: $_" -ForegroundColor Red
    }
    
    if (-not $solarProjectionsSheet) {
        Write-Host "Available sheet names:" -ForegroundColor Red
        try {
            $sheetCount = $workbook.Worksheets.Count
            for ($i = 1; $i -le $sheetCount; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                Write-Host "  '$sheetName'" -ForegroundColor Red
            }
        } catch {
            Write-Host "Error listing sheets in error message: $_" -ForegroundColor Red
        }
        throw "Solar Projections sheet not found. Available sheets listed above."
    }
    
    # Extract summary data with error handling
    Write-Host "Extracting summary data from sheet: $($solarProjectionsSheet.Name)" -ForegroundColor Yellow
    
    $summary = @{}
    
    # Extract payment type from Excel cell B7 first
    try {
        $paymentTypeFromExcel = $solarProjectionsSheet.Range("B7").Value2
        if ($paymentTypeFromExcel -ne $null -and $paymentTypeFromExcel -ne "") {
            $summary.paymentType = $paymentTypeFromExcel.ToString()
            Write-Host "Extracted paymentType from Excel B7: $($summary.paymentType)" -ForegroundColor Green
        } else {
            # Fallback to workflow payment type if Excel cell is empty
            $summary.paymentType = "${paymentType}"
            Write-Host "Excel B7 is empty, using workflow paymentType: $($summary.paymentType)" -ForegroundColor Yellow
        }
    } catch {
        # Fallback to workflow payment type if Excel extraction fails
        $summary.paymentType = "${paymentType}"
        Write-Host "Failed to extract paymentType from Excel B7, using workflow: $($summary.paymentType)" -ForegroundColor Yellow
    }
    
    # Extract fields based on payment type
    $paymentTypeLower = $summary.paymentType.ToLower()
    Write-Host "Processing fields for payment type: $($summary.paymentType) (lowercase: $paymentTypeLower)" -ForegroundColor Cyan
    
    # Common fields for all payment types
    try {
        $cellValue = $solarProjectionsSheet.Range("B20").Value2
        if ($cellValue -ne $null) {
            $numericValue = [double]$cellValue
            # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlySavingYear1 = $truncatedValue.ToString("F2")
        } else {
            $summary.yearlySavingYear1 = $null
        }
        Write-Host "Extracted yearlySavingYear1 from B20: $($summary.yearlySavingYear1)" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not extract yearlySavingYear1 from B20: $_" -ForegroundColor Yellow
        $summary.yearlySavingYear1 = $null
    }
    
    try {
        $cellValue = $solarProjectionsSheet.Range("B25").Value2
        if ($cellValue -ne $null) {
            $numericValue = [double]$cellValue
            # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.lifetimeProfit = $truncatedValue.ToString("F2")
        } else {
            $summary.lifetimeProfit = $null
        }
        Write-Host "Extracted lifetimeProfit from B25: $($summary.lifetimeProfit)" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not extract lifetimeProfit from B25: $_" -ForegroundColor Yellow
        $summary.lifetimeProfit = $null
    }
    
    # Payment type specific fields
    if ($paymentTypeLower -eq "hometree") {
        Write-Host "Processing Hometree payment type fields..." -ForegroundColor Green
        
        # Term from B10
        try {
            $cellValue = $solarProjectionsSheet.Range("B10").Value2
            $summary.term = if ($cellValue -ne $null) { $cellValue.ToString() } else { $null }
            Write-Host "Extracted term from B10: $($summary.term)" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not extract term from B10: $_" -ForegroundColor Yellow
            $summary.term = $null
        }
        
        # Yearly Plan Cost from B14 (explicit cell per design)
        try {
            $cellValue = $solarProjectionsSheet.Range("B14").Value2
            if ($cellValue -ne $null) {
                $cellText = $cellValue.ToString()
                $cleaned = [System.Text.RegularExpressions.Regex]::Replace($cellText, "[^0-9.\-]", "")
                $numericValue = [double]::Parse($cleaned, [System.Globalization.CultureInfo]::InvariantCulture)
                # Truncate to 2 decimal places (no rounding)
                $truncatedValue = [math]::Floor($numericValue * 100) / 100
                $summary.yearlyPlanCost = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyPlanCost = $null
            }
            Write-Host "Extracted yearlyPlanCost from B14: $($summary.yearlyPlanCost)" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not extract yearlyPlanCost from B14: $_" -ForegroundColor Yellow
            $summary.yearlyPlanCost = $null
        }
        
        # Yearly Contribution from D20
        try {
            $cellValue = $solarProjectionsSheet.Range("D20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlyContributionYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyContributionYear1 = $null
            }
            Write-Host "Extracted yearlyContributionYear1 from D20: $($summary.yearlyContributionYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not extract yearlyContributionYear1 from D20: $_" -ForegroundColor Yellow
            $summary.yearlyContributionYear1 = $null
        }
        
        # Set monthly plan cost to null for Hometree
        $summary.monthlyPlanCost = $null
        
    } elseif ($paymentTypeLower -eq "finance") {
        Write-Host "Processing Finance payment type fields..." -ForegroundColor Green
        
        # Term from B10
        try {
            $cellValue = $solarProjectionsSheet.Range("B10").Value2
            $summary.term = if ($cellValue -ne $null) { $cellValue.ToString() } else { $null }
            Write-Host "Extracted term from B10: $($summary.term)" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not extract term from B10: $_" -ForegroundColor Yellow
            $summary.term = $null
        }
        
        # Monthly Plan Cost from B14 (same cell as yearly for Hometree, but displayed as monthly)
        try {
            $cellValue = $solarProjectionsSheet.Range("B14").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.monthlyPlanCost = $truncatedValue.ToString("F2")
            } else {
                $summary.monthlyPlanCost = $null
            }
            Write-Host "Extracted monthlyPlanCost from B14: $($summary.monthlyPlanCost)" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not extract monthlyPlanCost from B14: $_" -ForegroundColor Yellow
            $summary.monthlyPlanCost = $null
        }
        
        # Yearly Contribution from D20
        try {
            $cellValue = $solarProjectionsSheet.Range("D20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlyContributionYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyContributionYear1 = $null
            }
            Write-Host "Extracted yearlyContributionYear1 from D20: $($summary.yearlyContributionYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not extract yearlyContributionYear1 from D20: $_" -ForegroundColor Yellow
            $summary.yearlyContributionYear1 = $null
        }
        
        # Set yearly plan cost to null for Finance
        $summary.yearlyPlanCost = $null
        
    } elseif ($paymentTypeLower -eq "cash") {
        Write-Host "Processing Cash payment type fields..." -ForegroundColor Green
        
        # For Cash, only yearly saving and lifetime profit are shown
        # Set other fields to null
        $summary.term = $null
        $summary.yearlyPlanCost = $null
        $summary.monthlyPlanCost = $null
        $summary.yearlyContributionYear1 = $null
        
        Write-Host "Cash payment type - only showing yearly saving and lifetime profit" -ForegroundColor Green
        
    } else {
        Write-Host "Unknown payment type '$($summary.paymentType)' - setting all optional fields to null" -ForegroundColor Yellow
        $summary.term = $null
        $summary.yearlyPlanCost = $null
        $summary.monthlyPlanCost = $null
        $summary.yearlyContributionYear1 = $null
    }
    
    # Set the display name for calculator type (Flux for both flux and epvs)
    $displayName = if ("${calculatorType}" -eq "flux" -or "${calculatorType}" -eq "epvs") { "Flux" } else { "Off Peak" }
    $summary.calculatorType = $displayName
    
    # Extract table headers (row 2, columns 6-14 - all data columns)
    Write-Host "Extracting table headers from row 2, columns 6-14..." -ForegroundColor Yellow
    $headers = @()
    for ($col = 6; $col -le 14; $col++) {
        try {
            $headerValue = $solarProjectionsSheet.Cells.Item(2, $col).Value2
            if ($headerValue -ne $null) {
                $headers += $headerValue.ToString()
                Write-Host "  Header $col : $($headerValue.ToString())" -ForegroundColor Green
            } else {
                $headers += ""
                Write-Host "  Header $col : (empty)" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Warning: Could not extract header from column $col : $_" -ForegroundColor Yellow
            $headers += ""
        }
    }

    # Extract table rows (rows 3-32, columns 6-14 - all data columns for 30 years)
    Write-Host "Extracting table rows from rows 3-32, columns 6-14..." -ForegroundColor Yellow
    $rows = @()
    for ($row = 3; $row -le 32; $row++) {
        $rowData = @()
        $hasData = $false
        for ($col = 6; $col -le 14; $col++) {
            try {
                $cellValue = $solarProjectionsSheet.Cells.Item($row, $col).Value2
                if ($cellValue -ne $null) {
                    if ($cellValue -is [double]) {
                        # Truncate to 2 decimal places (no rounding)
                        $truncatedValue = [math]::Floor([double]$cellValue * 100) / 100
                        $rowData += $truncatedValue.ToString("F2")
                    } else {
                        $rowData += $cellValue.ToString()
                    }
                    $hasData = $true
                } else {
                    $rowData += ""
                }
            } catch {
                Write-Host "  Warning: Could not extract cell ($row,$col): $_" -ForegroundColor Yellow
                $rowData += ""
            }
        }
        # Only add row if it has data
        if ($hasData) {
            # Create a proper array structure for each row
            $rows += ,@($rowData)
            Write-Host "  Row $row : Added with data" -ForegroundColor Green
        }
    }
    Write-Host "Extracted $($rows.Count) data rows" -ForegroundColor Green

    # Create the result object
    $result = @{
        summary = $summary
        title = "Lifetime Savings Projections - $($summary.calculatorType)"
        metadata = @{
            extractedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffffffK")
            sheetName = $solarProjectionsSheet.Name
            sourceFile = (Split-Path $filePath -Leaf)
        }
        table = @{
            headers = $headers
            rows = $rows
            totalRows = $rows.Count
        }
    }
    
    # Convert to JSON
    $jsonResult = $result | ConvertTo-Json -Depth 4 -Compress
    Write-Host "RESULT: $jsonResult" -ForegroundColor Green
    
    # Close Excel
    try {
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Solar projection data extraction completed successfully!" -ForegroundColor Green
    exit 0
    
} catch {
    Write-Host "Critical error in solar projection data extraction: $_" -ForegroundColor Red
    exit 1
}
`;
  }

  /**
   * Get presentation information
   */
  async getPresentationInfo(filename: string): Promise<any> {
    try {
      const filePath = path.join(this.outputDir, filename);
      
      if (!fs.existsSync(filePath)) {
        throw new Error(`Presentation file not found: ${filename}`);
      }

      const stats = fs.statSync(filePath);
      const ext = path.extname(filename).toLowerCase();
      const mimeType = ext === '.pdf' ? 'application/pdf' : 
                      ext === '.pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation' : 
                      'application/octet-stream';

      return {
        filename,
        filePath,
        mimeType,
        size: stats.size,
        created: stats.birthtime,
        modified: stats.mtime
      };
    } catch (error) {
      console.error('Error getting presentation info:', error);
      throw error;
    }
  }

  /**
   * List all presentations
   */
  async listPresentations(): Promise<any[]> {
    try {
      if (!fs.existsSync(this.outputDir)) {
        return [];
      }

      const files = fs.readdirSync(this.outputDir);
      const presentations: any[] = [];

      for (const file of files) {
        if (file.endsWith('.pdf') || file.endsWith('.pptx')) {
          const filePath = path.join(this.outputDir, file);
          const stats = fs.statSync(filePath);
          const ext = path.extname(file).toLowerCase();
          const mimeType = ext === '.pdf' ? 'application/pdf' : 
                          ext === '.pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation' : 
                          'application/octet-stream';

          presentations.push({
            filename: file,
            filePath,
            mimeType,
            size: stats.size,
            created: stats.birthtime,
            modified: stats.mtime
          });
        }
      }

      return presentations.sort((a, b) => b.modified.getTime() - a.modified.getTime());
    } catch (error) {
      console.error('Error listing presentations:', error);
      throw error;
    }
  }

  /**
   * Update payment method and extract solar projection data in one operation
   */
  async updatePaymentMethodAndExtractData(
    opportunityId: string,
    paymentMethod: string,
    calculatorType: 'flux' | 'off-peak' | 'epvs' = 'off-peak',
    fileName?: string
  ): Promise<any> {
    try {
      console.log(`🔧 updatePaymentMethodAndExtractData called with paymentMethod: ${paymentMethod}, calculatorType: ${calculatorType}, fileName: ${fileName}`);
      
      // Find the Excel file for this opportunity
      const excelFilePath = await this.findOpportunityExcelFile(opportunityId, calculatorType, fileName);
      if (!excelFilePath) {
        throw new Error(`No Excel file found for opportunity ${opportunityId}`);
      }
      
      console.log(`📁 Found Excel file: ${excelFilePath}`);
      
      // Set the payment method and extract data in one operation
      const result = await this.setPaymentMethodInExcel(excelFilePath, paymentMethod, calculatorType);
      
      if (result && result.success) {
        console.log(`✅ Successfully set payment method ${paymentMethod} and extracted solar projection data`);
        
        // Parse the JSON result from the PowerShell output
        return this.parsePowerShellJsonResult(result.output);
      } else {
        throw new Error('Failed to update payment method and extract data');
      }
      
    } catch (error) {
      console.error('❌ Error updating payment method:', error);
      throw error;
    }
  }

  /**
   * Update terms and extract solar projection data in one operation
   */
  async updateTermsAndExtractData(
    opportunityId: string,
    terms: number,
    calculatorType: 'flux' | 'off-peak' | 'epvs' = 'off-peak',
    fileName?: string
  ): Promise<any> {
    try {
      console.log(`🔧 updateTermsAndExtractData called with terms: ${terms}, calculatorType: ${calculatorType}, fileName: ${fileName}`);
      
      // Validate terms (1-30 years)
      if (terms < 1 || terms > 30) {
        throw new Error('Terms must be between 1 and 30 years');
      }
      
      // Find the Excel file for this opportunity
      const excelFilePath = await this.findOpportunityExcelFile(opportunityId, calculatorType, fileName);
      if (!excelFilePath) {
        throw new Error(`No Excel file found for opportunity ${opportunityId}`);
      }
      
      console.log(`📁 Found Excel file: ${excelFilePath}`);
      
      // Set the terms and extract data in one operation
      const result = await this.setTermsInExcel(excelFilePath, terms, calculatorType);
      
      if (result && result.success) {
        console.log(`✅ Successfully set terms ${terms} and extracted solar projection data`);
        
        // Parse the JSON result from the PowerShell output
        return this.parsePowerShellJsonResult(result.output);
      } else {
        throw new Error('Failed to update terms and extract data');
      }
      
    } catch (error) {
      console.error('❌ Error updating terms:', error);
      throw error;
    }
  }

  /**
   * Set payment method in Excel and extract solar projection data
   */
  private async setPaymentMethodInExcel(
    excelFilePath: string,
    paymentMethod: string,
    calculatorType: 'flux' | 'off-peak' | 'epvs'
  ): Promise<any> {
    try {
      console.log(`🔧 Setting payment method ${paymentMethod} in Excel...`);
      
      // Create PowerShell script to set payment method and extract data
      const scriptContent = `
# Set Payment Method and Extract Solar Projection Data
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
$password = "99"
$paymentMethod = "${paymentMethod}"
$calculatorType = "${calculatorType}"

Write-Host "Setting payment method '$paymentMethod' in Excel: $filePath" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    # Disable error dialogs to prevent VBA popups
    $excel.DisplayAlerts = $false
    $excel.AlertBeforeOverwriting = $false
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Check if file exists
    if (-not (Test-Path $filePath)) {
        throw "File does not exist: $filePath"
    }
    
    # Open workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        # Try to open with password first
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
        # Enable calculations and ensure workbook can be modified
        $excel.Calculation = -4105  # xlCalculationAutomatic
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $true
        Write-Host "Excel calculation mode set to automatic" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($filePath)
            Write-Host "Workbook opened without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook without password: $_" -ForegroundColor Red
            throw "Could not open workbook: $_"
        }
    }
    
    # Find the Inputs sheet
    $inputsSheet = $null
    try {
        $sheetCount = $workbook.Worksheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Inputs") {
                $inputsSheet = $workbook.Worksheets.Item($i)
                Write-Host "Found Inputs sheet: $sheetName" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "Error finding Inputs sheet: $_" -ForegroundColor Red
    }
    
    if (-not $inputsSheet) {
        Write-Host "Available sheet names:" -ForegroundColor Red
        try {
            $sheetCount = $workbook.Worksheets.Count
            for ($i = 1; $i -le $sheetCount; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                Write-Host "  '$sheetName'" -ForegroundColor Red
            }
        } catch {
            Write-Host "Error listing sheets in error message: $_" -ForegroundColor Red
        }
        throw "Inputs sheet not found. Available sheets listed above."
    }
    
    # Set payment method based on the selected option
    Write-Host "Setting payment method: $paymentMethod" -ForegroundColor Green
    
    # Map payment method names to Excel values
    $paymentMethodValue = switch ($paymentMethod.ToLower()) {
        "hometree" { 2 }
        "cash" { 1 }
        "finance" { 3 }
        default { 2 }  # Default to Hometree
    }
    
    Write-Host "Payment method value: $paymentMethodValue" -ForegroundColor Yellow
    
    # Try to use the SetOptionHomeTree macro approach first
    try {
        Write-Host "Using macro approach for payment selection..." -ForegroundColor Green
        $inputsSheet.Select()
        Start-Sleep -Seconds 1
        
        # Run the appropriate macro based on payment method
        switch ($paymentMethod.ToLower()) {
            "hometree" { 
                $excel.Run("SetOptionHomeTree")
                Write-Host "✅ SetOptionHomeTree macro executed" -ForegroundColor Green
            }
            "cash" { 
                $excel.Run("SetOptionCash")
                Write-Host "✅ SetOptionCash macro executed" -ForegroundColor Green
            }
            "finance" { 
                $excel.Run("SetOptionNewFinance")
                Write-Host "✅ SetOptionNewFinance macro executed" -ForegroundColor Green
            }
            default { 
                $excel.Run("SetOptionHomeTree")
                Write-Host "✅ SetOptionHomeTree macro executed (default)" -ForegroundColor Green
            }
        }
        
        Start-Sleep -Seconds 3
    } catch {
        Write-Host "⚠️ Macro approach failed: $_" -ForegroundColor Yellow
        
        # Fallback: Set values directly in Lookups sheet
        Write-Host "Fallback: Setting values directly in Lookups sheet..." -ForegroundColor Yellow
        try {
            $lookupsSheet = $workbook.Worksheets.Item("Lookups")
            if ($lookupsSheet) {
                # Set PaymentMethodOptionValue
                $lookupsSheet.Range("PaymentMethodOptionValue").Value2 = $paymentMethodValue
                Write-Host "Set Lookups PaymentMethodOptionValue to $paymentMethodValue" -ForegroundColor Green
                
                # Set NewFinance based on payment method
                $newFinanceValue = if ($paymentMethod.ToLower() -eq "finance") { "True" } else { "False" }
                $lookupsSheet.Range("NewFinance").Value2 = $newFinanceValue
                Write-Host "Set Lookups NewFinance to $newFinanceValue" -ForegroundColor Green
                
                $excel.Calculate()
                Start-Sleep -Seconds 3
            }
        } catch {
            Write-Host "Fallback method also failed: $_" -ForegroundColor Red
        }
    }
    
    # Now extract the solar projection data from the same Excel session
    Write-Host "Extracting solar projection data from the same Excel session..." -ForegroundColor Green
    
    # Find the Solar Projections sheet
    $solarProjectionsSheet = $null
    try {
        $sheetCount = $workbook.Worksheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Solar Projections") {
                $solarProjectionsSheet = $workbook.Worksheets.Item($i)
                Write-Host "Found Solar Projections sheet: $sheetName" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "Error finding Solar Projections sheet: $_" -ForegroundColor Red
    }
    
    if ($solarProjectionsSheet) {
        Write-Host "Extracting summary data from sheet: Solar Projections" -ForegroundColor Green
        
        # Extract summary data
        $summary = @{}
        
        # Extract payment type from B7
        try {
            $paymentTypeValue = $solarProjectionsSheet.Range("B7").Value2
            $summary.paymentType = if ($paymentTypeValue) { $paymentTypeValue.ToString() } else { $paymentMethod }
            Write-Host "Extracted paymentType from Excel B7: $($summary.paymentType)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract paymentType from B7: $_" -ForegroundColor Yellow
            $summary.paymentType = $paymentMethod
        }
        
        # Extract yearly saving year 1 from B20
        try {
            $cellValue = $solarProjectionsSheet.Range("B20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlySavingYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlySavingYear1 = $null
            }
            Write-Host "Extracted yearlySavingYear1 from B20: $($summary.yearlySavingYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlySavingYear1 from B20: $_" -ForegroundColor Yellow
            $summary.yearlySavingYear1 = $null
        }
        
        # Extract yearly plan cost from B14 (explicit cell per design)
        try {
            $cellValue = $solarProjectionsSheet.Range("B14").Value2
            if ($cellValue -ne $null) {
                $cellText = $cellValue.ToString()
                $cleaned = [System.Text.RegularExpressions.Regex]::Replace($cellText, "[^0-9.\-]", "")
                $numericValue = [double]::Parse($cleaned, [System.Globalization.CultureInfo]::InvariantCulture)
                # Truncate to 2 decimal places (no rounding)
                $truncatedValue = [math]::Floor($numericValue * 100) / 100
                $summary.yearlyPlanCost = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyPlanCost = $null
            }
            Write-Host "Extracted yearlyPlanCost from B14: $($summary.yearlyPlanCost)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlyPlanCost from B14: $_" -ForegroundColor Yellow
            $summary.yearlyPlanCost = $null
        }
        
        # Extract lifetime profit from B25
        try {
            $cellValue = $solarProjectionsSheet.Range("B25").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.lifetimeProfit = $truncatedValue.ToString("F2")
            } else {
                $summary.lifetimeProfit = $null
            }
            Write-Host "Extracted lifetimeProfit from B25: $($summary.lifetimeProfit)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract lifetimeProfit from B25: $_" -ForegroundColor Yellow
            $summary.lifetimeProfit = $null
        }
        
        # Extract term from B10
        try {
            $termValue = $solarProjectionsSheet.Range("B10").Value2
            $summary.term = if ($termValue) { $termValue.ToString() } else { $null }
            Write-Host "Extracted term from B10: $($summary.term)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract term from B10: $_" -ForegroundColor Yellow
            $summary.term = $null
        }
        
        # Extract yearly contribution year 1 from D20
        try {
            $cellValue = $solarProjectionsSheet.Range("D20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlyContributionYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyContributionYear1 = $null
            }
            Write-Host "Extracted yearlyContributionYear1 from D20: $($summary.yearlyContributionYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlyContributionYear1 from D20: $_" -ForegroundColor Yellow
            $summary.yearlyContributionYear1 = $null
        }
        
        # Set calculator type
        $summary.calculatorType = if ($calculatorType -eq "off-peak") { "Off Peak" } else { $calculatorType.ToUpper() }
        
        # Extract table headers (row 2, columns 6-14 - all data columns)
        Write-Host "Extracting table headers from row 2, columns 6-14..." -ForegroundColor Yellow
        $headers = @()
        for ($col = 6; $col -le 14; $col++) {
            try {
                $headerValue = $solarProjectionsSheet.Cells.Item(2, $col).Value2
                if ($headerValue -ne $null) {
                    $headers += $headerValue.ToString()
                    Write-Host "  Header $col : $($headerValue.ToString())" -ForegroundColor Green
                } else {
                    $headers += ""
                    Write-Host "  Header $col : (empty)" -ForegroundColor Gray
                }
            } catch {
                Write-Host "  Warning: Could not extract header from column $col : $_" -ForegroundColor Yellow
                $headers += ""
            }
        }

        # Extract table rows (rows 3-32, columns 6-14 - all data columns for 30 years)
        Write-Host "Extracting table rows from rows 3-32, columns 6-14..." -ForegroundColor Yellow
        $rows = @()
        for ($row = 3; $row -le 32; $row++) {
            $rowData = @()
            $hasData = $false
            for ($col = 6; $col -le 14; $col++) {
                try {
                    $cellValue = $solarProjectionsSheet.Cells.Item($row, $col).Value2
                    if ($cellValue -ne $null) {
                        if ($cellValue -is [double]) {
                            # Truncate to 2 decimal places (no rounding)
                        $truncatedValue = [math]::Floor([double]$cellValue * 100) / 100
                        $rowData += $truncatedValue.ToString("F2")
                        } else {
                            $rowData += $cellValue.ToString()
                        }
                        $hasData = $true
                    } else {
                        $rowData += ""
                    }
                } catch {
                    Write-Host "  Warning: Could not extract cell ($row,$col): $_" -ForegroundColor Yellow
                    $rowData += ""
                }
            }
            
            if ($hasData) {
                $rows += ,$rowData
                Write-Host "  Row $row : $($rowData -join ', ')" -ForegroundColor Green
            } else {
                Write-Host "  Row $row : (no data)" -ForegroundColor Gray
            }
        }

        # Create the result object
        $result = @{
            title = "Lifetime Savings Projections - $($summary.calculatorType)"
            summary = $summary
            table = @{
                headers = $headers
                rows = $rows
                totalRows = $rows.Count
            }
        }

        # Convert to JSON and output
        $jsonResult = $result | ConvertTo-Json -Depth 10
        Write-Host "RESULT: $jsonResult" -ForegroundColor Green
        
    } else {
        Write-Host "Solar Projections sheet not found" -ForegroundColor Red
        throw "Solar Projections sheet not found"
    }
    
    # Save the workbook
    Write-Host "Saving workbook..." -ForegroundColor Green
    try {
        $workbook.Save()
        Write-Host "Workbook saved successfully" -ForegroundColor Green
    } catch {
        Write-Host "Save failed, but continuing without save..." -ForegroundColor Yellow
        Write-Host "Warning: Changes may not be persisted" -ForegroundColor Yellow
    }
    
    # Close Excel properly to release file lock
    Write-Host "Closing Excel application to release file lock..." -ForegroundColor Yellow
    try {
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Payment method update and data extraction completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "Critical error in payment method update: $($_.Exception.Message)"
    exit 1
}
      `;
      
      const scriptPath = path.join(process.cwd(), `temp_payment_method_update_${Date.now()}.ps1`);
      await fsPromises.writeFile(scriptPath, scriptContent);
      
      try {
        const { stdout, stderr } = await this.execAsync(`powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`);
        
        if (stderr && stderr.trim()) {
          console.error('PowerShell stderr:', stderr);
        }
        
        return {
          success: true,
          output: stdout
        };
      } finally {
        // Clean up the temporary script file
        try {
          await fsPromises.unlink(scriptPath);
        } catch (cleanupError) {
          console.warn('Failed to clean up temporary script file:', cleanupError);
        }
      }
    } catch (error) {
      console.error('❌ Error setting payment method in Excel:', error);
      throw error;
    }
  }

  /**
   * Set terms in Excel and extract solar projection data
   */
  private async setTermsInExcel(
    excelFilePath: string,
    terms: number,
    calculatorType: 'flux' | 'off-peak' | 'epvs'
  ): Promise<any> {
    try {
      console.log(`🔧 Setting terms ${terms} in Excel...`);
      
      // Create PowerShell script to set terms and extract data
      const scriptContent = `
# Set Terms and Extract Solar Projection Data
$ErrorActionPreference = "Stop"

# Configuration
$filePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
$password = "99"
$terms = ${terms}
$calculatorType = "${calculatorType}"

Write-Host "Setting terms '$terms' in Excel: $filePath" -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    
    # Enable macros
    $excel.AutomationSecurity = 1  # Enable all macros
    
    # Disable error dialogs to prevent VBA popups
    $excel.DisplayAlerts = $false
    $excel.AlertBeforeOverwriting = $false
    
    Write-Host "Excel application created successfully" -ForegroundColor Green
    
    # Check if file exists
    if (-not (Test-Path $filePath)) {
        throw "File does not exist: $filePath"
    }
    
    # Open workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Yellow
    
    try {
        # Try to open with password first
        $workbook = $excel.Workbooks.Open($filePath, 0, $false, 5, $password)
        Write-Host "Workbook opened successfully with password" -ForegroundColor Green
        # Set calculation to MANUAL for performance - we'll calculate only what we need
        $excel.Calculation = -4135  # xlCalculationManual
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $true
        Write-Host "Excel calculation mode set to manual for performance" -ForegroundColor Green
    } catch {
        Write-Host "Failed to open with password, trying without..." -ForegroundColor Yellow
        try {
            $workbook = $excel.Workbooks.Open($filePath)
            Write-Host "Workbook opened without password" -ForegroundColor Green
        } catch {
            Write-Host "Failed to open workbook without password: $_" -ForegroundColor Red
            throw "Could not open workbook: $_"
        }
    }
    
    # Find the Inputs sheet
    $inputsSheet = $null
    try {
        $sheetCount = $workbook.Worksheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Inputs") {
                $inputsSheet = $workbook.Worksheets.Item($i)
                Write-Host "Found Inputs sheet: $sheetName" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "Error finding Inputs sheet: $_" -ForegroundColor Red
    }
    
    if (-not $inputsSheet) {
        Write-Host "Available sheet names:" -ForegroundColor Red
        try {
            $sheetCount = $workbook.Worksheets.Count
            for ($i = 1; $i -le $sheetCount; $i++) {
                $sheetName = $workbook.Worksheets.Item($i).Name
                Write-Host "  '$sheetName'" -ForegroundColor Red
            }
        } catch {
            Write-Host "Error listing sheets in error message: $_" -ForegroundColor Red
        }
        throw "Inputs sheet not found. Available sheets listed above."
    }
    
    # First, read the current payment method from Solar Projections sheet
    Write-Host "Reading current payment method from Excel file..." -ForegroundColor Green
    $currentPaymentMethod = "Hometree"  # Default fallback
    
    try {
        # Find the Solar Projections sheet to read current payment method
        $solarProjectionsSheet = $null
        $sheetCount = $workbook.Worksheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Solar Projections") {
                $solarProjectionsSheet = $workbook.Worksheets.Item($i)
                break
            }
        }
        
        if ($solarProjectionsSheet) {
            $paymentTypeValue = $solarProjectionsSheet.Range("B7").Value2
            if ($paymentTypeValue) {
                $currentPaymentMethod = $paymentTypeValue.ToString()
                Write-Host "Current payment method from B7: $currentPaymentMethod" -ForegroundColor Green
            } else {
                Write-Host "No payment method found in B7, using default: Hometree" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Solar Projections sheet not found, using default: Hometree" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Could not read current payment method: $_, using default: Hometree" -ForegroundColor Yellow
    }
    
    # Now ensure the current payment method is selected (preserve existing selection)
    Write-Host "Ensuring payment method '$currentPaymentMethod' is selected for calculations..." -ForegroundColor Green
    
    # Try to use the appropriate macro based on current payment method
    try {
        Write-Host "Using macro approach for payment selection..." -ForegroundColor Green
        $inputsSheet.Select()
        
        # Run the appropriate macro based on current payment method
        switch ($currentPaymentMethod.ToLower()) {
            "hometree" { 
                $excel.Run("SetOptionHomeTree")
                Write-Host "✅ SetOptionHomeTree macro executed" -ForegroundColor Green
            }
            "cash" { 
                $excel.Run("SetOptionCash")
                Write-Host "✅ SetOptionCash macro executed" -ForegroundColor Green
            }
            "finance" { 
                $excel.Run("SetOptionNewFinance")
                Write-Host "✅ SetOptionNewFinance macro executed" -ForegroundColor Green
            }
            default { 
                $excel.Run("SetOptionHomeTree")
                Write-Host "✅ SetOptionHomeTree macro executed (default)" -ForegroundColor Green
            }
        }
        
        # Wait for macro to complete and calculations to settle
        Start-Sleep -Milliseconds 1500
        # Ensure calculation is triggered after macro
        $excel.Calculate()
        Start-Sleep -Milliseconds 500
    } catch {
        Write-Host "⚠️ Macro approach failed: $_" -ForegroundColor Yellow
        
        # Fallback: Set values directly in Lookups sheet based on current payment method
        Write-Host "Fallback: Setting values directly in Lookups sheet..." -ForegroundColor Yellow
        try {
            $lookupsSheet = $workbook.Worksheets.Item("Lookups")
            if ($lookupsSheet) {
                # Set PaymentMethodOptionValue based on current payment method
                $paymentMethodValue = switch ($currentPaymentMethod.ToLower()) {
                    "hometree" { 2 }
                    "cash" { 1 }
                    "finance" { 3 }
                    default { 2 }  # Default to Hometree
                }
                $lookupsSheet.Range("PaymentMethodOptionValue").Value2 = $paymentMethodValue
                Write-Host "Set Lookups PaymentMethodOptionValue to $paymentMethodValue ($currentPaymentMethod)" -ForegroundColor Green
                
                # Set NewFinance based on current payment method
                $newFinanceValue = if ($currentPaymentMethod.ToLower() -eq "finance") { "True" } else { "False" }
                $lookupsSheet.Range("NewFinance").Value2 = $newFinanceValue
                Write-Host "Set Lookups NewFinance to $newFinanceValue" -ForegroundColor Green
                
                # Calculate the Lookups sheet and wait for completion
                $lookupsSheet.Calculate()
                Start-Sleep -Milliseconds 1000
                # Also trigger full calculation to ensure dependencies
                $excel.Calculate()
                Start-Sleep -Milliseconds 500
            }
        } catch {
            Write-Host "Fallback method also failed: $_" -ForegroundColor Red
        }
    }
    
    # Now set the terms value
    $termsCell = if ($calculatorType -eq "off-peak") { "H84" } else { "H85" }
    Write-Host "Setting terms: $terms in $termsCell" -ForegroundColor Green
    
    try {
        $inputsSheet.Range($termsCell).Value2 = $terms
        Write-Host "Set terms in $termsCell to $terms" -ForegroundColor Green
        $termsSet = $true
    } catch {
        Write-Host "Could not set terms in $($termsCell): $_" -ForegroundColor Yellow
        $termsSet = $false
    }
    
    if (-not $termsSet) {
        Write-Host "Warning: Could not set terms in $termsCell" -ForegroundColor Red
    }
    
    # Calculate the workbook and wait for calculation to complete
    Write-Host "Calculating workbook..." -ForegroundColor Green
    if ($solarProjectionsSheet) {
        # First calculate the entire workbook to ensure all dependencies are ready
        # This is especially important on the first run
        Write-Host "Calculating entire workbook to ensure dependencies are ready..." -ForegroundColor Yellow
        $excel.Calculate()
        
        # Wait for calculation to complete by checking CalculationState
        Write-Host "Waiting for calculation to complete..." -ForegroundColor Yellow
        $maxWaitTime = 10  # Maximum wait time in seconds (reduced from 30)
        $waitInterval = 0.5  # Check every 500ms
        $elapsed = 0
        
        while ($elapsed -lt $maxWaitTime) {
            $calcState = $excel.CalculationState
            # 0 = xlDone (calculation complete)
            if ($calcState -eq 0) {
                Write-Host "Calculation completed after $elapsed seconds" -ForegroundColor Green
                break
            }
            Start-Sleep -Milliseconds ($waitInterval * 1000)
            $elapsed += $waitInterval
        }
        
        if ($elapsed -ge $maxWaitTime) {
            Write-Host "Warning: Calculation may not be complete after $maxWaitTime seconds, proceeding anyway..." -ForegroundColor Yellow
        }
        
        # Additional safety wait to ensure values are fully updated (reduced from 1000ms)
        Start-Sleep -Milliseconds 500
    } else {
        # Fallback: calculate entire workbook only if sheet not found
        Write-Host "Solar Projections sheet not found, calculating entire workbook..." -ForegroundColor Yellow
        $excel.Calculate()
        
        # Wait for calculation to complete
        $maxWaitTime = 10  # Maximum wait time in seconds (reduced from 30)
        $waitInterval = 0.5
        $elapsed = 0
        
        while ($elapsed -lt $maxWaitTime) {
            $calcState = $excel.CalculationState
            if ($calcState -eq 0) {
                Write-Host "Calculation completed after $elapsed seconds" -ForegroundColor Green
                break
            }
            Start-Sleep -Milliseconds ($waitInterval * 1000)
            $elapsed += $waitInterval
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # Now extract the solar projection data from the same Excel session
    Write-Host "Extracting solar projection data from the same Excel session..." -ForegroundColor Green
    
    # Find the Solar Projections sheet
    $solarProjectionsSheet = $null
    try {
        $sheetCount = $workbook.Worksheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            if ($sheetName -eq "Solar Projections") {
                $solarProjectionsSheet = $workbook.Worksheets.Item($i)
                Write-Host "Found Solar Projections sheet: $sheetName" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "Error finding Solar Projections sheet: $_" -ForegroundColor Red
    }
    
    if ($solarProjectionsSheet) {
        Write-Host "Extracting summary data from sheet: Solar Projections" -ForegroundColor Green
        
        # Extract summary data
        $summary = @{}
        
        # Extract payment type from B7
        try {
            $paymentTypeValue = $solarProjectionsSheet.Range("B7").Value2
            $summary.paymentType = if ($paymentTypeValue) { $paymentTypeValue.ToString() } else { "Hometree" }
            Write-Host "Extracted paymentType from Excel B7: $($summary.paymentType)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract paymentType from B7: $_" -ForegroundColor Yellow
            $summary.paymentType = "Hometree"
        }
        
        # Extract yearly saving year 1 from B20
        try {
            $cellValue = $solarProjectionsSheet.Range("B20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlySavingYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlySavingYear1 = $null
            }
            Write-Host "Extracted yearlySavingYear1 from B20: $($summary.yearlySavingYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlySavingYear1 from B20: $_" -ForegroundColor Yellow
            $summary.yearlySavingYear1 = $null
        }
        
        # Extract yearly plan cost from Yearly Payments column (column 12, row 3)
        try {
            $cellValue = $solarProjectionsSheet.Cells.Item(3, 12).Value2
            if ($cellValue -ne $null) {
                $cellText = $cellValue.ToString()
                $cleaned = [System.Text.RegularExpressions.Regex]::Replace($cellText, "[^0-9.\-]", "")
                $numericValue = [double]::Parse($cleaned, [System.Globalization.CultureInfo]::InvariantCulture)
                # Truncate to 2 decimal places (no rounding)
                $truncatedValue = [math]::Floor($numericValue * 100) / 100
                $summary.yearlyPlanCost = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyPlanCost = $null
            }
            Write-Host "Extracted yearlyPlanCost from Yearly Payments column (column 12, row 3): $($summary.yearlyPlanCost)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlyPlanCost from Yearly Payments column: $_" -ForegroundColor Yellow
            $summary.yearlyPlanCost = $null
        }
        
        # Extract lifetime profit from B25
        try {
            $cellValue = $solarProjectionsSheet.Range("B25").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.lifetimeProfit = $truncatedValue.ToString("F2")
            } else {
                $summary.lifetimeProfit = $null
            }
            Write-Host "Extracted lifetimeProfit from B25: $($summary.lifetimeProfit)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract lifetimeProfit from B25: $_" -ForegroundColor Yellow
            $summary.lifetimeProfit = $null
        }
        
        # Extract term from B10
        try {
            $termValue = $solarProjectionsSheet.Range("B10").Value2
            $summary.term = if ($termValue) { $termValue.ToString() } else { $terms.ToString() }
            Write-Host "Extracted term from B10: $($summary.term)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract term from B10: $_" -ForegroundColor Yellow
            $summary.term = $terms.ToString()
        }
        
        # Extract yearly contribution year 1 from D20
        try {
            $cellValue = $solarProjectionsSheet.Range("D20").Value2
            if ($cellValue -ne $null) {
                $numericValue = [double]$cellValue
                # Truncate to 2 decimal places (no rounding)
            $truncatedValue = [math]::Floor($numericValue * 100) / 100
            $summary.yearlyContributionYear1 = $truncatedValue.ToString("F2")
            } else {
                $summary.yearlyContributionYear1 = $null
            }
            Write-Host "Extracted yearlyContributionYear1 from D20: $($summary.yearlyContributionYear1)" -ForegroundColor Green
        } catch {
            Write-Host "Could not extract yearlyContributionYear1 from D20: $_" -ForegroundColor Yellow
            $summary.yearlyContributionYear1 = $null
        }
        
        # Set calculator type
        $summary.calculatorType = if ($calculatorType -eq "off-peak") { "Off Peak" } else { $calculatorType.ToUpper() }
        
        # Extract table headers (row 2, columns 6-14 - all data columns)
        Write-Host "Extracting table headers from row 2, columns 6-14..." -ForegroundColor Yellow
        $headers = @()
        for ($col = 6; $col -le 14; $col++) {
            try {
                $headerValue = $solarProjectionsSheet.Cells.Item(2, $col).Value2
                if ($headerValue -ne $null) {
                    $headers += $headerValue.ToString()
                    Write-Host "  Header $col : $($headerValue.ToString())" -ForegroundColor Green
                } else {
                    $headers += ""
                    Write-Host "  Header $col : (empty)" -ForegroundColor Gray
                }
            } catch {
                Write-Host "  Warning: Could not extract header from column $col : $_" -ForegroundColor Yellow
                $headers += ""
            }
        }

        # Extract table rows (rows 3-32, columns 6-14 - all data columns for 30 years)
        Write-Host "Extracting table rows from rows 3-32, columns 6-14..." -ForegroundColor Yellow
        $rows = @()
        for ($row = 3; $row -le 32; $row++) {
            $rowData = @()
            $hasData = $false
            for ($col = 6; $col -le 14; $col++) {
                try {
                    $cellValue = $solarProjectionsSheet.Cells.Item($row, $col).Value2
                    if ($cellValue -ne $null) {
                        if ($cellValue -is [double]) {
                            # Truncate to 2 decimal places (no rounding)
                        $truncatedValue = [math]::Floor([double]$cellValue * 100) / 100
                        $rowData += $truncatedValue.ToString("F2")
                        } else {
                            $rowData += $cellValue.ToString()
                        }
                        $hasData = $true
                    } else {
                        $rowData += ""
                    }
                } catch {
                    Write-Host "  Warning: Could not extract cell ($row,$col): $_" -ForegroundColor Yellow
                    $rowData += ""
                }
            }
            
            if ($hasData) {
                $rows += ,$rowData
                Write-Host "  Row $row : $($rowData -join ', ')" -ForegroundColor Green
            } else {
                Write-Host "  Row $row : (no data)" -ForegroundColor Gray
            }
        }

        # Create the result object
        $result = @{
            title = "Lifetime Savings Projections - $($summary.calculatorType)"
            summary = $summary
            table = @{
                headers = $headers
                rows = $rows
                totalRows = $rows.Count
            }
        }

        # Convert to JSON and output
        $jsonResult = $result | ConvertTo-Json -Depth 10
        Write-Host "RESULT: $jsonResult" -ForegroundColor Green
        
    } else {
        Write-Host "Solar Projections sheet not found" -ForegroundColor Red
        throw "Solar Projections sheet not found"
    }
    
    # Skip saving the workbook - we're just reading data, no need to persist changes
    # This prevents "file already exists" dialogs and improves performance
    Write-Host "Skipping workbook save (read-only operation)..." -ForegroundColor Yellow
    
    # Close Excel properly to release file lock
    Write-Host "Closing Excel application to release file lock..." -ForegroundColor Yellow
    try {
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel application closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Error closing Excel: $_" -ForegroundColor Yellow
    }
    
    Write-Host "Terms update and data extraction completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "Critical error in terms update: $($_.Exception.Message)"
    exit 1
}
      `;
      
      const scriptPath = path.join(process.cwd(), `temp_terms_update_${Date.now()}.ps1`);
      await fsPromises.writeFile(scriptPath, scriptContent);
      
      try {
        const { stdout, stderr } = await this.execAsync(`powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`);
        
        if (stderr && stderr.trim()) {
          console.error('PowerShell stderr:', stderr);
        }
        
        return {
          success: true,
          output: stdout
        };
      } finally {
        // Clean up the temporary script file
        try {
          await fsPromises.unlink(scriptPath);
        } catch (cleanupError) {
          console.warn('Failed to clean up temporary script file:', cleanupError);
        }
      }
    } catch (error) {
      console.error('❌ Error setting terms in Excel:', error);
      throw error;
    }
  }

  /**
   * Parse JSON result from PowerShell output
   */
  private parsePowerShellJsonResult(output: string): any {
    try {
      const resultIndex = output.indexOf('RESULT:');
      if (resultIndex >= 0) {
        const firstBraceIndex = output.indexOf('{', resultIndex);
        if (firstBraceIndex >= 0) {
          let depth = 0;
          let endIndex = -1;
          for (let i = firstBraceIndex; i < output.length; i++) {
            const ch = output[i];
            if (ch === '{') {
              depth++;
            } else if (ch === '}') {
              depth--;
              if (depth === 0) {
                endIndex = i;
                break;
              }
            }
          }

          if (endIndex > firstBraceIndex) {
            const jsonString = output.substring(firstBraceIndex, endIndex + 1);
            console.log('Extracted JSON string:', jsonString.substring(0, 200) + (jsonString.length > 200 ? '...' : ''));
            const jsonData = JSON.parse(jsonString);
            console.log('✅ Successfully parsed JSON result');

            // Normalize numeric fields so the frontend receives numbers instead of formatted strings
            const coerceNumber = (value: any) => {
              if (value === null || value === undefined || value === '') {
                return null;
              }
              if (typeof value === 'number') {
                return value;
              }
              const cleaned = Number(String(value).replace(/[^0-9.-]/g, ''));
              return Number.isNaN(cleaned) ? null : cleaned;
            };

            if (jsonData?.summary) {
              jsonData.summary.yearlySavingYear1 = coerceNumber(jsonData.summary.yearlySavingYear1);
              jsonData.summary.yearlyPlanCost = coerceNumber(jsonData.summary.yearlyPlanCost);
              jsonData.summary.monthlyPlanCost = coerceNumber(jsonData.summary.monthlyPlanCost);
              jsonData.summary.yearlyContributionYear1 = coerceNumber(jsonData.summary.yearlyContributionYear1);
              jsonData.summary.lifetimeProfit = coerceNumber(jsonData.summary.lifetimeProfit);
              jsonData.summary.totalSavings = coerceNumber(jsonData.summary.totalSavings);
            }

            if (jsonData?.table?.rows) {
              jsonData.table.rows = jsonData.table.rows.map((row: any[]) =>
                row.map((cell: any) => (cell === null || cell === undefined ? '' : cell.toString()))
              );
              jsonData.table.totalRows = Array.isArray(jsonData.table.rows)
                ? jsonData.table.rows.length
                : 0;
            }

            console.log('🔍 Parsed table headers count:', jsonData?.table?.headers?.length ?? 0);
            console.log('🔍 Parsed table rows:', jsonData?.table?.rows?.length ?? 0);
            return jsonData;
          }
        }
      }

      console.error('❌ Could not locate complete JSON block in PowerShell output');
      console.error('PowerShell output preview:', output.substring(resultIndex >= 0 ? resultIndex : 0, Math.min(output.length, (resultIndex >= 0 ? resultIndex : 0) + 500)));
      return null;
    } catch (parseError) {
      console.error('❌ Failed to parse JSON result:', parseError);
      console.error('PowerShell output causing parse error:', output.substring(output.indexOf('RESULT:')));
      return null;
    }
  }

}
