import { Controller, Post, Body, HttpStatus, HttpException, Get, Logger, Param, Res } from '@nestjs/common';
import { Response } from 'express';
import { ExcelAutomationService } from './excel-automation.service';
import { SelectRadioButtonDto, RadioButtonResponseDto } from './dto/radio-button.dto';
import * as path from 'path';
import * as fs from 'fs';

@Controller('excel-automation')
export class ExcelAutomationController {
  private readonly logger = new Logger(ExcelAutomationController.name);
  
  constructor(private readonly excelAutomationService: ExcelAutomationService) {}

  /**
   * Perform Excel calculation with session isolation
   */
  @Post('session/calculate')
  async performExcelCalculationWithSession(
    @Body() body: {
      userId: string;
      opportunityId: string;
      customerDetails: { customerName: string; address: string; postcode: string };
      radioButtonSelections: string[];
      dynamicInputs?: Record<string, string>;
      templateFileName?: string;
    }
  ) {
    const result = await this.excelAutomationService.performCompleteCalculationWithSession(
      body.userId,
      body.opportunityId,
      body.customerDetails,
      body.radioButtonSelections,
      body.dynamicInputs,
      body.templateFileName
    );

    return result;
  }

  @Get('test')
  async test() {
    return {
      message: 'Excel Automation API is working!',
      timestamp: new Date().toISOString(),
      availableEndpoints: [
        'POST /excel-automation/session/calculate',
        'POST /excel-automation/select-radio-button',
        'GET /excel-automation/test',
        'POST /excel-automation/energy-use/single-rate',
        'POST /excel-automation/energy-use/dual-rate',
        'POST /excel-automation/battery/self-consumption',
        'POST /excel-automation/battery/overnight-charging',
        'POST /excel-automation/battery/none',
        'POST /excel-automation/existing-solar/yes',
        'POST /excel-automation/existing-solar/no',
        'POST /excel-automation/annual-consumption/yes',
        'POST /excel-automation/annual-consumption/no',
        'POST /excel-automation/export-tariff/yes',
        'POST /excel-automation/export-tariff/no',
        'POST /excel-automation/warranty/battery/yes',
        'POST /excel-automation/warranty/battery/no',
        'POST /excel-automation/warranty/solar-inverter/yes',
        'POST /excel-automation/warranty/solar-inverter/no',
        'POST /excel-automation/warranty/battery-inverter/yes',
        'POST /excel-automation/warranty/battery-inverter/no',
        'POST /excel-automation/payment/cash',
        'POST /excel-automation/payment/finance',
        'POST /excel-automation/payment/new-finance'
      ]
    };
  }

  @Post('select-radio-button')
  async selectRadioButton(@Body() selectRadioButtonDto: SelectRadioButtonDto): Promise<RadioButtonResponseDto> {
    try {
      const result = await this.excelAutomationService.selectRadioButton(selectRadioButtonDto.shapeName, selectRadioButtonDto.opportunityId);

      if (result.success) {
        return {
          success: true,
          message: result.message,
          shapeName: selectRadioButtonDto.shapeName
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
            shapeName: selectRadioButtonDto.shapeName
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error',
          error: error.message,
          shapeName: selectRadioButtonDto.shapeName
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('create-opportunity-with-template')
  async createOpportunityWithTemplate(@Body() body: { 
    opportunityId: string; 
    customerDetails: { customerName: string; address: string; postcode: string; };
    templateFileName: string;
  }): Promise<{
    success: boolean;
    message: string;
    filePath?: string;
  }> {
    try {
      const { opportunityId, customerDetails, templateFileName } = body;
      
      if (!opportunityId || !customerDetails || !templateFileName) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId, customerDetails, and templateFileName are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.excelAutomationService.createOpportunityWithTemplate(
        opportunityId, 
        customerDetails, 
        templateFileName
      );

      if (result.success) {
        return {
          success: true,
          message: result.message,
          filePath: result.filePath
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('create-opportunity-file')
  async createOpportunityFile(@Body() body: { 
    opportunityId: string; 
    customerDetails: { customerName: string; address: string; postcode: string; };
    templateFileName?: string;
    isTemplateSelection?: boolean;
  }): Promise<{
    success: boolean;
    message: string;
    filePath?: string;
  }> {
    try {
      const { opportunityId, customerDetails, templateFileName, isTemplateSelection = false } = body;
      
      // Auto-detect template selection: if templateFileName is provided and no isTemplateSelection is explicitly set to false
      // This handles the case where frontend calls create-opportunity-file for template selection
      const isActuallyTemplateSelection = isTemplateSelection !== false && !!templateFileName;
      
      this.logger.log(`Creating opportunity file for: ${opportunityId} with template: ${templateFileName || 'default'} (isTemplateSelection: ${isActuallyTemplateSelection})`);
      this.logger.log(`Customer details received: ${JSON.stringify(customerDetails)}`);
      
      if (!opportunityId || !customerDetails) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId and customerDetails are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.excelAutomationService.createOpportunityFile(opportunityId, customerDetails, templateFileName, isActuallyTemplateSelection);

      if (result.success) {
        return {
          success: true,
          message: result.message,
          filePath: result.filePath
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during file creation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('check-opportunity-file')
  async checkOpportunityFile(@Body() body: { 
    opportunityId: string;
  }): Promise<{
    success: boolean;
    exists: boolean;
    message: string;
    filePath?: string;
  }> {
    try {
      const { opportunityId } = body;
      
      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            exists: false,
            message: 'Invalid input: opportunityId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.excelAutomationService.checkOpportunityFileExists(opportunityId);

      return {
        success: true,
        exists: result.exists,
        message: result.message,
        filePath: result.filePath
      };
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          exists: false,
          message: 'Internal server error during file check',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('select-radio-buttons-batch')
  async selectRadioButtonsBatch(@Body() body: { shapeNames: string[]; opportunityId?: string }): Promise<{
    success: boolean;
    message: string;
    results: Array<{ shapeName: string; success: boolean; message: string }>;
  }> {
    try {
      const { shapeNames, opportunityId } = body;
      
      if (!shapeNames || !Array.isArray(shapeNames) || shapeNames.length === 0) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: shapeNames array is required and must not be empty',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const results: Array<{ shapeName: string; success: boolean; message: string }> = [];
      let allSuccessful = true;

      for (const shapeName of shapeNames) {
        try {
          const result = await this.excelAutomationService.selectRadioButton(shapeName, opportunityId);
          results.push({
            shapeName,
            success: result.success,
            message: result.message
          });
          
          if (!result.success) {
            allSuccessful = false;
          }
        } catch (error) {
          results.push({
            shapeName,
            success: false,
            message: error.message || 'Unknown error'
          });
          allSuccessful = false;
        }
      }

      return {
        success: allSuccessful,
        message: allSuccessful 
          ? `Successfully applied ${shapeNames.length} radio button selections`
          : `Applied ${shapeNames.length} selections with some errors`,
        results
      };

    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during batch operation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  // Energy Use Radio Buttons
  @Post('energy-use/single-rate')
  async setSingleRate(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'SingleRate' });
  }

  @Post('energy-use/dual-rate')
  async setDualRate(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'DualRate' });
  }

  // Battery Type Radio Buttons
  @Post('battery/self-consumption')
  async setBatterySelfConsumption(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'BatterySC' });
  }

  @Post('battery/overnight-charging')
  async setBatteryOvernightCharging(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'BatteryOC' });
  }

  @Post('battery/none')
  async setBatteryNone(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'BatteryNone' });
  }

  // Existing Solar Radio Buttons
  @Post('existing-solar/yes')
  async setExistingSolarYes(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'ExistingSolarYes' });
  }

  @Post('existing-solar/no')
  async setExistingSolarNo(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'ExistingSolarNo' });
  }

  // Annual Consumption Radio Buttons
  @Post('annual-consumption/yes')
  async setAnnualConsumptionYes(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'AnnualConsumptionYes' });
  }

  @Post('annual-consumption/no')
  async setAnnualConsumptionNo(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'AnnualConsumptionNo' });
  }

  // Export Tariff Radio Buttons
  @Post('export-tariff/yes')
  async setExportTariffYes(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'ExportYes' });
  }

  @Post('export-tariff/no')
  async setExportTariffNo(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'ExportNo' });
  }

  // Warranty Radio Buttons
  @Post('warranty/battery/yes')
  async setBatteryWarrantyYes(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'BatteryWarrantyYes' });
  }

  @Post('warranty/battery/no')
  async setBatteryWarrantyNo(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'BatteryWarrantyNo' });
  }

  @Post('warranty/solar-inverter/yes')
  async setSolarInverterWarrantyYes(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'SolarInverterWarrantyYes' });
  }

  @Post('warranty/solar-inverter/no')
  async setSolarInverterWarrantyNo(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'SolarInverterWarrantyNo' });
  }

  @Post('warranty/battery-inverter/yes')
  async setBatteryInverterWarrantyYes(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'BatteryInverterWarrantyYes' });
  }

  @Post('warranty/battery-inverter/no')
  async setBatteryInverterWarrantyNo(): Promise<RadioButtonResponseDto> {
    return this.selectRadioButton({ shapeName: 'BatteryInverterWarrantyNo' });
  }

  // Payment Method Radio Buttons
  @Post('payment/cash')
  async setPaymentCash(@Body() body: { opportunityId?: string }): Promise<RadioButtonResponseDto> {
    try {
      const result = await this.excelAutomationService.selectRadioButton('Cash', body.opportunityId);
      
      if (result.success) {
        return {
          success: true,
          message: result.message,
          shapeName: 'Cash'
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
            shapeName: 'Cash'
          },
          HttpStatus.INTERNAL_SERVER_ERROR
        );
      }
    } catch (error) {
      throw new HttpException(
        {
          success: false,
          message: 'Failed to select Cash payment method',
          error: error.message,
          shapeName: 'Cash'
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('payment/finance')
  async setPaymentFinance(@Body() body: { opportunityId?: string }): Promise<RadioButtonResponseDto> {
    try {
      const result = await this.excelAutomationService.selectRadioButton('Finance', body.opportunityId);
      
      if (result.success) {
        return {
          success: true,
          message: result.message,
          shapeName: 'Finance'
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
            shapeName: 'Finance'
          },
          HttpStatus.INTERNAL_SERVER_ERROR
        );
      }
    } catch (error) {
      throw new HttpException(
        {
          success: false,
          message: 'Failed to select Finance payment method',
          error: error.message,
          shapeName: 'Finance'
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('payment/new-finance')
  async setPaymentNewFinance(@Body() body: { opportunityId?: string }): Promise<RadioButtonResponseDto> {
    try {
      const result = await this.excelAutomationService.selectRadioButton('NewFinance', body.opportunityId);
      
      if (result.success) {
        return {
          success: true,
          message: result.message,
          shapeName: 'NewFinance'
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
            shapeName: 'NewFinance'
          },
          HttpStatus.INTERNAL_SERVER_ERROR
        );
      }
    } catch (error) {
      throw new HttpException(
        {
          success: false,
          message: 'Failed to select New Finance payment method',
          error: error.message,
          shapeName: 'NewFinance'
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }



  @Post('perform-complete-calculation')
  async performCompleteCalculation(@Body() body: {
    opportunityId: string;
    customerDetails: { customerName: string; address: string; postcode: string };
    radioButtonSelections: string[];
    dynamicInputs?: Record<string, string>;
    templateFileName?: string;
  }): Promise<{
    success: boolean;
    message: string;
    filePath?: string;
    error?: string;
  }> {
    try {
      const { opportunityId, customerDetails, radioButtonSelections, dynamicInputs, templateFileName } = body;
      
      if (!opportunityId || !customerDetails || !radioButtonSelections) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId, customerDetails, and radioButtonSelections are required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.excelAutomationService.performCompleteCalculation(
        opportunityId,
        customerDetails,
        radioButtonSelections,
        dynamicInputs,
        templateFileName
      );

      if (result.success) {
        return {
          success: true,
          message: result.message,
          filePath: result.filePath
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during complete calculation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('get-dynamic-inputs')
  async getDynamicInputs(@Body() body: { 
    opportunityId?: string; 
    templateFileName?: string;
  }): Promise<{
    success: boolean;
    message: string;
    inputFields?: any[];
    error?: string;
  }> {
    try {
      const { opportunityId, templateFileName } = body;
      this.logger.log(`🔍 Getting dynamic inputs for opportunity: ${opportunityId}, templateFileName: ${templateFileName}`);
      const result = await this.excelAutomationService.getDynamicInputs(opportunityId, templateFileName);
      this.logger.log(`📡 Dynamic inputs result: ${JSON.stringify(result)}`);

      if (result.success) {
        return {
          success: true,
          message: result.message,
          inputFields: result.inputFields
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during dynamic inputs retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('get-dropdown-options')
  async getDropdownOptions(@Body() body: { 
    opportunityId?: string; 
    fieldId: string;
    dependsOnValue?: string;
  }): Promise<{
    success: boolean;
    message: string;
    options?: string[];
    error?: string;
  }> {
    try {
      const { opportunityId, fieldId, dependsOnValue } = body;
      
      if (!fieldId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: fieldId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.excelAutomationService.getCascadingDropdownOptions(opportunityId, fieldId, dependsOnValue);

      if (result.success) {
        return {
          success: true,
          message: result.message,
          options: result.options
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during dropdown options retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('get-all-dropdown-options')
  async getAllDropdownOptions(@Body() body: { 
    opportunityId?: string; 
    templateFileName?: string;
  }): Promise<{
    success: boolean;
    message: string;
    dropdownOptions?: Record<string, string[]>;
    error?: string;
  }> {
    try {
      const { opportunityId, templateFileName } = body;
      
      const result = await this.excelAutomationService.getAllDropdownOptionsForFrontend(opportunityId, templateFileName);

      if (result.success) {
        return {
          success: true,
          message: result.message,
          dropdownOptions: result.dropdownOptions
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during dropdown options retrieval',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('save-dynamic-inputs')
  async saveDynamicInputs(@Body() body: { 
    opportunityId?: string; 
    templateFileName?: string;
    inputs: Record<string, string>;
    calculatorType?: 'flux' | 'off-peak';
  }): Promise<{
    success: boolean;
    message: string;
    error?: string;
  }> {
    try {
      const { opportunityId, templateFileName, inputs, calculatorType } = body;
      
      if (!inputs || Object.keys(inputs).length === 0) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: inputs object is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.excelAutomationService.saveDynamicInputs(opportunityId, inputs, templateFileName, calculatorType);

      if (result.success) {
        return {
          success: true,
          message: result.message
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during dynamic inputs save',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Post('generate-pdf')
  async generatePDF(@Body() body: { 
    opportunityId: string; 
    excelFilePath?: string;
    signatureData?: string;
    fileName?: string;
    selectedSheet?: string; // Frontend sends this parameter
  }): Promise<{
    success: boolean;
    message: string;
    pdfPath?: string;
    pdfUrl?: string;
    error?: string;
  }> {
    try {
      // Map selectedSheet to fileName if provided (frontend sends selectedSheet)
      const fileName = body.selectedSheet || body.fileName;
      const { opportunityId, excelFilePath, signatureData } = body;
      
      if (!opportunityId) {
        throw new HttpException(
          {
            success: false,
            message: 'Invalid input: opportunityId is required',
          },
          HttpStatus.BAD_REQUEST
        );
      }

      const result = await this.excelAutomationService.generatePDF(opportunityId, excelFilePath, signatureData, fileName);

      if (result.success) {
        // Create a URL for the PDF that can be accessed via HTTP
        const pdfUrl = `/excel-automation/pdf/${opportunityId}`;
        
        return {
          success: true,
          message: result.message,
          pdfPath: result.pdfPath,
          pdfUrl: pdfUrl
        };
      } else {
        throw new HttpException(
          {
            success: false,
            message: result.message,
            error: result.error,
          },
          HttpStatus.BAD_REQUEST
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Internal server error during PDF generation',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  @Get('pdf/:opportunityId')
  async servePDF(@Param('opportunityId') opportunityId: string, @Res() res: Response) {
    try {
      // Try different possible PDF paths
      const possiblePdfPaths = [
        path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities', 'pdfs', `Off Peak Calculator - ${opportunityId}.pdf`),
        path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities', 'pdfs', `EPVS Calculator - ${opportunityId}.pdf`),
        path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities', 'pdfs', `EPVS Calculator Creativ - 06.02 - ${opportunityId}.pdf`),
      ];
      
      let pdfPath: string | null = null;
      for (const possiblePath of possiblePdfPaths) {
        if (fs.existsSync(possiblePath)) {
          pdfPath = possiblePath;
          break;
        }
      }
      
      if (!pdfPath) {
        // If no specific PDF found, search for any PDF with the opportunity ID
        const searchFolders = [
          path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities', 'pdfs'),
          path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities', 'pdfs'),
        ];
        
        for (const folder of searchFolders) {
          if (fs.existsSync(folder)) {
            const files = fs.readdirSync(folder);
            const matchingFile = files.find(file => file.includes(opportunityId) && file.endsWith('.pdf'));
            
            if (matchingFile) {
              pdfPath = path.join(folder, matchingFile);
              break;
            }
          }
        }
      }
      
      if (!pdfPath || !fs.existsSync(pdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'PDF file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `inline; filename="Calculator - ${opportunityId}.pdf"`);
      res.sendFile(pdfPath);
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }

      throw new HttpException(
        {
          success: false,
          message: 'Error serving PDF file',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}
