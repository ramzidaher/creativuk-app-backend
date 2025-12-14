import { Controller, Post, Body, Get, Param, Query, Res, HttpException, HttpStatus } from '@nestjs/common';
import { Response } from 'express';
import { EPVSAutomationService } from './epvs-automation.service';
import * as path from 'path';
import * as fs from 'fs';

@Controller('epvs-automation')
export class EPVSAutomationController {
  constructor(private readonly epvsAutomationService: EPVSAutomationService) {}

  /**
   * Perform EPVS calculation with session isolation
   */
  @Post('session/calculate')
  async performEPVSCalculationWithSession(
    @Body() body: {
      userId: string;
      opportunityId: string;
      customerDetails: { customerName: string; address: string; postcode: string };
      radioButtonSelections: string[];
      dynamicInputs?: Record<string, string>;
      templateFileName?: string;
    }
  ) {
    const result = await this.epvsAutomationService.performCompleteCalculationWithSession(
      body.userId,
      body.opportunityId,
      body.customerDetails,
      body.radioButtonSelections,
      body.dynamicInputs,
      body.templateFileName
    );

    return result;
  }

  @Post('select-radio-button')
  async selectRadioButton(
    @Body() body: { shapeName: string; opportunityId?: string }
  ) {
    return await this.epvsAutomationService.selectRadioButton(
      body.shapeName,
      body.opportunityId
    );
  }

  @Post('select-multiple-radio-buttons')
  async selectMultipleRadioButtons(
    @Body() body: { shapeNames: string[]; opportunityId?: string }
  ) {
    return await this.epvsAutomationService.selectMultipleRadioButtons(
      body.shapeNames,
      body.opportunityId
    );
  }

  // Energy Use endpoints
  @Post('energy-use/single-rate')
  async selectSingleRate(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('SingleRate', body.opportunityId);
  }

  @Post('energy-use/dual-rate')
  async selectDualRate(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('DualRate', body.opportunityId);
  }

  // Annual Usage endpoints
  @Post('annual-usage/yes')
  async selectAnnualUsageYes(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('AnnualConsumptionYes', body.opportunityId);
  }

  @Post('annual-usage/no')
  async selectAnnualUsageNo(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('AnnualConsumptionNo', body.opportunityId);
  }

  // Existing Customer endpoints
  @Post('existing-customer/yes')
  async selectExistingCustomerYes(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('ExistingSolarYes', body.opportunityId);
  }

  @Post('existing-customer/no')
  async selectExistingCustomerNo(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('ExistingSolarNo', body.opportunityId);
  }

  // Battery Warranty endpoints
  @Post('battery-warranty/yes')
  async selectBatteryWarrantyYes(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('BatteryWarrantyYes', body.opportunityId);
  }

  @Post('battery-warranty/no')
  async selectBatteryWarrantyNo(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('BatteryWarrantyNo', body.opportunityId);
  }

  // Solar/Hybrid Warranty endpoints
  @Post('solar-hybrid-warranty/yes')
  async selectSolarHybridWarrantyYes(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('SolarInverterWarrantyYes', body.opportunityId);
  }

  @Post('solar-hybrid-warranty/no')
  async selectSolarHybridWarrantyNo(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('SolarInverterWarrantyNo', body.opportunityId);
  }

  // Battery Inverter endpoints
  @Post('battery-inverter/yes')
  async selectBatteryInverterYes(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('BatteryInverterWarrantyYes', body.opportunityId);
  }

  @Post('battery-inverter/no')
  async selectBatteryInverterNo(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('BatteryInverterWarrantyNo', body.opportunityId);
  }

  // Payment endpoints
  @Post('payment/cash')
  async selectPaymentCash(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('Cash', body.opportunityId);
  }

  @Post('payment/finance')
  async selectPaymentFinance(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('Finance', body.opportunityId);
  }

  @Post('payment/new-finance')
  async selectPaymentNewFinance(@Body() body: { opportunityId?: string }) {
    return await this.epvsAutomationService.selectRadioButton('NewFinance', body.opportunityId);
  }


  @Post('create-opportunity-file')
  async createOpportunityFile(
    @Body() body: {
      opportunityId: string;
      customerDetails: {
        customerName: string;
        address: string;
        postcode: string;
      };
      templateFileName?: string;
      isTemplateSelection?: boolean;
    }
  ) {
    // Auto-detect template selection: if templateFileName is provided and no isTemplateSelection is explicitly set to false
    // This handles the case where frontend calls create-opportunity-file for template selection
    const isActuallyTemplateSelection = body.isTemplateSelection !== false && !!body.templateFileName;
    
    return await this.epvsAutomationService.createOpportunityFile(
      body.opportunityId,
      body.customerDetails,
      body.templateFileName,
      isActuallyTemplateSelection
    );
  }

  @Post('create-opportunity-with-template')
  async createOpportunityWithTemplate(
    @Body() body: {
      opportunityId: string;
      customerDetails: {
        customerName: string;
        address: string;
        postcode: string;
      };
      templateFileName: string;
    }
  ) {
    return await this.epvsAutomationService.createOpportunityWithTemplate(
      body.opportunityId,
      body.customerDetails,
      body.templateFileName
    );
  }

  @Get('dynamic-inputs')
  async getDynamicInputs(
    @Query('opportunityId') opportunityId?: string,
    @Query('templateFileName') templateFileName?: string
  ) {
    return await this.epvsAutomationService.getDynamicInputs(
      opportunityId,
      templateFileName
    );
  }

  @Post('save-dynamic-inputs')
  async saveDynamicInputs(
    @Body() body: {
      opportunityId?: string;
      inputs: Record<string, string>;
      templateFileName?: string;
    }
  ) {
    return await this.epvsAutomationService.saveDynamicInputs(
      body.opportunityId,
      body.inputs,
      body.templateFileName
    );
  }

  @Post('save-array-data-simple')
  async saveArrayDataSimple(
    @Body() body: {
      opportunityId: string;
      arrayData: {
        no_of_arrays: string;
        [key: string]: string | undefined; // Allow any array field (array1_panels, array2_panels, etc.)
      };
    }
  ) {
    return await this.epvsAutomationService.saveArrayDataSimple(
      body.opportunityId,
      body.arrayData
    );
  }

  @Get('dropdown-options')
  async getAllDropdownOptionsForFrontendNoId(
    @Query('templateFileName') templateFileName?: string
  ) {
    return await this.epvsAutomationService.getAllDropdownOptionsForFrontend(
      undefined,
      templateFileName
    );
  }

  @Get('dropdown-options/:opportunityId')
  async getAllDropdownOptionsForFrontend(
    @Param('opportunityId') opportunityId: string,
    @Query('templateFileName') templateFileName?: string
  ) {
    return await this.epvsAutomationService.getAllDropdownOptionsForFrontend(
      opportunityId,
      templateFileName
    );
  }

  @Get('manufacturer-models/:fieldId/:manufacturer')
  async getManufacturerSpecificModelsNoId(
    @Param('fieldId') fieldId: string,
    @Param('manufacturer') manufacturer: string
  ) {
    return await this.epvsAutomationService.getManufacturerSpecificModels(
      fieldId,
      manufacturer,
      undefined
    );
  }

  @Get('manufacturer-models/:fieldId/:manufacturer/:opportunityId')
  async getManufacturerSpecificModels(
    @Param('fieldId') fieldId: string,
    @Param('manufacturer') manufacturer: string,
    @Param('opportunityId') opportunityId: string
  ) {
    return await this.epvsAutomationService.getManufacturerSpecificModels(
      fieldId,
      manufacturer,
      opportunityId
    );
  }

  @Get('cascading-dropdown/:fieldId')
  async getCascadingDropdownOptionsNoId(
    @Param('fieldId') fieldId: string,
    @Query('dependsOnValue') dependsOnValue?: string
  ) {
    return await this.epvsAutomationService.getCascadingDropdownOptions(
      undefined,
      fieldId,
      dependsOnValue
    );
  }

  @Get('cascading-dropdown/:fieldId/:opportunityId')
  async getCascadingDropdownOptions(
    @Param('fieldId') fieldId: string,
    @Param('opportunityId') opportunityId: string,
    @Query('dependsOnValue') dependsOnValue?: string
  ) {
    return await this.epvsAutomationService.getCascadingDropdownOptions(
      opportunityId,
      fieldId,
      dependsOnValue
    );
  }

  @Get('flux-rates/:postcode')
  async getOctopusFluxRates(@Param('postcode') postcode: string) {
    return await this.epvsAutomationService.getOctopusFluxRates(postcode);
  }

  @Post('populate-flux-rates')
  async populateFluxRatesInExcel(
    @Body() body: { opportunityId: string; postcode: string }
  ) {
    return await this.epvsAutomationService.populateFluxRatesInExcel(
      body.opportunityId,
      body.postcode
    );
  }

  @Post('complete-calculation')
  async performCompleteCalculation(
    @Body() body: {
      opportunityId: string;
      customerDetails: {
        customerName: string;
        address: string;
        postcode: string;
      };
      radioButtonSelections: string[];
      dynamicInputs?: Record<string, string>;
      templateFileName?: string;
    }
  ) {
    return await this.epvsAutomationService.performCompleteCalculation(
      body.opportunityId,
      body.customerDetails,
      body.radioButtonSelections,
      body.dynamicInputs,
      body.templateFileName
    );
  }

  @Post('generate-pdf')
  async generatePDF(
    @Body() body: {
      opportunityId: string;
      excelFilePath?: string;
      signatureData?: string;
      fileName?: string;
      selectedSheet?: string; // Frontend sends this parameter
    }
  ) {
    // Map selectedSheet to fileName if provided (frontend sends selectedSheet)
    const fileName = body.selectedSheet || body.fileName;
    
    const result = await this.epvsAutomationService.generatePDF(
      body.opportunityId,
      body.excelFilePath,
      body.signatureData,
      fileName
    );

    if (result.success) {
      // Create a URL for the PDF that can be accessed via HTTP
      const pdfUrl = `/epvs-automation/pdf/${body.opportunityId}`;
      
      return {
        ...result,
        pdfUrl: pdfUrl
      };
    }

    return result;
  }

  @Get('pdf/:opportunityId')
  async servePDF(@Param('opportunityId') opportunityId: string, @Res() res: Response) {
    try {
      const pdfPath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities', 'pdfs', `EPVS Calculator - ${opportunityId}.pdf`);
      
      if (!fs.existsSync(pdfPath)) {
        throw new HttpException(
          {
            success: false,
            message: 'PDF file not found',
          },
          HttpStatus.NOT_FOUND
        );
      }

      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `inline; filename="EPVS Calculator - ${opportunityId}.pdf"`);
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
