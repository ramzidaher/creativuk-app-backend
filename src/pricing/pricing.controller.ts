import { Controller, Post, Body, Param, Get } from '@nestjs/common';
import { PricingService } from './pricing.service';

@Controller('pricing')
export class PricingController {
  constructor(private readonly pricingService: PricingService) {}

  /**
   * Save pricing with session isolation
   */
  @Post('session/save')
  async savePricingWithSession(
    @Body() body: {
      userId: string;
      opportunityId: string;
      pricingData: {
        batteryType: string;
        numberOfPanels: number;
        additionalItems: string[];
        totalSystemCost: number;
        targetCell: string;
        calculatorType: string;
      };
    }
  ) {
    try {
      const result = await this.pricingService.savePricingToExcelWithSession(
        body.userId,
        body.opportunityId,
        body.pricingData
      );
      
      return {
        success: true,
        message: 'Pricing saved successfully with session isolation',
        data: result
      };
    } catch (error) {
      return {
        success: false,
        message: 'Failed to save pricing with session isolation',
        error: error.message
      };
    }
  }

  @Post(':opportunityId/save')
  async savePricing(
    @Param('opportunityId') opportunityId: string,
    @Body() pricingData: {
      batteryType: string;
      numberOfPanels: number;
      additionalItems: string[];
      totalSystemCost: number;
      targetCell: string;
      calculatorType: string;
    }
  ) {
    try {
      const result = await this.pricingService.savePricingToExcel(
        opportunityId,
        pricingData
      );
      
      return {
        success: true,
        message: 'Pricing saved successfully to Excel',
        data: result
      };
    } catch (error) {
      console.error('Error saving pricing:', error);
      return {
        success: false,
        message: error.message || 'Failed to save pricing',
        error: error.toString()
      };
    }
  }

  @Get(':opportunityId/system-costs-inputs')
  async getSystemCostsInputs(@Param('opportunityId') opportunityId: string) {
    try {
      const inputs = await this.pricingService.getSystemCostsInputs(opportunityId);
      
      return {
        success: true,
        message: 'System costs inputs retrieved successfully',
        data: inputs
      };
    } catch (error) {
      console.error('Error getting system costs inputs:', error);
      return {
        success: false,
        message: error.message || 'Failed to get system costs inputs',
        error: error.toString()
      };
    }
  }
}
