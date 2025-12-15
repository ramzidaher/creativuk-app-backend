import { Controller, Post, Get, Body, Param, Query, HttpException, HttpStatus } from '@nestjs/common';
import { SolarProjectionService } from './solar-projection.service';

@Controller('presentation/solar-projection')
export class SolarProjectionController {
  constructor(private readonly solarProjectionService: SolarProjectionService) {}

  /**
   * Get solar projection data for an opportunity
   */
  @Get(':opportunityId')
  async getSolarProjection(
    @Param('opportunityId') opportunityId: string,
    @Query('calculatorType') calculatorType: string,
    @Query('fileName') fileName?: string
  ) {
    return await this.solarProjectionService.getSolarProjectionData(
      opportunityId,
      calculatorType,
      fileName
    );
  }

  /**
   * Update payment method for solar projection
   */
  @Post(':opportunityId/payment-method')
  async updatePaymentMethod(
    @Param('opportunityId') opportunityId: string,
    @Body() body: { paymentMethod: string; calculatorType: string }
  ) {
    return await this.solarProjectionService.updatePaymentMethod(
      opportunityId,
      body.paymentMethod,
      body.calculatorType
    );
  }

  /**
   * Update terms for solar projection
   */
  @Post(':opportunityId/terms')
  async updateTerms(
    @Param('opportunityId') opportunityId: string,
    @Body() body: { terms: number; calculatorType: string }
  ) {
    return await this.solarProjectionService.updateTerms(
      opportunityId,
      body.terms,
      body.calculatorType
    );
  }
}


