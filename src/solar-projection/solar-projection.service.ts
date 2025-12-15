import { Injectable, Logger } from '@nestjs/common';
import { PresentationService } from '../excel-file-calculator/presentation.service';

@Injectable()
export class SolarProjectionService {
  private readonly logger = new Logger(SolarProjectionService.name);

  constructor(
    private readonly presentationService: PresentationService
  ) {}

  /**
   * Get solar projection data for an opportunity
   */
  async getSolarProjectionData(
    opportunityId: string,
    calculatorType: string,
    fileName?: string
  ): Promise<{ success: boolean; data?: any; error?: string }> {
    try {
      this.logger.log(`Getting solar projection data for opportunity: ${opportunityId}, calculatorType: ${calculatorType}`);

      // Use the presentation service to extract actual solar projection data from Excel
      const solarProjectionData = await this.presentationService.extractSolarProjectionData(
        opportunityId,
        calculatorType as 'flux' | 'off-peak' | 'epvs'
      );

      if (!solarProjectionData) {
        throw new Error('Failed to extract solar projection data from Excel file');
      }

      return {
        success: true,
        data: solarProjectionData
      };
    } catch (error) {
      this.logger.error(`Error getting solar projection data: ${error.message}`);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Update payment method for solar projection
   */
  async updatePaymentMethod(
    opportunityId: string,
    paymentMethod: string,
    calculatorType: string
  ): Promise<{ success: boolean; data?: any; error?: string }> {
    try {
      this.logger.log(`Updating payment method to ${paymentMethod} for opportunity: ${opportunityId}, calculatorType: ${calculatorType}`);

      // Use the PresentationService method that includes the VBA fix
      // This method runs SetOptionHomeTree macro and keeps the sheet open during operations
      const updatedData = await this.presentationService.updatePaymentMethodAndExtractData(
          opportunityId,
        paymentMethod,
        calculatorType as 'flux' | 'off-peak' | 'epvs'
        );
        
        if (!updatedData) {
        throw new Error('Failed to update payment method and extract data');
          }

          return {
            success: true,
            data: updatedData
          };
    } catch (error) {
      this.logger.error(`Error updating payment method: ${error.message}`);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Update terms for solar projection
   */
  async updateTerms(
    opportunityId: string,
    terms: number,
    calculatorType: string
  ): Promise<{ success: boolean; data?: any; error?: string }> {
    try {
      this.logger.log(`Updating terms to ${terms} years for opportunity: ${opportunityId}, calculatorType: ${calculatorType}`);

      // Use the PresentationService method that includes the VBA fix
      // This method runs SetOptionHomeTree macro and keeps the sheet open during operations
      const updatedData = await this.presentationService.updateTermsAndExtractData(
        opportunityId,
        terms,
        calculatorType as 'flux' | 'off-peak' | 'epvs'
      );
      
      if (!updatedData) {
        throw new Error('Failed to update terms and extract data');
          }

          return {
            success: true,
            data: updatedData
          };
    } catch (error) {
      this.logger.error(`Error updating terms: ${error.message}`);
      return {
        success: false,
        error: error.message
      };
    }
  }
}
