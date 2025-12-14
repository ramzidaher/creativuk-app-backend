import { Controller, Post, Body, Logger } from '@nestjs/common';
import { ExcelCellDetectorService } from './excel-cell-detector.service';

interface GetEnabledInputsRequest {
  opportunityId?: string;
}

@Controller('excel-cell-detector')
export class ExcelCellDetectorController {
  private readonly logger = new Logger(ExcelCellDetectorController.name);

  constructor(private readonly excelCellDetectorService: ExcelCellDetectorService) {}

  @Post('get-enabled-inputs')
  async getEnabledInputs(@Body() request: GetEnabledInputsRequest) {
    this.logger.log(`Getting enabled inputs for opportunity: ${request.opportunityId || 'template'}`);
    
    try {
      const result = await this.excelCellDetectorService.getEnabledInputFields(request.opportunityId);
      
      this.logger.log(`Successfully retrieved ${result.inputFields?.length || 0} enabled input fields`);
      
      return result;
    } catch (error) {
      this.logger.error(`Error in getEnabledInputs: ${error.message}`);
      return {
        success: false,
        message: 'Failed to get enabled input fields',
        error: error.message
      };
    }
  }
}
