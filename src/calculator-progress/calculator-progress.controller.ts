import { Controller, Post, Body, Get, Query, Delete, Put, Logger } from '@nestjs/common';
import { CalculatorProgressService } from './calculator-progress.service';

interface SaveProgressDto {
  userId: string;
  opportunityId: string;
  calculatorType: 'off-peak' | 'flux' | 'epvs';
  progressData: any;
}

interface GetProgressDto {
  userId: string;
  opportunityId: string;
  calculatorType: 'off-peak' | 'flux' | 'epvs';
  existingFileName?: string;
}

interface CheckChangesDto {
  userId: string;
  opportunityId: string;
  calculatorType: 'off-peak' | 'flux' | 'epvs';
  newData: any;
}

interface ClearProgressDto {
  userId: string;
  opportunityId: string;
  calculatorType?: 'off-peak' | 'flux' | 'epvs';
}

@Controller('calculator-progress')
export class CalculatorProgressController {
  private readonly logger = new Logger(CalculatorProgressController.name);

  constructor(private readonly calculatorProgressService: CalculatorProgressService) {}

  @Post('save')
  async saveProgress(@Body() saveProgressDto: SaveProgressDto) {
    try {
      const { userId, opportunityId, calculatorType, progressData } = saveProgressDto;
      
      this.logger.log(`Saving progress for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
      
      const result = await this.calculatorProgressService.saveProgress(
        userId,
        opportunityId,
        calculatorType,
        progressData
      );
      
      return result;
    } catch (error) {
      this.logger.error(`Error saving progress: ${error.message}`);
      return {
        success: false,
        message: `Error saving progress: ${error.message}`,
      };
    }
  }

  @Get('get')
  async getProgress(@Query() getProgressDto: GetProgressDto) {
    try {
      const { userId, opportunityId, calculatorType } = getProgressDto;
      
      this.logger.log(`Getting progress for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
      
      const progress = await this.calculatorProgressService.getProgress(
        userId,
        opportunityId,
        calculatorType
      );
      
      return {
        success: true,
        data: progress,
        message: progress ? 'Progress found' : 'No progress found',
      };
    } catch (error) {
      this.logger.error(`Error getting progress: ${error.message}`);
      return {
        success: false,
        message: `Error getting progress: ${error.message}`,
        data: null,
      };
    }
  }

  @Post('check-changes')
  async checkChanges(@Body() checkChangesDto: CheckChangesDto) {
    try {
      const { userId, opportunityId, calculatorType, newData } = checkChangesDto;
      
      this.logger.log(`Checking changes for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
      
      const result = await this.calculatorProgressService.hasDataChanged(
        userId,
        opportunityId,
        calculatorType,
        newData
      );
      
      return {
        success: true,
        ...result,
        message: result.hasChanged ? 'Data has changed' : 'No changes detected',
      };
    } catch (error) {
      this.logger.error(`Error checking changes: ${error.message}`);
      return {
        success: false,
        message: `Error checking changes: ${error.message}`,
        hasChanged: true,
      };
    }
  }

  @Get('summary')
  async getProgressSummary(@Query() getProgressDto: GetProgressDto) {
    try {
      const { userId, opportunityId, calculatorType } = getProgressDto;
      
      this.logger.log(`Getting progress summary for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
      
      const summary = await this.calculatorProgressService.getProgressSummary(
        userId,
        opportunityId,
        calculatorType
      );
      
      return {
        success: true,
        data: summary,
        message: 'Progress summary retrieved',
      };
    } catch (error) {
      this.logger.error(`Error getting progress summary: ${error.message}`);
      return {
        success: false,
        message: `Error getting progress summary: ${error.message}`,
        data: null,
      };
    }
  }

  @Delete('clear')
  async clearProgress(@Body() clearProgressDto: ClearProgressDto) {
    try {
      const { userId, opportunityId, calculatorType } = clearProgressDto;
      
      this.logger.log(`Clearing progress for user ${userId}, opportunity ${opportunityId}, type ${calculatorType || 'all'}`);
      
      const result = await this.calculatorProgressService.clearProgress(
        userId,
        opportunityId,
        calculatorType
      );
      
      return result;
    } catch (error) {
      this.logger.error(`Error clearing progress: ${error.message}`);
      return {
        success: false,
        message: `Error clearing progress: ${error.message}`,
      };
    }
  }

  @Post('submit')
  async submitCalculator(@Body() submitDto: GetProgressDto) {
    try {
      this.logger.log(`🚀 SUBMIT ENDPOINT CALLED - Received request: ${JSON.stringify(submitDto)}`);
      
      const { userId, opportunityId, calculatorType, existingFileName } = submitDto;
      
      if (!userId || !opportunityId || !calculatorType) {
        this.logger.error(`❌ Missing required fields: userId=${userId}, opportunityId=${opportunityId}, calculatorType=${calculatorType}`);
        return {
          success: false,
          message: 'Missing required fields: userId, opportunityId, and calculatorType are required',
          error: 'Missing required fields',
        };
      }
      
      this.logger.log(`📤 Submitting calculator for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}${existingFileName ? ` (editing existing file: ${existingFileName})` : ''}`);
      
      const result = await this.calculatorProgressService.submitCalculator(
        userId,
        opportunityId,
        calculatorType,
        existingFileName
      );
      
      this.logger.log(`✅ Submit result: ${JSON.stringify(result)}`);
      
      return {
        success: result.success,
        message: result.message,
        filePath: result.filePath,
        error: result.error,
      };
    } catch (error) {
      this.logger.error(`❌ Error submitting calculator: ${error.message}`);
      this.logger.error(`❌ Error stack: ${error.stack}`);
      return {
        success: false,
        message: `Error submitting calculator: ${error.message}`,
        error: error.message,
      };
    }
  }

  @Get('submit')
  async getSubmissionStatus(@Query() getProgressDto: GetProgressDto) {
    try {
      const { userId, opportunityId, calculatorType } = getProgressDto;
      
      this.logger.log(`Checking submission status for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
      
      const result = await this.calculatorProgressService.getSubmissionStatus(
        userId,
        opportunityId,
        calculatorType
      );
      
      return {
        success: result.success,
        submitted: result.submitted,
        filePath: result.filePath,
        message: result.message,
        error: result.error,
      };
    } catch (error) {
      this.logger.error(`Error checking submission status: ${error.message}`);
      return {
        success: false,
        submitted: false,
        message: `Error checking submission status: ${error.message}`,
        error: error.message,
      };
    }
  }

  @Put('submit')
  async retrySubmission(@Body() submitDto: GetProgressDto) {
    try {
      const { userId, opportunityId, calculatorType } = submitDto;
      
      this.logger.log(`Retrying calculator submission for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
      
      // PUT is idempotent - retries the submission
      const result = await this.calculatorProgressService.submitCalculator(
        userId,
        opportunityId,
        calculatorType
      );
      
      return {
        success: result.success,
        message: result.message,
        filePath: result.filePath,
        error: result.error,
      };
    } catch (error) {
      this.logger.error(`Error retrying submission: ${error.message}`);
      return {
        success: false,
        message: `Error retrying submission: ${error.message}`,
        error: error.message,
      };
    }
  }
}

