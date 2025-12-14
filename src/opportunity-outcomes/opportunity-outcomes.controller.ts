import {
  Controller,
  Get,
  Post,
  Put,
  Delete,
  Body,
  Param,
  Query,
  Request,
  UseGuards,
  ParseUUIDPipe,
} from '@nestjs/common';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { OpportunityOutcomesService, OpportunityOutcomeData } from './opportunity-outcomes.service';

@Controller('opportunity-outcomes')
@UseGuards(JwtAuthGuard)
export class OpportunityOutcomesController {
  constructor(private readonly opportunityOutcomesService: OpportunityOutcomesService) {}

  /**
   * Record an opportunity outcome
   */
  @Post()
  async recordOutcome(@Request() req, @Body() data: OpportunityOutcomeData) {
    return this.opportunityOutcomesService.recordOutcome({
      ...data,
      userId: req.user.sub,
    });
  }

  /**
   * Get win/loss statistics for the current user
   */
  @Get('stats')
  async getUserStats(
    @Request() req,
    @Query('startDate') startDate?: string,
    @Query('endDate') endDate?: string,
  ) {
    const start = startDate ? new Date(startDate) : undefined;
    const end = endDate ? new Date(endDate) : undefined;
    
    return this.opportunityOutcomesService.getUserWinLossStats(req.user.sub, start, end);
  }

  /**
   * Get recent outcomes for the current user
   */
  @Get('recent')
  async getRecentOutcomes(
    @Request() req,
    @Query('limit') limit?: string,
  ) {
    const limitNum = limit ? parseInt(limit, 10) : 10;
    return this.opportunityOutcomesService.getRecentOutcomes(req.user.sub, limitNum);
  }

  /**
   * Get win/loss statistics for all users (admin only)
   */
  @Get('admin/all-stats')
  async getAllUsersStats(
    @Request() req,
    @Query('startDate') startDate?: string,
    @Query('endDate') endDate?: string,
  ) {
    // TODO: Add admin role check
    const start = startDate ? new Date(startDate) : undefined;
    const end = endDate ? new Date(endDate) : undefined;
    
    return this.opportunityOutcomesService.getAllUsersWinLossStats(start, end);
  }

  /**
   * Get overall company statistics (admin only)
   */
  @Get('admin/overall-stats')
  async getOverallStats(
    @Request() req,
    @Query('startDate') startDate?: string,
    @Query('endDate') endDate?: string,
  ) {
    // TODO: Add admin role check
    const start = startDate ? new Date(startDate) : undefined;
    const end = endDate ? new Date(endDate) : undefined;
    
    return this.opportunityOutcomesService.getOverallStats(start, end);
  }

  /**
   * Sync opportunity outcomes from GHL API
   */
  @Post('sync-ghl')
  async syncFromGHL(@Request() req) {
    return this.opportunityOutcomesService.syncFromGHL(req.user.sub);
  }

  /**
   * Update an existing outcome
   */
  @Put(':id')
  async updateOutcome(
    @Request() req,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() data: Partial<OpportunityOutcomeData>,
  ) {
    // TODO: Add validation to ensure user can only update their own outcomes
    return this.opportunityOutcomesService.recordOutcome({
      ...data,
      userId: req.user.sub,
    } as OpportunityOutcomeData);
  }

  /**
   * Get outcome by GHL opportunity ID
   */
  @Get('opportunity/:opportunityId')
  async getOutcomeByOpportunityId(@Request() req, @Param('opportunityId') opportunityId: string) {
    return this.opportunityOutcomesService.getOutcomeByOpportunityId(opportunityId);
  }

  /**
   * Get specific outcome by ID
   */
  @Get(':id')
  async getOutcome(@Request() req, @Param('id', ParseUUIDPipe) id: string) {
    // TODO: Implement get specific outcome
    throw new Error('Not implemented yet');
  }

  /**
   * Delete an outcome (admin only)
   */
  @Delete(':id')
  async deleteOutcome(@Request() req, @Param('id', ParseUUIDPipe) id: string) {
    // TODO: Implement delete outcome with admin check
    throw new Error('Not implemented yet');
  }
}

