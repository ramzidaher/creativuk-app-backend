import { Controller, Get, Param, UseGuards } from '@nestjs/common';
import { AdminOpportunityDetailsService } from './admin-opportunity-details.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AdminGuard } from '../auth/admin.guard';

@Controller('admin/opportunities')
@UseGuards(JwtAuthGuard, AdminGuard)
export class AdminOpportunityDetailsController {
  constructor(
    private readonly adminOpportunityDetailsService: AdminOpportunityDetailsService
  ) {}

  /**
   * Get all users with their opportunities (summary)
   */
  @Get('users')
  async getAllUsersWithOpportunities() {
    return this.adminOpportunityDetailsService.getAllUsersWithOpportunitiesSummary();
  }

  /**
   * Get all users with their opportunities (full details)
   */
  @Get('users/full')
  async getAllUsersWithOpportunitiesFull() {
    return this.adminOpportunityDetailsService.getAllUsersWithOpportunities();
  }

  /**
   * Get complete opportunity details including all related data
   */
  @Get('details/:opportunityId')
  async getOpportunityDetails(@Param('opportunityId') opportunityId: string) {
    return this.adminOpportunityDetailsService.getOpportunityDetails(opportunityId);
  }
}


