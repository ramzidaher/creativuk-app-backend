import { Controller, Get, Query, UseGuards, Request, Param } from '@nestjs/common';
import { AdminAnalyticsService } from './admin-analytics.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AdminGuard } from '../auth/admin.guard';

@Controller('admin/analytics')
@UseGuards(JwtAuthGuard, AdminGuard)
export class AdminAnalyticsController {
  constructor(private readonly adminAnalyticsService: AdminAnalyticsService) {}

  /**
   * Get comprehensive system analytics
   */
  @Get('system')
  async getSystemAnalytics(@Request() req) {
    console.log(`📊 Admin ${req.user.id} requested system analytics`);
    return this.adminAnalyticsService.getSystemAnalytics();
  }

  /**
   * Get system logs
   */
  @Get('logs')
  async getSystemLogs(
    @Query('filter') filter?: string,
    @Query('limit') limit?: string
  ) {
    const limitNumber = limit ? parseInt(limit, 10) : 100;
    return this.adminAnalyticsService.getSystemLogs(filter, limitNumber);
  }

  /**
   * Get user activity summary
   */
  @Get('user-activity')
  async getUserActivitySummary(@Query('userId') userId?: string) {
    return this.adminAnalyticsService.getUserActivitySummary(userId);
  }

  /**
   * Get system performance metrics
   */
  @Get('performance')
  async getSystemPerformanceMetrics() {
    return this.adminAnalyticsService.getSystemPerformanceMetrics();
  }

  /**
   * Get all users
   */
  @Get('users')
  async getAllUsers() {
    return this.adminAnalyticsService.getAllUsers();
  }

  /**
   * Get user opportunities summary (lightweight version)
   */
  @Get('users/:userId/opportunities/summary')
  async getUserOpportunitiesSummary(@Param('userId') userId: string) {
    return this.adminAnalyticsService.getUserOpportunitiesSummary(userId);
  }

  /**
   * Get user opportunities by user ID with pagination
   */
  @Get('users/:userId/opportunities')
  async getUserOpportunities(
    @Param('userId') userId: string,
    @Query('page') page?: string,
    @Query('limit') limit?: string
  ) {
    const pageNumber = page ? parseInt(page, 10) : 1;
    const limitNumber = limit ? parseInt(limit, 10) : 50;
    
    // Validate pagination parameters
    if (pageNumber < 1) {
      throw new Error('Page number must be greater than 0');
    }
    if (limitNumber < 1 || limitNumber > 100) {
      throw new Error('Limit must be between 1 and 100');
    }
    
    return this.adminAnalyticsService.getUserOpportunities(userId, pageNumber, limitNumber);
  }

  /**
   * Get user autosaved survey data
   */
  @Get('users/:userId/survey-data')
  async getUserAutosavedSurveyData(
    @Param('userId') userId: string,
    @Query('opportunityId') opportunityId?: string
  ) {
    return this.adminAnalyticsService.getUserAutosavedSurveyData(userId, opportunityId);
  }

  /**
   * Get user autosaved calculator data
   */
  @Get('users/:userId/calculator-data')
  async getUserAutosavedCalculatorData(
    @Param('userId') userId: string,
    @Query('opportunityId') opportunityId?: string
  ) {
    return this.adminAnalyticsService.getUserAutosavedCalculatorData(userId, opportunityId);
  }

  /**
   * Get comprehensive user data (opportunities + autosaved data) with pagination
   */
  @Get('users/:userId/comprehensive')
  async getUserComprehensiveData(
    @Param('userId') userId: string,
    @Query('page') page?: string,
    @Query('limit') limit?: string
  ) {
    const pageNumber = page ? parseInt(page, 10) : 1;
    const limitNumber = limit ? parseInt(limit, 10) : 25;
    
    // Validate pagination parameters
    if (pageNumber < 1) {
      throw new Error('Page number must be greater than 0');
    }
    if (limitNumber < 1 || limitNumber > 50) {
      throw new Error('Limit must be between 1 and 50 for comprehensive data');
    }
    
    return this.adminAnalyticsService.getUserComprehensiveData(userId, pageNumber, limitNumber);
  }
}

