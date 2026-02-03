import { Controller, Get, Post, Put, Param, Query, UseGuards, Request, Req, Body } from '@nestjs/common';
import { OpportunitiesService } from './opportunities.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { RolesGuard } from '../auth/roles.guard';
import { Roles } from '../auth/roles.decorator';
import { UserRole } from '../auth/dto/auth.dto';
import { UnauthorizedException, NotFoundException } from '@nestjs/common';
import { UserService } from '../user/user.service';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { Logger } from '@nestjs/common';
import { DynamicSurveyorService } from './dynamic-surveyor.service';

@Controller('opportunities')
@UseGuards(JwtAuthGuard)
export class OpportunitiesController {
  private readonly logger = new Logger(OpportunitiesController.name);

  constructor(
    private readonly opportunitiesService: OpportunitiesService,
    private readonly userService: UserService,
    private readonly goHighLevelService: GoHighLevelService,
    private readonly dynamicSurveyorService: DynamicSurveyorService
  ) {}

  @Get()
  async getOpportunities(@Request() req) {
    // Get opportunities based on user role
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    if (userRole === UserRole.ADMIN) {
      // Admins can see all opportunities
      return this.opportunitiesService.getOpportunities(userId);
    } else {
      // Surveyors can only see their assigned opportunities
      return this.opportunitiesService.getOpportunities(userId);
    }
  }

  @Get('ai-vs-manual')
  async getAiVsManualOpportunities(@Request() req) {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    if (userRole === UserRole.ADMIN) {
      return this.opportunitiesService.getAiVsManualOpportunities(userId);
    } else {
      return this.opportunitiesService.getAiVsManualOpportunities(userId);
    }
  }

  @Get('assigned-to-james')
  async getOpportunitiesAssignedToJames(@Request() req) {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    if (userRole === UserRole.ADMIN) {
      return this.opportunitiesService.getOpportunitiesByAssignedTeamMember(userId, 'James Barnett');
    } else {
      return this.opportunitiesService.getOpportunitiesByAssignedTeamMember(userId, 'James Barnett');
    }
  }

  @Get('all')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getAllOpportunities(@Request() req) {
    return this.opportunitiesService.getAllOpportunities(req.user.sub);
  }

  @Get('stages')
  async getOpportunityStages(@Request() req) {
    return this.opportunitiesService.getOpportunityStages(req.user.sub);
  }

  @Get('simple/list')
  async getSimpleOpportunities(@Request() req) {
    return this.opportunitiesService.getSimpleOpportunities(req.user.sub);
  }

  @Get('with-appointments')
  async getOpportunitiesWithAppointments(@Request() req) {
    return this.opportunitiesService.getOpportunitiesWithAppointments(req.user.sub);
  }

  @Get('with-appointments-optimized')
  async getOpportunitiesWithAppointmentsOptimized(@Request() req) {
    return this.opportunitiesService.getOpportunitiesWithAppointmentsOptimized(req.user.sub);
  }

  @Get('with-appointments-hybrid')
  async getOpportunitiesWithAppointmentsHybrid(@Request() req) {
    return this.opportunitiesService.getOpportunitiesWithAppointmentsHybrid(req.user.sub);
  }

  @Get('with-appointments-unified')
  async getOpportunitiesWithAppointmentsUnified(@Request() req) {
    return this.opportunitiesService.getOpportunitiesWithAppointmentsUnified(req.user.sub);
  }

  @Get(':opportunityId/details')
  async getOpportunityDetails(@Param('opportunityId') opportunityId: string, @Request() req) {
    this.logger.log(`🔍 Request for opportunity details: ${opportunityId}`);
    return this.opportunitiesService.getOpportunityDetails(opportunityId, req.user.sub);
  }

  @Get(':opportunityId/customer-details')
  async getCustomerDetails(@Param('opportunityId') opportunityId: string, @Request() req) {
    this.logger.log(`🔍 Request for customer details: ${opportunityId}`);
    return this.opportunitiesService.getCustomerDetails(opportunityId, req.user.sub);
  }

  @Get('performance-comparison')
  async performanceComparison(@Request() req) {
    const startTime = Date.now();
    
    try {
      // Test original method
      const originalStart = Date.now();
      const originalResult = await this.opportunitiesService.getOpportunitiesWithAppointments(req.user.sub);
      const originalTime = Date.now() - originalStart;
      
      // Test optimized method
      const optimizedStart = Date.now();
      const optimizedResult = await this.opportunitiesService.getOpportunitiesWithAppointmentsOptimized(req.user.sub);
      const optimizedTime = Date.now() - optimizedStart;
      
      const totalTime = Date.now() - startTime;
      
      return {
        comparison: {
          original: {
            time: originalTime,
            opportunities: originalResult.total,
            classification: originalResult.classification
          },
          optimized: {
            time: optimizedTime,
            opportunities: optimizedResult.total,
            classification: optimizedResult.classification,
            performance: optimizedResult.performance
          },
          improvement: {
            timeSaved: originalTime - optimizedTime,
            percentageFaster: Math.round(((originalTime - optimizedTime) / originalTime) * 100)
          }
        },
        totalTime,
        recommendation: optimizedTime < originalTime ? 'Use optimized method' : 'Use original method'
      };
    } catch (error) {
      this.logger.error(`Error in performance comparison: ${error.message}`);
      throw error;
    }
  }

  @Get('test-all-opportunities')
  async testAllOpportunities(@Request() req) {
    const userId = req.user.sub; // Use req.user.sub instead of req.user?.id
    if (!userId) {
      throw new UnauthorizedException('User not authenticated');
    }

    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      const credentials = this.opportunitiesService['getGhlCredentials']();
      if (!credentials) {
        throw new NotFoundException('GHL credentials not configured');
      }

      // Get opportunities from AI and Manual stages
      const aiStageId = '77740d71-fd7e-47df-a9de-a7f1e4db0b87'; // (AI Bot) Home Survey Booked
      const manualStageId = '08f2f487-14c5-44ef-b2f7-0021c605efb2'; // (Manual) Home Survey Booked
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      // Filter opportunities based on user role
      const filteredOpportunities = await this.opportunitiesService['filterOpportunitiesByUserRoleAndAppointments'](
        allOpportunities,
        user,
        credentials
      );

      // Separate AI and Manual opportunities
      const aiOpportunities = filteredOpportunities.filter(opp => opp.pipelineStageId === aiStageId);
      const manualOpportunities = filteredOpportunities.filter(opp => opp.pipelineStageId === manualStageId);

      return {
        summary: {
          total: filteredOpportunities.length,
          ai: aiOpportunities.length,
          manual: manualOpportunities.length,
          user: {
            id: user.id,
            name: user.name,
            role: user.role
          }
        },
        ai: {
          stageName: '(AI Bot) Home Survey Booked',
          stageId: aiStageId,
          opportunities: aiOpportunities
        },
        manual: {
          stageName: '(Manual) Home Survey Booked',
          stageId: manualStageId,
          opportunities: manualOpportunities
        },
        all: filteredOpportunities
      };
    } catch (error) {
      this.logger.error(`Error in test all opportunities: ${error.message}`);
      throw error;
    }
  }

  @Get('test-appointment-analysis')
  async testAppointmentAnalysis(@Request() req) {
    const userId = req.user.sub;
    if (!userId) {
      throw new UnauthorizedException('User not authenticated');
    }

    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      const credentials = this.opportunitiesService['getGhlCredentials']();
      if (!credentials) {
        throw new NotFoundException('GHL credentials not configured');
      }

      // Get the detailed analysis
      const result = await this.opportunitiesService.getOpportunitiesWithAppointments(userId);

      return {
        success: true,
        data: {
          ...result,
          analysis: {
            totalOpportunities: result.total,
            confirmedWithAppointments: result.classification?.confirmedWithAppointments || 0,
            multipleAppointments: result.classification?.multipleAppointments || 0,
            noAppointments: result.classification?.noAppointments || 0,
            accuracy: {
              confirmedPercentage: Math.round(((result.classification?.confirmedWithAppointments || 0) / result.total) * 100),
              multiplePercentage: Math.round(((result.classification?.multipleAppointments || 0) / result.total) * 100),
              noAppointmentPercentage: Math.round(((result.classification?.noAppointments || 0) / result.total) * 100),
            }
          },
          user: {
            id: user.id,
            name: user.name,
            role: user.role,
          },
          timestamp: new Date().toISOString(),
        }
      };
    } catch (error) {
      this.logger.error(`Error in test appointment analysis: ${error.message}`);
      throw error;
    }
  }

  @Get('debug-appointments/:opportunityId')
  async debugAppointments(@Request() req, @Param('opportunityId') opportunityId: string) {
    const userId = req.user.sub; // Use req.user.sub instead of req.user?.id
    if (!userId) {
      throw new UnauthorizedException('User not authenticated');
    }

    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      const credentials = this.opportunitiesService['getGhlCredentials']();
      if (!credentials) {
        throw new NotFoundException('GHL credentials not configured');
      }

      // Get the specific opportunity
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        ['77740d71-fd7e-47df-a9de-a7f1e4db0b87', '08f2f487-14c5-44ef-b2f7-0021c605efb2']
      );

      const opportunity = allOpportunities.find(opp => opp.id === opportunityId);
      if (!opportunity) {
        throw new NotFoundException('Opportunity not found');
      }

      const contactId = opportunity.contactId || opportunity.contact?.id;
      
      let appointments: any[] = [];
      if (contactId) {
        // Try direct method
        appointments = await this.goHighLevelService.getAppointmentsByContactId(
          credentials.accessToken,
          credentials.locationId,
          contactId
        );

        // If no appointments, try alternative
        if (appointments.length === 0) {
          appointments = await this.goHighLevelService.getAppointmentsByContactIdAlternative(
            credentials.accessToken,
            credentials.locationId,
            contactId
          );
        }
      }

      return {
        opportunity: {
          id: opportunity.id,
          name: opportunity.name,
          contactId: contactId,
          contact: opportunity.contact
        },
        appointments: appointments,
        appointmentCount: appointments.length,
        hasAppointment: appointments.length > 0
      };
    } catch (error) {
      this.logger.error(`Error in debug appointments: ${error.message}`);
      throw error;
    }
  }

  @Get('contact-appointments/:contactId')
  async getContactAppointments(@Request() req, @Param('contactId') contactId: string) {
    const userId = req.user.sub;
    if (!userId) {
      throw new UnauthorizedException('User not authenticated');
    }

    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      const credentials = this.opportunitiesService['getGhlCredentials']();
      if (!credentials) {
        throw new NotFoundException('GHL credentials not configured');
      }

      // Get appointments for the specific contact
      const appointments = await this.goHighLevelService.getAppointmentsByContactId(
        credentials.accessToken,
        credentials.locationId,
        contactId
      );

      // If no appointments found, try alternative method
      let allAppointments = appointments;
      if (appointments.length === 0) {
        const alternativeAppointments = await this.goHighLevelService.getAppointmentsByContactIdAlternative(
          credentials.accessToken,
          credentials.locationId,
          contactId
        );
        allAppointments = alternativeAppointments;
      }

      return {
        contactId: contactId,
        appointments: allAppointments,
        appointmentCount: allAppointments.length,
        hasAppointments: allAppointments.length > 0,
        user: {
          id: user.id,
          name: user.name,
          role: user.role
        }
      };
    } catch (error) {
      this.logger.error(`Error getting contact appointments: ${error.message}`);
      throw error;
    }
  }

  @Get('by-stage/:stageName')
  async getOpportunitiesByStage(@Request() req, @Param('stageName') stageName: string) {
    return this.opportunitiesService.getOpportunitiesByStage(req.user.sub, stageName);
  }

  @Get('pipeline-stages/all')
  async getAllPipelineStagesWithIds(@Request() req) {
    return this.opportunitiesService.getAllPipelineStagesWithIds(req.user.sub);
  }

  @Get('sales-performance')
  async getSalesPerformanceStats(@Request() req, @Query('month') month?: string, @Query('year') year?: string) {
    return this.opportunitiesService.getSalesPerformanceStats(req.user.sub, month, year);
  }

  @Get('debug-all-won')
  async getAllWonOpportunities(@Request() req) {
    return this.opportunitiesService.getAllWonOpportunities(req.user.sub);
  }

  @Get('opportunities-with-won')
  async getOpportunitiesWithWon(@Request() req) {
    return this.opportunitiesService.getOpportunitiesWithWon(req.user.sub);
  }

  @Get('debug-all-user-opportunities')
  async debugAllUserOpportunities(@Request() req) {
    return this.opportunitiesService.debugAllUserOpportunities(req.user.sub);
  }

  @Get('pipelines')
  async getPipelines(@Request() req) {
    return this.opportunitiesService.getPipelines(req.user.sub);
  }

  @Get('pipelines/:pipelineId/opportunities')
  async getOpportunitiesByPipeline(@Request() req, @Param('pipelineId') pipelineId: string) {
    return this.opportunitiesService.getOpportunitiesByPipeline(req.user.sub, pipelineId);
  }

  @Get('stage-progression/:stageName')
  async getOpportunitiesByStageProgression(@Request() req, @Param('stageName') stageName: string) {
    return this.opportunitiesService.getOpportunitiesByStageProgression(req.user.sub, stageName);
  }

  @Get('stage-progression/:stageName/unfiltered')
  async getOpportunitiesByStageProgressionUnfiltered(@Request() req, @Param('stageName') stageName: string) {
    return this.opportunitiesService.getOpportunitiesByStageProgressionUnfiltered(req.user.sub, stageName);
  }

  @Get(':id')
  async getOpportunityById(@Request() req, @Param('id') id: string) {
    return this.opportunitiesService.getOpportunityById(req.user.sub, id);
  }

  @Get(':id/notes')
  async getOpportunityNotes(@Request() req, @Param('id') id: string) {
    return this.opportunitiesService.getOpportunityNotes(req.user.sub, id);
  }

  @Put(':id/status')
  async updateOpportunityStatus(
    @Request() req,
    @Param('id') opportunityId: string,
    @Body() body: { status: 'open' | 'won' | 'lost' | 'abandoned'; stageId?: string }
  ) {
    return this.opportunitiesService.updateOpportunityStatus(
      opportunityId,
      body.status,
      body.stageId
    );
  }

  @Post(':id/move-to-signed-contract')
  async moveOpportunityToSignedContract(
    @Request() req,
    @Param('id') opportunityId: string
  ) {
    return this.opportunitiesService.moveOpportunityToSignedContract(opportunityId);
  }

  // Surveyor management endpoints
  @Get('surveyors')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getAllSurveyors() {
    return this.dynamicSurveyorService.getAllSurveyors();
  }

  @Get('surveyors/needing-configuration')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getSurveyorsNeedingConfiguration() {
    return this.dynamicSurveyorService.getSurveyorsNeedingConfiguration();
  }

  @Get('surveyors/:name')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getSurveyorByName(@Param('name') name: string) {
    return this.dynamicSurveyorService.getSurveyorByName(name);
  }

  /**
   * Sync opportunities from GHL to database (Admin only)
   */
  @Post('sync-from-ghl')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async syncOpportunitiesFromGHL(
    @Request() req,
    @Query('limit') limit?: string,
  ) {
    const limitNum = limit ? parseInt(limit, 10) : undefined;
    return this.opportunitiesService.syncOpportunitiesFromGHL(req.user.sub, limitNum);
  }

  /**
   * Get opportunities from database (Admin only)
   */
  @Get('db/all')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async getOpportunitiesFromDB(
    @Request() req,
    @Query('userId') userId?: string,
    @Query('outcome') outcome?: string,
    @Query('status') status?: string,
    @Query('startDate') startDate?: string,
    @Query('endDate') endDate?: string,
    @Query('limit') limit?: string,
  ) {
    const filters: any = {};
    if (userId) filters.userId = userId;
    if (outcome) filters.outcome = outcome;
    if (status) filters.status = status;
    if (startDate) filters.startDate = new Date(startDate);
    if (endDate) filters.endDate = new Date(endDate);
    if (limit) filters.limit = parseInt(limit, 10);

    return this.opportunitiesService.getOpportunitiesFromDB(filters);
  }
} 