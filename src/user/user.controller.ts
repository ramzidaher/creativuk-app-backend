import { Controller, Get, Post, Put, Delete, Body, Param, Query, UseGuards, Req, Logger } from '@nestjs/common';
import { UserService } from './user.service';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { GHLUserLookupService } from '../integrations/ghl-user-lookup.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { RolesGuard } from '../auth/roles.guard';
import { Roles } from '../auth/roles.decorator';
import { UserRole } from '../auth/dto/auth.dto';
import { PrismaService } from '../prisma/prisma.service';

@Controller('user')
export class UserController {
  private readonly logger = new Logger(UserController.name);

  constructor(
    private readonly userService: UserService,
    private readonly goHighLevelService: GoHighLevelService,
    private readonly ghlUserLookupService: GHLUserLookupService,
    private readonly prisma: PrismaService,
  ) {}

  @Get('ghl-sync')
  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  async syncUsersWithGHL(@Req() req: any) {
    this.logger.log('Starting GHL user sync...');
    
    try {
      // Get GHL credentials from environment or config
      const accessToken = process.env.GOHIGHLEVEL_API_TOKEN || process.env.GHL_ACCESS_TOKEN;
      const locationId = process.env.GHL_LOCATION_ID;
      
      if (!accessToken || !locationId) {
        throw new Error('GHL credentials not configured');
      }

      // Get all users from database
      const dbUsers = await this.userService.findAll();
      this.logger.log(`Found ${dbUsers.length} users in database`);

      // Get all users from GHL
      const ghlUsers = await this.goHighLevelService.getAllUsers(accessToken, locationId);
      this.logger.log(`Found ${ghlUsers.length} users in GHL`);

      const results = {
        synced: 0,
        updated: 0,
        errors: 0,
        details: [] as any[]
      };

      // For each database user, try to find matching GHL user
      for (const dbUser of dbUsers) {
        try {
          if (dbUser.ghlUserId) {
            // User already has GHL ID, verify it still exists
            const ghlUser = ghlUsers.find(u => u.id === dbUser.ghlUserId);
            if (ghlUser) {
              results.details.push({
                username: dbUser.username,
                name: dbUser.name,
                ghlUserId: dbUser.ghlUserId,
                status: 'verified'
              });
              results.synced++;
            } else {
              // GHL user ID no longer exists, try to find by name
              const foundUser = await this.goHighLevelService.findUserByName(
                accessToken, 
                locationId, 
                dbUser.name || dbUser.username
              );
              
              if (foundUser) {
                await this.userService.update(dbUser.id, { ghlUserId: foundUser.id });
                results.details.push({
                  username: dbUser.username,
                  name: dbUser.name,
                  oldGhlUserId: dbUser.ghlUserId,
                  newGhlUserId: foundUser.id,
                  status: 'updated'
                });
                results.updated++;
              } else {
                results.details.push({
                  username: dbUser.username,
                  name: dbUser.name,
                  ghlUserId: dbUser.ghlUserId,
                  status: 'not_found_in_ghl'
                });
                results.errors++;
              }
            }
          } else {
            // User doesn't have GHL ID, try to find by name
            const foundUser = await this.goHighLevelService.findUserByName(
              accessToken, 
              locationId, 
              dbUser.name || dbUser.username
            );
            
            if (foundUser) {
              await this.userService.update(dbUser.id, { ghlUserId: foundUser.id });
              results.details.push({
                username: dbUser.username,
                name: dbUser.name,
                newGhlUserId: foundUser.id,
                status: 'assigned'
              });
              results.updated++;
            } else {
              results.details.push({
                username: dbUser.username,
                name: dbUser.name,
                status: 'not_found_in_ghl'
              });
              results.errors++;
            }
          }
        } catch (error) {
          this.logger.error(`Error processing user ${dbUser.username}: ${error.message}`);
          results.details.push({
            username: dbUser.username,
            name: dbUser.name,
            error: error.message,
            status: 'error'
          });
          results.errors++;
        }
      }

      this.logger.log(`GHL sync completed: ${results.synced} verified, ${results.updated} updated, ${results.errors} errors`);
      
      return {
        success: true,
        message: 'GHL user sync completed',
        results
      };
    } catch (error) {
      this.logger.error(`GHL sync failed: ${error.message}`);
      throw error;
    }
  }

  @Get('ghl-status')
  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  async getGHLUserStatus() {
    const users = await this.userService.findAll();
    
    const status = {
      total: users.length,
      withGhlId: users.filter(u => u.ghlUserId).length,
      withoutGhlId: users.filter(u => !u.ghlUserId).length,
      users: users.map(user => ({
        id: user.id,
        username: user.username,
        name: user.name,
        ghlUserId: user.ghlUserId,
        role: user.role,
        hasGhlId: !!user.ghlUserId
      }))
    };

    return status;
  }

  @Post('assign-ghl-id/:userId')
  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  async assignGHLUserId(
    @Param('userId') userId: string,
    @Body() body: { ghlUserId: string }
  ) {
    const user = await this.userService.findById(userId);
    if (!user) {
      throw new Error('User not found');
    }

    const updatedUser = await this.userService.update(userId, { 
      ghlUserId: body.ghlUserId 
    });

    return {
      success: true,
      message: `GHL user ID assigned to ${user.username}`,
      user: {
        id: updatedUser.id,
        username: updatedUser.username,
        name: updatedUser.name,
        ghlUserId: updatedUser.ghlUserId
      }
    };
  }

  @Get('ghl-users')
  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  async getGHLUsers() {
    try {
      const ghlUsers = await this.ghlUserLookupService.getAvailableGHLUsers();
      
      return {
        success: true,
        count: ghlUsers.length,
        users: ghlUsers.map(user => ({
          id: user.id,
          firstName: user.firstName,
          lastName: user.lastName,
          email: user.email,
          fullName: `${user.firstName || ''} ${user.lastName || ''}`.trim()
        }))
      };
    } catch (error) {
      this.logger.error(`Error fetching GHL users: ${error.message}`);
      throw error;
    }
  }

  @Post('assign-ghl-id-manual/:userId')
  @UseGuards(JwtAuthGuard, RolesGuard)
  @Roles(UserRole.ADMIN)
  async assignGHLUserIdManual(
    @Param('userId') userId: string,
    @Body() body: { ghlUserId?: string; ghlUserName?: string }
  ) {
    const user = await this.userService.findById(userId);
    if (!user) {
      throw new Error('User not found');
    }

    // If ghlUserId is provided, assign it directly
    if (body.ghlUserId) {
      const updatedUser = await this.userService.update(userId, { 
        ghlUserId: body.ghlUserId 
      });

      return {
        success: true,
        message: `GHL user ID ${body.ghlUserId} assigned to ${user.username}`,
        user: {
          id: updatedUser.id,
          username: updatedUser.username,
          name: updatedUser.name,
          ghlUserId: updatedUser.ghlUserId
        }
      };
    }

    // If ghlUserName is provided, try to find and assign
    if (body.ghlUserName) {
      const ghlLookupResult = await this.ghlUserLookupService.findGHLUserByName(body.ghlUserName);
      
      if (ghlLookupResult.found && ghlLookupResult.ghlUser) {
        const updatedUser = await this.userService.update(userId, { 
          ghlUserId: ghlLookupResult.ghlUser.id 
        });

        return {
          success: true,
          message: `GHL user found and assigned: ${ghlLookupResult.ghlUser.firstName} ${ghlLookupResult.ghlUser.lastName} (ID: ${ghlLookupResult.ghlUser.id})`,
          user: {
            id: updatedUser.id,
            username: updatedUser.username,
            name: updatedUser.name,
            ghlUserId: updatedUser.ghlUserId
          }
        };
      } else {
        return {
          success: false,
          message: ghlLookupResult.message,
          availableUsers: await this.ghlUserLookupService.getAvailableGHLUsers()
        };
      }
    }

    throw new Error('Either ghlUserId or ghlUserName must be provided');
  }

  @Get('analytics/progress')
  @UseGuards(JwtAuthGuard)
  async getUserAnalytics(
    @Req() req: any,
    @Query('timeRange') timeRange: string = 'all'
  ) {
    try {
      const userId = req.user.id;
      this.logger.log(`Getting analytics for user ${userId} with time range ${timeRange}`);

      // Calculate date range based on timeRange parameter
      const now = new Date();
      let startDate: Date | undefined;
      
      switch (timeRange) {
        case '7d':
          startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
          break;
        case '30d':
          startDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
          break;
        case '90d':
          startDate = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
          break;
        case 'all':
        default:
          startDate = undefined;
          break;
      }

      // Build where clause for date filtering
      const dateFilter = startDate ? { gte: startDate } : undefined;

      // Get user's opportunities
      const opportunities = await this.prisma.opportunityProgress.findMany({
        where: {
          userId,
          ...(dateFilter && { createdAt: dateFilter })
        },
        orderBy: { createdAt: 'desc' }
      });

      // Get survey data
      const surveys = await this.prisma.survey.findMany({
        where: {
          createdBy: userId,
          ...(dateFilter && { createdAt: dateFilter })
        },
        orderBy: { createdAt: 'desc' }
      });

      // Get calculator progress data
      const calculatorProgress = await this.prisma.calculatorProgress.findMany({
        where: {
          userId,
          ...(dateFilter && { createdAt: dateFilter })
        },
        orderBy: { createdAt: 'desc' }
      });

      // Process survey data
      const surveyData = surveys.map(survey => {
        const pages = ['page1', 'page2', 'page3', 'page4', 'page5', 'page6', 'page7', 'page8'];
        const surveyPages: Record<string, boolean> = {};
        let completedPages = 0;

        pages.forEach(page => {
          const hasData = survey[page] !== null && survey[page] !== undefined;
          surveyPages[page] = hasData;
          if (hasData) completedPages++;
        });

        const completionPercentage = Math.round((completedPages / pages.length) * 100);
        const completed = survey.status === 'COMPLETED' || survey.status === 'SUBMITTED' || survey.status === 'APPROVED';

        return {
          opportunityId: survey.ghlOpportunityId,
          lastSavedAt: survey.updatedAt.toISOString(),
          surveyPages,
          completed,
          completionPercentage
        };
      });

      // Process calculator data
      const calculatorData = calculatorProgress.map(calc => {
        const data = calc.data as any;
        const calculatorTypes = {
          'off-peak': data.calculatorType === 'off-peak',
          'flux': data.calculatorType === 'flux',
          'epvs': data.calculatorType === 'epvs'
        };

        // Calculate progress based on completed steps
        const steps = ['template-selection', 'radio-buttons', 'dynamic-inputs', 'arrays', 'pricing', 'completed'];
        const completedSteps = data.completedSteps || {};
        const completedCount = Object.values(completedSteps).filter(Boolean).length;
        const progressPercentage = Math.round((completedCount / steps.length) * 100);
        
        const completed = data.currentStep === 'completed' || progressPercentage === 100;

        return {
          opportunityId: calc.opportunityId,
          lastSavedAt: calc.updatedAt.toISOString(),
          calculatorTypes,
          currentStep: data.currentStep || 'template-selection',
          completed,
          progressPercentage
        };
      });

      // Calculate statistics
      const totalSurveys = surveys.length;
      const completedSurveys = surveys.filter(s => 
        s.status === 'COMPLETED' || s.status === 'SUBMITTED' || s.status === 'APPROVED'
      ).length;

      const totalCalculators = calculatorProgress.length;
      const completedCalculators = calculatorProgress.filter(c => {
        const data = c.data as any;
        return data.currentStep === 'completed' || 
               (data.completedSteps && Object.values(data.completedSteps).filter(Boolean).length === 6);
      }).length;

      // Calculate average completion time (in minutes)
      const completedSurveyTimes = surveys
        .filter(s => s.submittedAt && s.createdAt)
        .map(s => s.submittedAt!.getTime() - s.createdAt.getTime())
        .map(time => time / (1000 * 60)); // Convert to minutes

      const completedCalculatorTimes = calculatorProgress
        .filter(c => {
          const data = c.data as any;
          return data.currentStep === 'completed';
        })
        .map(c => c.updatedAt.getTime() - c.createdAt.getTime())
        .map(time => time / (1000 * 60)); // Convert to minutes

      const allCompletionTimes = [...completedSurveyTimes, ...completedCalculatorTimes];
      const averageCompletionTime = allCompletionTimes.length > 0 
        ? Math.round(allCompletionTimes.reduce((a, b) => a + b, 0) / allCompletionTimes.length)
        : 0;

      // Find last activity
      const allActivities = [
        ...surveys.map(s => s.updatedAt),
        ...calculatorProgress.map(c => c.updatedAt),
        ...opportunities.map(o => o.lastActivityAt)
      ];
      const lastActivity = allActivities.length > 0 
        ? new Date(Math.max(...allActivities.map(d => d.getTime()))).toISOString()
        : null;

      return {
        success: true,
        data: {
          opportunities: opportunities.map(o => ({
            id: o.ghlOpportunityId,
            status: o.status,
            currentStep: o.currentStep,
            startedAt: o.startedAt.toISOString(),
            lastActivityAt: o.lastActivityAt.toISOString()
          })),
          surveyData,
          calculatorData,
          totalSurveys,
          totalCalculators,
          completedSurveys,
          completedCalculators,
          lastActivity,
          averageCompletionTime
        }
      };
    } catch (error) {
      this.logger.error(`Error getting user analytics: ${error.message}`);
      throw error;
    }
  }
}
