import { Injectable, Logger, HttpException, HttpStatus } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { UserService } from '../user/user.service';
import { OpportunitiesService } from '../opportunities/opportunities.service';
import { AutoSaveService } from '../opportunities/auto-save.service';

@Injectable()
export class AdminAnalyticsService {
  private readonly logger = new Logger(AdminAnalyticsService.name);
  private readonly MAX_JSON_SIZE = 50 * 1024 * 1024; // 50MB limit

  constructor(
    private readonly prisma: PrismaService,
    private readonly userService: UserService,
    private readonly opportunitiesService: OpportunitiesService,
    private readonly autoSaveService: AutoSaveService
  ) {}

  /**
   * Safely serialize data to JSON with size checking
   */
  private safeJsonSerialize(data: any, context: string): string {
    try {
      const jsonString = JSON.stringify(data);
      
      if (jsonString.length > this.MAX_JSON_SIZE) {
        this.logger.error(`Data too large for JSON serialization in ${context}: ${jsonString.length} bytes (max: ${this.MAX_JSON_SIZE})`);
        throw new HttpException(
          `Response data too large (${Math.round(jsonString.length / 1024 / 1024)}MB). Please use pagination to reduce data size.`,
          HttpStatus.PAYLOAD_TOO_LARGE
        );
      }
      
      return jsonString;
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }
      
      this.logger.error(`JSON serialization error in ${context}:`, error);
      throw new HttpException(
        'Data serialization failed. The response data may be too large or contain circular references.',
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Get comprehensive system analytics
   */
  async getSystemAnalytics() {
    try {
      this.logger.log('Fetching system analytics...');

      // Get user statistics
      const totalUsers = await this.prisma.user.count();
      const activeUsers = await this.prisma.user.count({
        where: { status: 'ACTIVE' }
      });
      const adminUsers = await this.prisma.user.count({
        where: { role: 'ADMIN' }
      });
      const surveyorUsers = await this.prisma.user.count({
        where: { role: 'SURVEYOR' }
      });

      // Get survey statistics
      const totalSurveys = await this.prisma.survey.count();
      const completedSurveys = await this.prisma.survey.count({
        where: { status: 'COMPLETED' }
      });

      // Get opportunity statistics
      const totalOpportunities = await this.prisma.opportunityProgress.count();
      const activeOpportunities = await this.prisma.opportunityProgress.count({
        where: { status: 'IN_PROGRESS' }
      });

      // Get autosave statistics
      const totalAutoSaves = await this.prisma.autoSave.count();
      const recentAutoSaves = await this.prisma.autoSave.count({
        where: {
          lastSavedAt: {
            gte: new Date(Date.now() - 24 * 60 * 60 * 1000) // Last 24 hours
          }
        }
      });

      // Get system health metrics
      const systemHealth = await this.getSystemHealth();

      // Calculate completion rates
      const surveyCompletionRate = totalSurveys > 0 ? (completedSurveys / totalSurveys) * 100 : 0;

      // Get user growth data (last 6 months)
      const userGrowth = await this.getUserGrowthData();

      // Get top features usage (mock data for now)
      const topFeatures = await this.getTopFeaturesUsage();

      // Get daily/weekly/monthly active users
      const dailyActiveUsers = await this.getActiveUsersCount(1);
      const weeklyActiveUsers = await this.getActiveUsersCount(7);
      const monthlyActiveUsers = await this.getActiveUsersCount(30);

      const analytics = {
        // User metrics
        totalUsers,
        activeUsers,
        adminUsers,
        surveyorUsers,
        dailyActiveUsers,
        weeklyActiveUsers,
        monthlyActiveUsers,

        // Survey metrics
        totalSurveys,
        completedSurveys,
        surveyCompletionRate: Math.round(surveyCompletionRate * 100) / 100,

        // Opportunity metrics
        totalOpportunities,
        activeOpportunities,

        // System metrics
        totalAutoSaves,
        recentAutoSaves,
        systemUptime: '99.9%', // This would come from monitoring system
        averageResponseTime: '245ms', // This would come from monitoring system
        errorRate: '0.1%', // This would come from monitoring system

        // Growth and usage data
        userGrowth,
        topFeatures,

        // System health
        systemHealth,

        // Timestamps
        lastUpdated: new Date().toISOString(),
        generatedAt: new Date().toISOString()
      };

      this.logger.log(`Analytics generated successfully: ${totalUsers} users, ${totalSurveys} surveys`);
      return analytics;

    } catch (error) {
      this.logger.error('Error fetching system analytics:', error);
      throw new Error('Failed to fetch system analytics');
    }
  }

  /**
   * Get system health status
   */
  private async getSystemHealth() {
    try {
      // Check database connectivity
      const dbHealth = await this.checkDatabaseHealth();
      
      // Check API health (mock for now)
      const apiHealth = 'healthy';
      
      // Check storage health (mock for now)
      const storageHealth = 'healthy';
      
      // Check integrations health
      const integrationsHealth = await this.checkIntegrationsHealth();

      return {
        database: dbHealth,
        api: apiHealth,
        storage: storageHealth,
        integrations: integrationsHealth
      };
    } catch (error) {
      this.logger.error('Error checking system health:', error);
      return {
        database: 'error',
        api: 'error',
        storage: 'error',
        integrations: 'error'
      };
    }
  }

  /**
   * Check database health
   */
  private async checkDatabaseHealth(): Promise<string> {
    try {
      await this.prisma.$queryRaw`SELECT 1`;
      return 'healthy';
    } catch (error) {
      this.logger.error('Database health check failed:', error);
      return 'error';
    }
  }

  /**
   * Check integrations health
   */
  private async checkIntegrationsHealth(): Promise<string> {
    try {
      // Check GHL integration (mock for now)
      // In real implementation, you would ping GHL API
      return 'warning'; // Mock warning status
    } catch (error) {
      this.logger.error('Integrations health check failed:', error);
      return 'error';
    }
  }

  /**
   * Get user growth data for the last 6 months
   */
  private async getUserGrowthData() {
    try {
      const sixMonthsAgo = new Date();
      sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);

      const userGrowth: Array<{ month: string; users: number }> = [];
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      
      for (let i = 5; i >= 0; i--) {
        const startDate = new Date();
        startDate.setMonth(startDate.getMonth() - i);
        startDate.setDate(1);
        startDate.setHours(0, 0, 0, 0);

        const endDate = new Date(startDate);
        endDate.setMonth(endDate.getMonth() + 1);
        endDate.setDate(0);
        endDate.setHours(23, 59, 59, 999);

        const userCount = await this.prisma.user.count({
          where: {
            createdAt: {
              gte: startDate,
              lte: endDate
            }
          }
        });

        userGrowth.push({
          month: months[startDate.getMonth()],
          users: userCount
        });
      }

      return userGrowth;
    } catch (error) {
      this.logger.error('Error fetching user growth data:', error);
      return [];
    }
  }

  /**
   * Get top features usage (mock data for now)
   */
  private async getTopFeaturesUsage() {
    try {
      // In real implementation, you would track feature usage
      // For now, return mock data based on available features
      return [
        { name: 'Survey System', usage: 89 },
        { name: 'Calculator', usage: 76 },
        { name: 'Document Signing', usage: 65 },
        { name: 'Admin Panel', usage: 23 },
        { name: 'User Management', usage: 45 },
        { name: 'GHL Integration', usage: 67 }
      ];
    } catch (error) {
      this.logger.error('Error fetching top features usage:', error);
      return [];
    }
  }

  /**
   * Get active users count for a specific period
   */
  private async getActiveUsersCount(days: number): Promise<number> {
    try {
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const activeUsers = await this.prisma.user.count({
        where: {
          lastLoginAt: {
            gte: startDate
          }
        }
      });

      return activeUsers;
    } catch (error) {
      this.logger.error(`Error fetching active users for ${days} days:`, error);
      return 0;
    }
  }

  /**
   * Get system logs
   */
  async getSystemLogs(filter?: string, limit: number = 100) {
    try {
      this.logger.log(`Fetching system logs with filter: ${filter || 'all'}`);

      // In a real implementation, you would have a logs table or use a logging service
      // For now, return mock data
      const mockLogs = [
        {
          id: '1',
          timestamp: new Date().toISOString(),
          level: 'INFO',
          message: 'User authentication successful',
          userId: 'user123',
          action: 'LOGIN',
          details: { ip: '192.168.1.1', userAgent: 'Mozilla/5.0...' }
        },
        {
          id: '2',
          timestamp: new Date(Date.now() - 300000).toISOString(),
          level: 'WARN',
          message: 'High memory usage detected',
          userId: null,
          action: 'SYSTEM',
          details: { memoryUsage: '85%', threshold: '80%' }
        },
        {
          id: '3',
          timestamp: new Date(Date.now() - 600000).toISOString(),
          level: 'ERROR',
          message: 'Failed to sync with GHL',
          userId: 'user456',
          action: 'GHL_SYNC',
          details: { error: 'Connection timeout', retryCount: 3 }
        },
        {
          id: '4',
          timestamp: new Date(Date.now() - 900000).toISOString(),
          level: 'INFO',
          message: 'Survey completed successfully',
          userId: 'user789',
          action: 'SURVEY_COMPLETE',
          details: { opportunityId: 'opp123', duration: '15m 32s' }
        },
        {
          id: '5',
          timestamp: new Date(Date.now() - 1200000).toISOString(),
          level: 'INFO',
          message: 'Admin panel accessed',
          userId: 'admin123',
          action: 'ADMIN_ACCESS',
          details: { section: 'users', duration: '5m 12s' }
        }
      ];

      // Filter logs if filter is provided
      let filteredLogs = mockLogs;
      if (filter && filter !== 'all') {
        filteredLogs = mockLogs.filter(log => 
          log.level.toLowerCase() === filter.toLowerCase()
        );
      }

      // Limit results
      filteredLogs = filteredLogs.slice(0, limit);

      this.logger.log(`Returning ${filteredLogs.length} system logs`);
      return filteredLogs;

    } catch (error) {
      this.logger.error('Error fetching system logs:', error);
      throw new Error('Failed to fetch system logs');
    }
  }

  /**
   * Get user activity summary
   */
  async getUserActivitySummary(userId?: string) {
    try {
      this.logger.log(`Fetching user activity summary${userId ? ` for user ${userId}` : ''}`);

      const whereClause = userId ? { userId } : {};

      // Get recent activities
      const recentActivities = await this.prisma.autoSave.findMany({
        where: {
          ...whereClause,
          lastSavedAt: {
            gte: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000) // Last 7 days
          }
        },
        orderBy: { lastSavedAt: 'desc' },
        take: 50,
        include: {
          user: {
            select: {
              id: true,
              name: true,
              email: true,
              role: true
            }
          }
        }
      });

      // Get user statistics
      const userStats = {
        totalAutoSaves: await this.prisma.autoSave.count({ where: whereClause }),
        recentAutoSaves: recentActivities.length,
        lastActivity: recentActivities[0]?.lastSavedAt || null
      };

      return {
        userStats,
        recentActivities: recentActivities.map(activity => ({
          id: activity.id,
          userId: activity.userId,
          userName: activity.user?.name,
          userEmail: activity.user?.email,
          userRole: activity.user?.role,
          opportunityId: activity.opportunityId,
          lastSavedAt: activity.lastSavedAt,
          dataSize: JSON.stringify(activity.data).length
        }))
      };

    } catch (error) {
      this.logger.error('Error fetching user activity summary:', error);
      throw new Error('Failed to fetch user activity summary');
    }
  }

  /**
   * Get system performance metrics
   */
  async getSystemPerformanceMetrics() {
    try {
      this.logger.log('Fetching system performance metrics...');

      // Get database performance metrics
      const dbMetrics = await this.getDatabaseMetrics();

      // Get API performance metrics (mock for now)
      const apiMetrics = {
        averageResponseTime: '245ms',
        requestsPerMinute: 156,
        errorRate: '0.1%',
        uptime: '99.9%'
      };

      // Get memory and CPU usage (mock for now)
      const systemMetrics = {
        memoryUsage: '65%',
        cpuUsage: '23%',
        diskUsage: '45%',
        networkLatency: '12ms'
      };

      return {
        database: dbMetrics,
        api: apiMetrics,
        system: systemMetrics,
        lastUpdated: new Date().toISOString()
      };

    } catch (error) {
      this.logger.error('Error fetching system performance metrics:', error);
      throw new Error('Failed to fetch system performance metrics');
    }
  }

  /**
   * Get database performance metrics
   */
  private async getDatabaseMetrics() {
    try {
      const startTime = Date.now();
      
      // Test query performance
      await this.prisma.user.findMany({ take: 1 });
      
      const queryTime = Date.now() - startTime;

      return {
        queryTime: `${queryTime}ms`,
        connectionPool: 'healthy',
        slowQueries: 0,
        deadlocks: 0
      };
    } catch (error) {
      this.logger.error('Error getting database metrics:', error);
      return {
        queryTime: 'error',
        connectionPool: 'error',
        slowQueries: 0,
        deadlocks: 0
      };
    }
  }

  /**
   * Get all users with their basic information
   */
  async getAllUsers() {
    try {
      this.logger.log('Fetching all users...');

      const users = await this.prisma.user.findMany({
        select: {
          id: true,
          name: true,
          email: true,
          role: true,
          status: true,
          createdAt: true,
          lastLoginAt: true,
          ghlUserId: true
        },
        orderBy: { createdAt: 'desc' }
      });

      this.logger.log(`Found ${users.length} users`);
      return users;

    } catch (error) {
      this.logger.error('Error fetching all users:', error);
      throw new Error('Failed to fetch users');
    }
  }

  /**
   * Get user opportunities summary (lightweight version)
   */
  async getUserOpportunitiesSummary(userId: string) {
    try {
      this.logger.log(`Fetching opportunities summary for user: ${userId}`);

      // Get user details first
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new Error('User not found');
      }

      // Get opportunities for the user
      const opportunities = await this.opportunitiesService.getOpportunities(userId);
      const allOpportunities = opportunities.opportunities || [];

      // Return only essential information
      const summary = allOpportunities.map(opp => ({
        id: opp.id,
        name: opp.name,
        stageName: opp.stageName,
        status: opp.status,
        createdAt: opp.createdAt,
        updatedAt: opp.updatedAt
      }));

      this.logger.log(`Found ${allOpportunities.length} opportunities for user ${user.name}`);
      return {
        user: {
          id: user.id,
          name: user.name,
          email: user.email,
          role: user.role
        },
        opportunitiesSummary: summary,
        total: allOpportunities.length
      };

    } catch (error) {
      this.logger.error(`Error fetching opportunities summary for user ${userId}:`, error);
      throw new Error(`Failed to fetch opportunities summary for user: ${error.message}`);
    }
  }

  /**
   * Get user opportunities by user ID with pagination and size limits
   */
  async getUserOpportunities(userId: string, page: number = 1, limit: number = 50) {
    try {
      this.logger.log(`Fetching opportunities for user: ${userId} (page: ${page}, limit: ${limit})`);

      // Get user details first
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new Error('User not found');
      }

      // Get opportunities for the user
      const opportunities = await this.opportunitiesService.getOpportunities(userId);
      const allOpportunities = opportunities.opportunities || [];
      const total = opportunities.total || 0;

      this.logger.log(`Found ${allOpportunities.length} opportunities for user ${user.name}`);

      // Implement pagination to prevent large responses
      const startIndex = (page - 1) * limit;
      const endIndex = startIndex + limit;
      const paginatedOpportunities = allOpportunities.slice(startIndex, endIndex);

      // Calculate pagination metadata
      const totalPages = Math.ceil(total / limit);
      const hasNextPage = page < totalPages;
      const hasPrevPage = page > 1;

      // Log warning if data is large
      if (allOpportunities.length > 100) {
        this.logger.warn(`Large dataset detected: ${allOpportunities.length} opportunities. Using pagination to prevent JSON serialization errors.`);
      }

      return {
        user: {
          id: user.id,
          name: user.name,
          email: user.email,
          role: user.role,
          ghlUserId: user.ghlUserId
        },
        opportunities: paginatedOpportunities,
        pagination: {
          currentPage: page,
          totalPages,
          totalItems: total,
          itemsPerPage: limit,
          hasNextPage,
          hasPrevPage
        }
      };

    } catch (error) {
      this.logger.error(`Error fetching opportunities for user ${userId}:`, error);
      throw new Error(`Failed to fetch opportunities for user: ${error.message}`);
    }
  }

  /**
   * Get user autosaved survey data
   */
  async getUserAutosavedSurveyData(userId: string, opportunityId?: string) {
    try {
      this.logger.log(`Fetching autosaved survey data for user: ${userId}${opportunityId ? `, opportunity: ${opportunityId}` : ''}`);

      // Get user details first
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new Error('User not found');
      }

      // Build where clause
      const whereClause: any = { userId };
      if (opportunityId) {
        whereClause.opportunityId = opportunityId;
      }

      // Get autosaved data
      const autosavedData = await this.prisma.autoSave.findMany({
        where: whereClause,
        orderBy: { lastSavedAt: 'desc' },
        include: {
          user: {
            select: {
              id: true,
              name: true,
              email: true,
              role: true
            }
          }
        }
      });

      // Process the data to extract survey-specific information
      const surveyData = autosavedData.map(item => {
        const data = item.data as any;
        return {
          id: item.id,
          opportunityId: item.opportunityId,
          userId: item.userId,
          userName: item.user?.name,
          userEmail: item.user?.email,
          lastSavedAt: item.lastSavedAt,
          createdAt: item.createdAt,
          updatedAt: item.updatedAt,
          // Extract survey page data
          surveyPages: {
            page1: data?.page1 || null,
            page2: data?.page2 || null,
            page3: data?.page3 || null,
            page4: data?.page4 || null,
            page5: data?.page5 || null,
            page6: data?.page6 || null,
            page7: data?.page7 || null,
            page8: data?.page8 || null
          },
          // Extract images if any
          images: data?.images || {},
          // Raw data for debugging
          rawData: data
        };
      });

      this.logger.log(`Found ${surveyData.length} autosaved survey records for user ${user.name}`);
      return {
        user: {
          id: user.id,
          name: user.name,
          email: user.email,
          role: user.role
        },
        surveyData,
        total: surveyData.length
      };

    } catch (error) {
      this.logger.error(`Error fetching autosaved survey data for user ${userId}:`, error);
      throw new Error(`Failed to fetch autosaved survey data: ${error.message}`);
    }
  }

  /**
   * Get user autosaved calculator data
   */
  async getUserAutosavedCalculatorData(userId: string, opportunityId?: string) {
    try {
      this.logger.log(`Fetching autosaved calculator data for user: ${userId}${opportunityId ? `, opportunity: ${opportunityId}` : ''}`);

      // Get user details first
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new Error('User not found');
      }

      // Build where clause
      const whereClause: any = { userId };
      if (opportunityId) {
        whereClause.opportunityId = opportunityId;
      }

      // Get autosaved data
      const autosavedData = await this.prisma.autoSave.findMany({
        where: whereClause,
        orderBy: { lastSavedAt: 'desc' },
        include: {
          user: {
            select: {
              id: true,
              name: true,
              email: true,
              role: true
            }
          }
        }
      });

      // Process the data to extract calculator-specific information
      const calculatorData = autosavedData.map(item => {
        const data = item.data as any;
        return {
          id: item.id,
          opportunityId: item.opportunityId,
          userId: item.userId,
          userName: item.user?.name,
          userEmail: item.user?.email,
          lastSavedAt: item.lastSavedAt,
          createdAt: item.createdAt,
          updatedAt: item.updatedAt,
          // Extract calculator types and their data
          calculatorTypes: {
            offPeak: data?.offPeak || null,
            flux: data?.flux || null,
            epvs: data?.epvs || null
          },
          // Extract progress data
          progress: data?.progress || {},
          // Raw data for debugging
          rawData: data
        };
      });

      this.logger.log(`Found ${calculatorData.length} autosaved calculator records for user ${user.name}`);
      return {
        user: {
          id: user.id,
          name: user.name,
          email: user.email,
          role: user.role
        },
        calculatorData,
        total: calculatorData.length
      };

    } catch (error) {
      this.logger.error(`Error fetching autosaved calculator data for user ${userId}:`, error);
      throw new Error(`Failed to fetch autosaved calculator data: ${error.message}`);
    }
  }

  /**
   * Get comprehensive user data (opportunities + autosaved data) with pagination
   */
  async getUserComprehensiveData(userId: string, page: number = 1, limit: number = 25) {
    try {
      this.logger.log(`Fetching comprehensive data for user: ${userId} (page: ${page}, limit: ${limit})`);

      // Get user details
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new Error('User not found');
      }

      // Get all data in parallel with pagination
      const [opportunities, surveyData, calculatorData] = await Promise.all([
        this.getUserOpportunities(userId, page, limit),
        this.getUserAutosavedSurveyData(userId),
        this.getUserAutosavedCalculatorData(userId)
      ]);

      // Check total data size to prevent JSON serialization errors
      const totalDataSize = JSON.stringify({
        opportunities: opportunities.opportunities,
        surveyData: surveyData.surveyData,
        calculatorData: calculatorData.calculatorData
      }).length;

      if (totalDataSize > 1000000) { // 1MB limit
        this.logger.warn(`Large comprehensive dataset detected (${totalDataSize} bytes). Consider using pagination.`);
      }

      return {
        user: {
          id: user.id,
          name: user.name,
          email: user.email,
          role: user.role,
          status: user.status,
          createdAt: user.createdAt,
          lastLoginAt: user.lastLoginAt,
          ghlUserId: user.ghlUserId
        },
        opportunities: opportunities.opportunities,
        opportunitiesPagination: opportunities.pagination,
        surveyData: surveyData.surveyData,
        surveyDataTotal: surveyData.total,
        calculatorData: calculatorData.calculatorData,
        calculatorDataTotal: calculatorData.total,
        lastUpdated: new Date().toISOString()
      };

    } catch (error) {
      this.logger.error(`Error fetching comprehensive data for user ${userId}:`, error);
      throw new Error(`Failed to fetch comprehensive user data: ${error.message}`);
    }
  }
}

