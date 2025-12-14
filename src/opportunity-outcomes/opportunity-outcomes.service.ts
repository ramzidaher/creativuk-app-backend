import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { ConfigService } from '@nestjs/config';

export interface OpportunityOutcomeData {
  ghlOpportunityId: string;
  userId: string;
  outcome: 'WON' | 'LOST' | 'ABANDONED' | 'IN_PROGRESS';
  value?: number;
  notes?: string;
  stageAtOutcome?: string;
}

export interface WinLossStats {
  totalOpportunities: number;
  won: number;
  lost: number;
  abandoned: number;
  inProgress: number;
  totalValue: number;
  wonValue: number;
  conversionRate: number;
  averageDealValue: number;
  averageDuration: number;
}

export interface UserWinLossStats extends WinLossStats {
  userId: string;
  userName: string;
  userEmail: string;
  period: {
    start: Date;
    end: Date;
  };
}

@Injectable()
export class OpportunityOutcomesService {
  private readonly logger = new Logger(OpportunityOutcomesService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly goHighLevelService: GoHighLevelService,
    private readonly configService: ConfigService,
  ) {}

  private getGhlCredentials() {
    const accessToken = this.configService.get<string>('GOHIGHLEVEL_API_TOKEN');
    
    if (!accessToken) {
      this.logger.warn('GHL API token not configured - returning empty response');
      return null;
    }
    
    // Extract location ID from the JWT token (same as OpportunitiesService)
    try {
      const tokenData = this.extractTokenData(accessToken);
      const locationId = tokenData.locationId;
      
      if (!locationId) {
        this.logger.warn('Location ID not found in GHL token - returning empty response');
        return null;
      }
      
      return { accessToken, locationId };
    } catch (error) {
      this.logger.warn('Error extracting location ID from GHL token:', error.message);
      return null;
    }
  }

  private extractTokenData(token: string): any {
    try {
      // JWT tokens have 3 parts separated by dots: header.payload.signature
      const parts = token.split('.');
      if (parts.length !== 3) {
        throw new Error('Invalid JWT token format');
      }
      
      // Decode the payload (second part)
      const payload = parts[1];
      // Add padding if needed for base64 decoding
      const paddedPayload = payload + '='.repeat((4 - payload.length % 4) % 4);
      const decodedPayload = Buffer.from(paddedPayload, 'base64').toString('utf-8');
      
      return JSON.parse(decodedPayload);
    } catch (error) {
      throw new Error(`Failed to decode JWT token: ${error.message}`);
    }
  }

  /**
   * Record an opportunity outcome (win/loss/abandoned)
   */
  async recordOutcome(data: OpportunityOutcomeData): Promise<any> {
    this.logger.log(`Recording outcome for opportunity ${data.ghlOpportunityId}: ${data.outcome}`);

    try {
      // Check if outcome already exists
      const existingOutcome = await this.prisma.opportunityOutcome.findUnique({
        where: { ghlOpportunityId: data.ghlOpportunityId },
      });

      if (existingOutcome) {
        // Update existing outcome
        const updatedOutcome = await this.prisma.opportunityOutcome.update({
          where: { ghlOpportunityId: data.ghlOpportunityId },
          data: {
            outcome: data.outcome,
            value: data.value,
            notes: data.notes,
            stageAtOutcome: data.stageAtOutcome,
            ghlUpdatedAt: new Date(),
            updatedAt: new Date(),
          },
        });

        this.logger.log(`Updated outcome for opportunity ${data.ghlOpportunityId}`);
        return updatedOutcome;
      } else {
        // Create new outcome
        const newOutcome = await this.prisma.opportunityOutcome.create({
          data: {
            ghlOpportunityId: data.ghlOpportunityId,
            userId: data.userId,
            outcome: data.outcome,
            value: data.value,
            notes: data.notes,
            stageAtOutcome: data.stageAtOutcome,
            ghlUpdatedAt: new Date(),
          },
        });

        this.logger.log(`Created new outcome for opportunity ${data.ghlOpportunityId}`);
        return newOutcome;
      }
    } catch (error) {
      this.logger.error(`Error recording outcome for opportunity ${data.ghlOpportunityId}:`, error.stack);
      throw error;
    }
  }

  /**
   * Get win/loss statistics for a specific user
   */
  async getUserWinLossStats(
    userId: string,
    startDate?: Date,
    endDate?: Date
  ): Promise<UserWinLossStats> {
    this.logger.log(`Getting win/loss stats for user ${userId}`);

    try {
      // Get user info
      const user = await this.prisma.user.findUnique({
        where: { id: userId },
        select: { id: true, name: true, email: true, ghlUserId: true },
      });

      if (!user) {
        throw new NotFoundException('User not found');
      }

      // Build date filter
      const dateFilter: any = {};
      if (startDate || endDate) {
        dateFilter.createdAt = {};
        if (startDate) dateFilter.createdAt.gte = startDate;
        if (endDate) dateFilter.createdAt.lte = endDate;
      }

      // Get all outcomes for the user
      const outcomes = await this.prisma.opportunityOutcome.findMany({
        where: {
          userId: userId,
          ...dateFilter,
        },
        orderBy: { createdAt: 'desc' },
      });

      // Get opportunities in progress from GHL
      let inProgressCount = 0;
      try {
        const credentials = this.getGhlCredentials();
        if (credentials && user.ghlUserId) {
          const ghlOpportunities = await this.goHighLevelService.getOpportunities(
            credentials.accessToken,
            credentials.locationId
          );
          
          // Filter opportunities assigned to this user that are not won/lost
          const userOpportunities = ghlOpportunities.filter(opp => 
            opp.assignedTo === user.ghlUserId && 
            opp.status !== 'won' && 
            opp.status !== 'lost'
          );
          
          inProgressCount = userOpportunities.length;
          this.logger.log(`Found ${inProgressCount} opportunities in progress for user ${userId}`);
        }
      } catch (error) {
        this.logger.warn(`Could not fetch GHL opportunities for user ${userId}:`, error.message);
      }

      // Calculate statistics
      const stats = this.calculateWinLossStats(outcomes, inProgressCount);

      return {
        ...stats,
        userId: user.id,
        userName: user.name || 'Unknown',
        userEmail: user.email,
        period: {
          start: startDate || new Date(0),
          end: endDate || new Date(),
        },
      };
    } catch (error) {
      this.logger.error(`Error getting win/loss stats for user ${userId}:`, error.stack);
      throw error;
    }
  }

  /**
   * Get win/loss statistics for all users (admin view)
   */
  async getAllUsersWinLossStats(
    startDate?: Date,
    endDate?: Date
  ): Promise<UserWinLossStats[]> {
    this.logger.log('Getting win/loss stats for all users');

    try {
      // Get all users
      const users = await this.prisma.user.findMany({
        where: { status: 'ACTIVE' },
        select: { id: true, name: true, email: true },
      });

      // Get stats for each user
      const userStats = await Promise.all(
        users.map(async (user) => {
          return this.getUserWinLossStats(user.id, startDate, endDate);
        })
      );

      // Sort by total value (descending)
      return userStats.sort((a, b) => b.wonValue - a.wonValue);
    } catch (error) {
      this.logger.error('Error getting all users win/loss stats:', error.stack);
      throw error;
    }
  }

  /**
   * Get overall company statistics
   */
  async getOverallStats(startDate?: Date, endDate?: Date): Promise<WinLossStats> {
    this.logger.log('Getting overall company win/loss stats');

    try {
      // Build date filter
      const dateFilter: any = {};
      if (startDate || endDate) {
        dateFilter.createdAt = {};
        if (startDate) dateFilter.createdAt.gte = startDate;
        if (endDate) dateFilter.createdAt.lte = endDate;
      }

      // Get all outcomes
      const outcomes = await this.prisma.opportunityOutcome.findMany({
        where: dateFilter,
        orderBy: { createdAt: 'desc' },
      });

      return this.calculateWinLossStats(outcomes, 0);
    } catch (error) {
      this.logger.error('Error getting overall stats:', error.stack);
      throw error;
    }
  }

  /**
   * Sync opportunity outcomes from GHL API
   */
  async syncFromGHL(userId: string): Promise<{ synced: number; errors: number }> {
    this.logger.log(`Syncing opportunity outcomes from GHL for user ${userId}`);

    try {
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        throw new Error('GHL credentials not configured');
      }

      // Get user's opportunities from GHL
      const opportunities = await this.goHighLevelService.getOpportunities(
        credentials.accessToken,
        credentials.locationId
      );

      let synced = 0;
      let errors = 0;

      // Process each opportunity
      for (const opportunity of opportunities) {
        try {
          // Determine outcome based on GHL data
          const outcome = this.determineOutcomeFromGHL(opportunity);
          
          if (outcome && outcome !== 'IN_PROGRESS') {
            await this.recordOutcome({
              ghlOpportunityId: opportunity.id,
              userId: userId,
              outcome: outcome,
              value: opportunity.monetaryValue || 0,
              stageAtOutcome: opportunity.pipelineStageId,
              notes: `Synced from GHL - Status: ${opportunity.status}`,
            });
            synced++;
          }
        } catch (error) {
          this.logger.error(`Error syncing opportunity ${opportunity.id}:`, error.message);
          errors++;
        }
      }

      this.logger.log(`Synced ${synced} opportunities, ${errors} errors`);
      return { synced, errors };
    } catch (error) {
      this.logger.error('Error syncing from GHL:', error.stack);
      throw error;
    }
  }

  /**
   * Get recent outcomes for a user
   */
  async getRecentOutcomes(userId: string, limit: number = 10): Promise<any[]> {
    this.logger.log(`Getting recent outcomes for user ${userId}`);

    try {
      const outcomes = await this.prisma.opportunityOutcome.findMany({
        where: { userId },
        orderBy: { createdAt: 'desc' },
        take: limit,
        include: {
          user: {
            select: { name: true, email: true },
          },
        },
      });

      return outcomes;
    } catch (error) {
      this.logger.error(`Error getting recent outcomes for user ${userId}:`, error.stack);
      throw error;
    }
  }

  /**
   * Get outcome by GHL opportunity ID
   */
  async getOutcomeByOpportunityId(ghlOpportunityId: string): Promise<any | null> {
    this.logger.log(`Getting outcome for opportunity ${ghlOpportunityId}`);

    try {
      const outcome = await this.prisma.opportunityOutcome.findUnique({
        where: { ghlOpportunityId },
        include: {
          user: {
            select: { name: true, email: true },
          },
        },
      });

      if (!outcome) {
        this.logger.log(`No outcome found for opportunity ${ghlOpportunityId}`);
        return null;
      }

      this.logger.log(`Found outcome for opportunity ${ghlOpportunityId}: ${outcome.outcome}`);
      return outcome;
    } catch (error) {
      this.logger.error(`Error getting outcome for opportunity ${ghlOpportunityId}:`, error.stack);
      throw error;
    }
  }

  /**
   * Calculate win/loss statistics from outcomes array
   */
  private calculateWinLossStats(outcomes: any[], inProgressCount: number = 0): WinLossStats {
    const won = outcomes.filter(o => o.outcome === 'WON').length;
    const lost = outcomes.filter(o => o.outcome === 'LOST').length;
    const abandoned = outcomes.filter(o => o.outcome === 'ABANDONED').length;
    const inProgress = inProgressCount; // Use the count from GHL

    const totalOpportunities = won + lost + abandoned + inProgress;

    const totalValue = outcomes.reduce((sum, o) => sum + (o.value || 0), 0);
    const wonValue = outcomes
      .filter(o => o.outcome === 'WON')
      .reduce((sum, o) => sum + (o.value || 0), 0);

    const conversionRate = totalOpportunities > 0 ? (won / totalOpportunities) * 100 : 0;
    const averageDealValue = won > 0 ? wonValue / won : 0;

    // Calculate average duration (simplified - would need more complex logic for actual duration)
    const averageDuration = 0; // TODO: Implement duration calculation

    return {
      totalOpportunities,
      won,
      lost,
      abandoned,
      inProgress,
      totalValue,
      wonValue,
      conversionRate,
      averageDealValue,
      averageDuration,
    };
  }

  /**
   * Determine outcome from GHL opportunity data
   */
  private determineOutcomeFromGHL(opportunity: any): 'WON' | 'LOST' | 'ABANDONED' | 'IN_PROGRESS' | null {
    const status = opportunity.status?.toLowerCase();
    const stageId = opportunity.pipelineStageId;

    // Check for won status/tags
    if (status === 'won' || status === 'closed won' || status === 'sold') {
      return 'WON';
    }

    // Check for lost status/tags
    if (status === 'lost' || status === 'closed lost' || status === 'no sale') {
      return 'LOST';
    }

    // Check for abandoned status
    if (status === 'abandoned' || status === 'inactive') {
      return 'ABANDONED';
    }

    // Check tags for won/lost indicators
    const tags = opportunity.contact?.tags || opportunity.tags || [];
    if (Array.isArray(tags)) {
      for (const tag of tags) {
        const tagText = (typeof tag === 'string' ? tag : tag.name || tag.title || '').toLowerCase();
        if (tagText.includes('won') || tagText.includes('sold') || tagText.includes('closed won')) {
          return 'WON';
        }
        if (tagText.includes('lost') || tagText.includes('no sale') || tagText.includes('closed lost')) {
          return 'LOST';
        }
      }
    }

    // Default to in progress
    return 'IN_PROGRESS';
  }
}
