import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { UserService } from '../user/user.service';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { ConfigService } from '@nestjs/config';
import { UserRole } from '../auth/dto/auth.dto';
import { DynamicSurveyorService } from './dynamic-surveyor.service';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class OpportunitiesService {
  private readonly logger = new Logger(OpportunitiesService.name);

  constructor(
    private readonly userService: UserService,
    private readonly goHighLevelService: GoHighLevelService,
    private readonly configService: ConfigService,
    private readonly dynamicSurveyorService: DynamicSurveyorService,
    private readonly prisma: PrismaService,
  ) {}

  private getGhlCredentials() {
    const accessToken = this.configService.get<string>('GOHIGHLEVEL_API_TOKEN');
    
    if (!accessToken) {
      this.logger.warn('GHL API token not configured - returning empty response');
      return null;
    }
    
          // Extract location ID from the JWT token
      try {
        const tokenData = this.extractTokenData(accessToken);
        const locationId = tokenData.locationId;
        
        if (!locationId) {
          this.logger.warn('Location ID not found in GHL token - returning empty response');
          return null;
        }
        
        this.logger.log(`Using GHL credentials - Location ID: ${locationId}`);
    return { accessToken, locationId };
      } catch (error) {
        this.logger.error('Error extracting location ID from GHL token:', error);
        return null;
      }
  }

  private async filterOpportunitiesByUserRoleAndAppointments(
    allOpportunities: any[],
    user: any,
    credentials: { accessToken: string; locationId: string }
  ): Promise<any[]> {
    let filteredOpportunities = allOpportunities;

    if (user.role === UserRole.SURVEYOR) {
      // For surveyors, show only opportunities assigned to them by GHL User ID
      this.logger.log(`Filtering opportunities for surveyor: ${user.name} (GHL ID: ${user.ghlUserId})`);
      
      this.logger.log(`Total opportunities before filtering: ${allOpportunities.length}`);

      // Filter opportunities by GHL User ID - much more reliable than name matching
      filteredOpportunities = allOpportunities.filter(opp => {
        const isAssigned = opp.assignedTo === user.ghlUserId;
        if (isAssigned) {
          this.logger.log(`✅ Opportunity assigned to ${user.name}: ${opp.name || opp.id} (assignedTo: ${opp.assignedTo})`);
        }
        return isAssigned;
      });

      this.logger.log(`Filtered to ${filteredOpportunities.length} opportunities for surveyor ${user.name}`);
      
      // Log some examples of filtered opportunities
      if (filteredOpportunities.length > 0) {
        const examples = filteredOpportunities.slice(0, 3).map(opp => 
          `${opp.name} (${opp.contact?.address?.postcode || 'No postcode'})`
        );
        this.logger.log(`Examples of filtered opportunities: ${examples.join(', ')}`);
      }
    } else if (user.role === UserRole.ADMIN) {
      // For admins, show all opportunities (no filtering)
      this.logger.log(`Filtering opportunities for admin: ${user.name}`);
      
      // Show all opportunities for admin
      filteredOpportunities = allOpportunities;
      this.logger.log(`Showing all ${filteredOpportunities.length} opportunities for admin`);
    }

    return filteredOpportunities;
  }

  async getOpportunities(userId: string) {
    const credentials = this.getGhlCredentials();
    if (!credentials) {
      this.logger.warn('GHL credentials not configured - returning empty response');
      return { opportunities: [], total: 0 };
    }

    try {
      // Get user details
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      this.logger.log(`🔍 MAIN OPPORTUNITIES: Starting analysis for user: ${user.name} (${user.role})`);

      // Get opportunities from AI and Manual stages
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      this.logger.log(`📊 Found ${allOpportunities.length} opportunities in survey stages`);

      // Filter opportunities based on user role
      const filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        allOpportunities,
        user,
        credentials
      );

      this.logger.log(`👤 Filtered to ${filteredOpportunities.length} opportunities for user ${user.name}`);

      // ENHANCE OPPORTUNITIES WITH APPOINTMENT INFO
      const opportunitiesWithAppointmentInfo = await this.enhanceOpportunitiesWithAppointmentInfo(
        filteredOpportunities,
        credentials
      );

      // Log summary
      const withAppointments = opportunitiesWithAppointmentInfo.filter(opp => opp.hasAppointment);
      const withoutAppointments = opportunitiesWithAppointmentInfo.filter(opp => !opp.hasAppointment);
      
      this.logger.log(`📈 MAIN OPPORTUNITIES SUMMARY:`);
      this.logger.log(`✅ With appointments: ${withAppointments.length}`);
      this.logger.log(`❌ Without appointments: ${withoutAppointments.length}`);
      this.logger.log(`📊 Total opportunities: ${opportunitiesWithAppointmentInfo.length}`);

      return {
        opportunities: opportunitiesWithAppointmentInfo,
        total: opportunitiesWithAppointmentInfo.length,
        summary: {
          withAppointments: withAppointments.length,
          withoutAppointments: withoutAppointments.length,
          total: opportunitiesWithAppointmentInfo.length,
        },
        user: {
          id: user.id,
          name: user.name,
          role: user.role,
        },
      };
    } catch (error) {
      this.logger.error(`Error in getOpportunities: ${error.message}`);
      throw error;
    }
  }

  async getOpportunitiesByStage(userId: string, stageName: string) {
    this.logger.log(`Fetching opportunities by stage: ${stageName} for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty stage opportunities response');
        return [];
      }
      
      const all = await this.goHighLevelService.getOpportunitiesWithStageNames(
        credentials.accessToken,
        credentials.locationId,
      );
      
      // Filter opportunities based on user role and name
      const filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        all,
        user,
        credentials
      );
      
      // Then filter by stage name
      return filteredOpportunities.filter(opp => 
        opp.stageName && opp.stageName.toLowerCase().includes(stageName.toLowerCase())
      );
    } catch (error) {
      this.logger.error('Error in getOpportunitiesByStage:', error.stack);
      throw error;
    }
  }

  async getAiVsManualOpportunities(userId: string) {
    this.logger.log(`Fetching AI vs Manual opportunities for userId: ${userId}`);
    try {
      let user = await this.userService.findById(userId);
      if (!user) {
        this.logger.warn(`User not found in DB for userId: ${userId} - attempting to create from GHL`);
        
        // Try to find user by GHL user ID if the ID looks like a GHL ID
        if (userId.length > 20) { // GHL user IDs are typically long
          user = await this.userService.findByGhlUserId(userId);
          if (user) {
            this.logger.log(`Found user by GHL user ID: ${user.username} (${user.id})`);
          }
        }
        
        if (!user) {
          this.logger.error(`User not found in DB and cannot be created for userId: ${userId}`);
          // For now, return empty opportunities instead of throwing error
          // This allows the workflow to continue even if user is not in DB
          this.logger.warn(`Returning empty opportunities for missing user: ${userId}`);
          return {
            ai: {
              opportunities: [],
              count: 0,
              totalValue: 0,
              stageName: '(AI Bot) Home Survey Booked'
            },
            manual: {
              opportunities: [],
              count: 0,
              totalValue: 0,
              stageName: '(Manual) Home Survey Booked'
            },
            summary: {
              totalOpportunities: 0,
              totalValue: 0
            }
          };
        }
      }

      this.logger.log(`User found in DB: ${user.username || user.name || 'Unknown'} (${user.id})`);

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty opportunities response');
        return {
          ai: {
            opportunities: [],
            count: 0,
            totalValue: 0,
            stageName: '(AI Bot) Home Survey Booked'
          },
          manual: {
            opportunities: [],
            count: 0,
            totalValue: 0,
            stageName: '(Manual) Home Survey Booked'
          },
          summary: {
            totalOpportunities: 0,
            totalValue: 0
          }
        };
      }
      
      this.logger.log(`Using GHL v1 API credentials for locationId: ${credentials.locationId}`);
      
      // Define the specific stage IDs we want
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      this.logger.log(`Fetching opportunities for AI stage: ${aiStageId}`);
      this.logger.log(`Fetching opportunities for Manual stage: ${manualStageId}`);
      
      // Add timeout to prevent hanging
      const timeoutPromise = new Promise((_, reject) => {
        setTimeout(() => reject(new Error('Request timeout after 60 seconds')), 60000);
      });

      const fetchPromise = this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      const all = await Promise.race([fetchPromise, timeoutPromise]) as any[];
      this.logger.log(`Fetched ${all.length} total opportunities from specific stages.`);

      // Filter opportunities based on user role and name
      const filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        all,
        user,
        credentials
      );

      // Log all stage names to help debug filtering
      const stageNames = [...new Set(filteredOpportunities.map(opp => opp.stageName).filter(Boolean))];
      this.logger.log(`Available stage names: ${stageNames.join(', ')}`);

      // Filter by specific stage IDs for AI and Manual opportunities
      const aiOpportunities = filteredOpportunities.filter(opp => opp.pipelineStageId === aiStageId);
      const manualOpportunities = filteredOpportunities.filter(opp => opp.pipelineStageId === manualStageId);
      
      this.logger.log(`Found ${aiOpportunities.length} AI opportunities and ${manualOpportunities.length} manual opportunities.`);

      // Enhance opportunities with contact location data (temporarily disabled to avoid rate limiting)
      // const enhancedAiOpportunities = await this.enhanceOpportunitiesWithContactData(aiOpportunities, credentials.accessToken);
      // const enhancedManualOpportunities = await this.enhanceOpportunitiesWithContactData(manualOpportunities, credentials.accessToken);
      
      // Extract postcodes from opportunity names (no API calls - addresses fetched on-demand)
      const enhancedAiOpportunities = aiOpportunities.map(opp => {
        let contactPostcode: string | null = null;
        
        // Extract postcode from opportunity name (e.g., "N12 9JA, Lisa Jones")
        if (opp.name) {
          const postcodeMatch = opp.name.match(/[A-Z]{1,2}\d{1,2}\s?\d[A-Z]{2}/i);
          if (postcodeMatch) {
            contactPostcode = postcodeMatch[0].toUpperCase();
          }
        }
        
        return {
          ...opp,
          contactPostcode,
          contactAddress: null, // Will be fetched on-demand
          address: null // Will be fetched on-demand
        };
      });
      
      const enhancedManualOpportunities = manualOpportunities.map(opp => {
        let contactPostcode: string | null = null;
        
        // Extract postcode from opportunity name (e.g., "N12 9JA, Lisa Jones")
        if (opp.name) {
          const postcodeMatch = opp.name.match(/[A-Z]{1,2}\d{1,2}\s?\d[A-Z]{2}/i);
          if (postcodeMatch) {
            contactPostcode = postcodeMatch[0].toUpperCase();
          }
        }
        
        return {
          ...opp,
          contactPostcode,
          contactAddress: null, // Will be fetched on-demand
          address: null // Will be fetched on-demand
        };
      });

      // Log some examples of opportunity IDs
      if (enhancedAiOpportunities.length > 0) {
        this.logger.log(`AI opportunity examples: ${enhancedAiOpportunities.slice(0, 3).map(opp => `${opp.name} (ID: ${opp.id})`).join(', ')}`);
      }
      if (enhancedManualOpportunities.length > 0) {
        this.logger.log(`Manual opportunity examples: ${enhancedManualOpportunities.slice(0, 3).map(opp => `${opp.name} (ID: ${opp.id})`).join(', ')}`);
      }

      const aiTotal = enhancedAiOpportunities.reduce((sum, opp) => sum + (opp.monetaryValue || 0), 0);
      const manualTotal = enhancedManualOpportunities.reduce((sum, opp) => sum + (opp.monetaryValue || 0), 0);
      
      const result = {
        ai: {
          opportunities: enhancedAiOpportunities,
          count: enhancedAiOpportunities.length,
          totalValue: aiTotal,
          stageName: '(AI Bot) Home Survey Booked'
        },
        manual: {
          opportunities: enhancedManualOpportunities,
          count: enhancedManualOpportunities.length,
          totalValue: manualTotal,
          stageName: '(Manual) Home Survey Booked'
        },
        summary: {
          totalOpportunities: filteredOpportunities.length,
          totalValue: aiTotal + manualTotal
        }
      };
      this.logger.log(`Successfully prepared AI vs Manual opportunities response.`);
      return result;
    } catch (error) {
      this.logger.error('Error in getAiVsManualOpportunities:', error.stack);
      throw error;
    }
  }

  async getAllOpportunities(userId: string) {
    this.logger.log(`Fetching ALL opportunities for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty all opportunities response');
        return [];
      }
      
      const all = await this.goHighLevelService.getOpportunitiesWithStageNames(
        credentials.accessToken,
        credentials.locationId,
      );
      
      this.logger.log(`Fetched ${all.length} total opportunities from GHL v1 API.`);
      return all;
    } catch (error) {
      this.logger.error('Error in getAllOpportunities:', error.stack);
      throw error;
    }
  }

  async getOpportunityStages(userId: string) {
    this.logger.log(`Fetching opportunity stages for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty stages response');
        return { stages: [] };
      }
      
      const all = await this.goHighLevelService.getOpportunitiesWithStageNames(
        credentials.accessToken,
        credentials.locationId,
      );
      
      const stages = [...new Set(all.map(opp => opp.stageName).filter(Boolean))];
      this.logger.log(`Found ${stages.length} unique stages: ${stages.join(', ')}`);
      return { stages };
    } catch (error) {
      this.logger.error('Error in getOpportunityStages:', error.stack);
      throw error;
    }
  }

  async getAllPipelineStagesWithIds(userId: string) {
    this.logger.log(`Fetching all pipeline stages with IDs for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty stages response');
        return { pipelines: [], stages: [], stageIdToName: {} };
      }
      
      // Get all pipelines with their stages
      const pipelines = await this.goHighLevelService.getPipelines(
        credentials.accessToken,
        credentials.locationId
      );

      // Extract all stages with their IDs and names
      const allStages: any[] = [];
      const stageIdToName: Record<string, string> = {};
      
      for (const pipeline of pipelines) {
        this.logger.log(`Pipeline: ${pipeline.name} (ID: ${pipeline.id})`);
        for (const stage of pipeline.stages || []) {
          if (stage && stage.id && stage.name) {
            allStages.push({
              id: stage.id,
              name: stage.name,
              pipelineId: pipeline.id,
              pipelineName: pipeline.name
            });
            stageIdToName[stage.id] = stage.name;
            this.logger.log(`  Stage: ${stage.name} (ID: ${stage.id})`);
          }
        }
      }

      this.logger.log(`Found ${allStages.length} stages across ${pipelines.length} pipelines`);
      
      return {
        pipelines: pipelines,
        stages: allStages,
        stageIdToName: stageIdToName
      };
    } catch (error) {
      this.logger.error('Error in getAllPipelineStagesWithIds:', error.stack);
      throw error;
    }
  }

  async getOpportunityById(userId: string, opportunityId: string) {
    this.logger.log(`Fetching opportunity by ID: ${opportunityId} for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return error
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - cannot fetch opportunity by ID');
        throw new NotFoundException('GHL credentials not configured');
      }
      
      // Use direct ID lookup instead of fetching all opportunities
      try {
        const opportunity = await this.goHighLevelService.getOpportunityById(
          credentials.accessToken,
          opportunityId,
          credentials.locationId,
        );
        
        // Get pipelines to map stage names
        const pipelines = await this.goHighLevelService.getPipelines(
          credentials.accessToken,
          credentials.locationId,
        );
        
        // Build a map of stageId -> stageName for each pipeline
        const stageIdToName: Record<string, string> = {};
        for (const pipeline of pipelines) {
          for (const stage of pipeline.stages || []) {
            if (stage && stage.id && stage.name) {
              stageIdToName[stage.id] = stage.name;
            }
          }
        }
        
        // Add stage name to the opportunity
        const opportunityWithStageName = {
          ...opportunity,
          stageName: stageIdToName[opportunity.pipelineStageId] || opportunity.stageName || '',
        };
        
        this.logger.log(`Found opportunity: ${opportunityWithStageName.name} (ID: ${opportunityWithStageName.id})`);
        return opportunityWithStageName;
      } catch (error) {
        this.logger.warn(`Opportunity with ID ${opportunityId} not found in any pipeline`);
        throw new NotFoundException('Opportunity not found');
      }
    } catch (error) {
      this.logger.error('Error fetching opportunity by ID:', error.stack);
      throw error;
    }
  }

  async getOpportunityNotes(userId: string, opportunityId: string) {
    this.logger.log(`Fetching notes for opportunity: ${opportunityId} for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty notes response');
        return [];
      }
      
      const notes = await this.goHighLevelService.getOpportunityNotes(
        credentials.accessToken,
        opportunityId,
      );
      
      this.logger.log(`Fetched ${notes.length} notes for opportunity ${opportunityId}`);
      return notes;
    } catch (error) {
      this.logger.error('Error fetching opportunity notes:', error.stack);
      throw error;
    }
  }

  async getSimpleOpportunities(userId: string) {
    this.logger.log(`Fetching simple opportunities for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty simple opportunities response');
        return [];
      }
      
      const all = await this.goHighLevelService.getOpportunitiesWithStageNames(
        credentials.accessToken,
        credentials.locationId,
      );
      
      // Filter opportunities based on user role and name (same logic as getOpportunities)
      const filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        all,
        user,
        credentials
      );
      
      // Return simplified opportunity data
      const simpleOpportunities = filteredOpportunities.map(opp => ({
        id: opp.id,
        name: opp.name,
        stageName: opp.stageName,
        monetaryValue: opp.monetaryValue,
        createdAt: opp.createdAt,
        updatedAt: opp.updatedAt,
        contact: opp.contact ? {
          id: opp.contact.id,
          firstName: opp.contact.firstName,
          lastName: opp.contact.lastName,
          email: opp.contact.email,
          phone: opp.contact.phone
        } : null
      }));
      
      this.logger.log(`Returning ${simpleOpportunities.length} simplified opportunities`);
      return simpleOpportunities;
    } catch (error) {
      this.logger.error('Error in getSimpleOpportunities:', error.stack);
      throw error;
    }
  }

  // Get opportunities assigned to James Barnett via appointments
  async getOpportunitiesByAssignedTeamMember(userId: string, teamMemberName: string = 'James Barnett') {
    this.logger.log(`Fetching opportunities assigned to ${teamMemberName} for userId: ${userId}`);
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found in DB for userId: ${userId}`);
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      
      // If GHL credentials are not configured, return empty response
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty team member opportunities response');
        return [];
      }
      
      // Fetch all opportunities
      const allOpportunities = await this.goHighLevelService.getOpportunitiesWithStageNames(
        credentials.accessToken,
        credentials.locationId,
      );
      
      // Fetch appointments assigned to the team member
      const appointments = await this.goHighLevelService.getAppointmentsByTeamMember(
        credentials.accessToken,
        credentials.locationId,
        teamMemberName
      );
      
      this.logger.log(`Found ${appointments.length} appointments assigned to ${teamMemberName}`);
      
      // Extract opportunity IDs from appointments
      const appointmentOpportunityIds = appointments.map((appointment: any) => {
        // Try different possible fields where opportunity ID might be stored
        return appointment.opportunityId || appointment.opportunity?.id || appointment.leadId || appointment.contactId;
      }).filter(Boolean);
      
      this.logger.log(`Found ${appointmentOpportunityIds.length} unique opportunity IDs from appointments`);
      
      // Filter opportunities that have appointments assigned to the team member
      const filteredOpportunities = allOpportunities.filter(opp => 
        appointmentOpportunityIds.includes(opp.id)
      );
      
      this.logger.log(`Found ${filteredOpportunities.length} opportunities with appointments assigned to ${teamMemberName}`);
      
      // Apply the same AI/Manual filtering as before
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      const aiOpportunities = filteredOpportunities.filter(opp => opp.pipelineStageId === aiStageId);
      const manualOpportunities = filteredOpportunities.filter(opp => opp.pipelineStageId === manualStageId);
      
      this.logger.log(`Found ${aiOpportunities.length} AI opportunities and ${manualOpportunities.length} manual opportunities assigned to ${teamMemberName}`);

      const aiTotal = aiOpportunities.reduce((sum, opp) => sum + (opp.monetaryValue || 0), 0);
      const manualTotal = manualOpportunities.reduce((sum, opp) => sum + (opp.monetaryValue || 0), 0);
      
      const result = {
        ai: {
          opportunities: aiOpportunities,
          count: aiOpportunities.length,
          totalValue: aiTotal,
          stageName: '(AI Bot) Home Survey Booked'
        },
        manual: {
          opportunities: manualOpportunities,
          count: manualOpportunities.length,
          totalValue: manualTotal,
          stageName: '(Manual) Home Survey Booked'
        },
        summary: {
          totalOpportunities: filteredOpportunities.length,
          totalValue: aiTotal + manualTotal,
          assignedTo: teamMemberName
        }
      };
      
      this.logger.log(`Successfully prepared opportunities assigned to ${teamMemberName}`);
      return result;
    } catch (error) {
      this.logger.error(`Error in getOpportunitiesByAssignedTeamMember:`, error.stack);
      throw error;
    }
  }

  async getOpportunitiesWithAppointments(userId: string) {
    const credentials = this.getGhlCredentials();
    if (!credentials) {
      this.logger.warn('GHL credentials not configured - returning empty response');
      return { opportunities: [], total: 0 };
    }

    try {
      // Get user details
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      this.logger.log(`🔍 ACCURATE APPOINTMENT MATCHING: Starting analysis for user: ${user.name} (${user.role})`);

      // Get opportunities from AI and Manual stages
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      this.logger.log(`📊 Found ${allOpportunities.length} opportunities in survey stages`);

      // Filter opportunities based on user role
      const filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        allOpportunities,
        user,
        credentials
      );

      this.logger.log(`👤 Filtered to ${filteredOpportunities.length} opportunities for user ${user.name}`);

      // ACCURATE APPOINTMENT ANALYSIS
      const opportunitiesWithAnalysis = await this.analyzeOpportunityAppointments(
        filteredOpportunities,
        credentials
      );

      // Classify opportunities based on appointment status
      const classifiedOpportunities = this.classifyOpportunities(opportunitiesWithAnalysis);

      // Log classification summary
      this.logger.log(`📈 APPOINTMENT CLASSIFICATION SUMMARY:`);
      this.logger.log(`✅ Confirmed with real appointments: ${classifiedOpportunities.confirmedWithAppointments.length}`);
      this.logger.log(`⚠️ Tagged as booked but no appointment: ${classifiedOpportunities.taggedButNoAppointment.length}`);
      this.logger.log(`❓ Multiple appointments (low confidence): ${classifiedOpportunities.multipleAppointments.length}`);
      this.logger.log(`❌ No appointments found: ${classifiedOpportunities.noAppointments.length}`);

      // Separate opportunities by type (AI vs Manual) for frontend compatibility
      const aiOpportunities = classifiedOpportunities.all.filter(opp => opp.type === 'ai');
      const manualOpportunities = classifiedOpportunities.all.filter(opp => opp.type === 'manual');

      return {
        opportunities: classifiedOpportunities.all,
        total: classifiedOpportunities.all.length,
        ai: {
          opportunities: aiOpportunities,
          total: aiOpportunities.length
        },
        manual: {
          opportunities: manualOpportunities,
          total: manualOpportunities.length
        },
        classification: {
          confirmedWithAppointments: classifiedOpportunities.confirmedWithAppointments.length,
          taggedButNoAppointment: classifiedOpportunities.taggedButNoAppointment.length,
          multipleAppointments: classifiedOpportunities.multipleAppointments.length,
          noAppointments: classifiedOpportunities.noAppointments.length,
        },
        user: {
          id: user.id,
          name: user.name,
          role: user.role,
        },
      };
    } catch (error) {
      this.logger.error(`Error in getOpportunitiesWithAppointments: ${error.message}`);
      throw error;
    }
  }

  // NEW: Optimized method for getting opportunities with appointments
  async getOpportunitiesWithAppointmentsOptimized(userId: string) {
    const credentials = this.getGhlCredentials();
    if (!credentials) {
      this.logger.warn('GHL credentials not configured - returning empty response');
      return { opportunities: [], total: 0 };
    }

    try {
      // Get user details
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      this.logger.log(`🚀 OPTIMIZED APPOINTMENT MATCHING: Starting analysis for user: ${user.name} (${user.role})`);

      // Get opportunities from AI and Manual stages
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      this.logger.log(`📊 Found ${allOpportunities.length} opportunities in survey stages`);

      // Filter opportunities based on user role
      const filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        allOpportunities,
        user,
        credentials
      );

      this.logger.log(`👤 Filtered to ${filteredOpportunities.length} opportunities for user ${user.name}`);

      // BATCH PROCESSING: Get all unique contact IDs
      const contactIds = [...new Set(filteredOpportunities
        .map(opp => opp.contactId || opp.contact?.id)
        .filter(id => id))];

      this.logger.log(`📋 Processing ${contactIds.length} unique contacts for appointments`);

      // BATCH PROCESSING: Get all appointments in parallel
      const appointmentPromises = contactIds.map(async (contactId) => {
        try {
          const appointments = await this.goHighLevelService.getAppointmentsByContactId(
            credentials.accessToken,
            credentials.locationId,
            contactId
          );
          return { contactId, appointments };
        } catch (error) {
          this.logger.warn(`Failed to get appointments for contact ${contactId}: ${error.message}`);
          return { contactId, appointments: [] };
        }
      });

      const appointmentResults = await Promise.all(appointmentPromises);
      
      // Create a map for quick lookup
      const appointmentMap = new Map(
        appointmentResults.map(result => [result.contactId, result.appointments])
      );

      this.logger.log(`✅ Retrieved appointments for ${appointmentResults.length} contacts`);

      // Process opportunities with cached appointment data
      const opportunitiesWithAnalysis = filteredOpportunities.map(opportunity => {
        const contactId = opportunity.contactId || opportunity.contact?.id;
        const appointments = contactId ? (appointmentMap.get(contactId) || []) : [];
        
        const analysis = this.analyzeAppointmentsForOpportunity(appointments, opportunity);
        
        return {
          ...opportunity,
          appointmentAnalysis: analysis
        };
      });

      // Classify opportunities based on appointment status
      const classifiedOpportunities = this.classifyOpportunities(opportunitiesWithAnalysis);

      // Log classification summary
      this.logger.log(`📈 OPTIMIZED APPOINTMENT CLASSIFICATION SUMMARY:`);
      this.logger.log(`✅ Confirmed with real appointments: ${classifiedOpportunities.confirmedWithAppointments.length}`);
      this.logger.log(`⚠️ Tagged as booked but no appointment: ${classifiedOpportunities.taggedButNoAppointment.length}`);
      this.logger.log(`❓ Multiple appointments (low confidence): ${classifiedOpportunities.multipleAppointments.length}`);
      this.logger.log(`❌ No appointments found: ${classifiedOpportunities.noAppointments.length}`);

      return {
        opportunities: classifiedOpportunities.all,
        total: classifiedOpportunities.all.length,
        classification: {
          confirmedWithAppointments: classifiedOpportunities.confirmedWithAppointments.length,
          taggedButNoAppointment: classifiedOpportunities.taggedButNoAppointment.length,
          multipleAppointments: classifiedOpportunities.multipleAppointments.length,
          noAppointments: classifiedOpportunities.noAppointments.length,
        },
        user: {
          id: user.id,
          name: user.name,
          role: user.role,
        },
        performance: {
          totalContacts: contactIds.length,
          totalAppointments: appointmentResults.reduce((sum, result) => sum + result.appointments.length, 0),
          processingTime: 'optimized'
        }
      };
    } catch (error) {
      this.logger.error(`Error in getOpportunitiesWithAppointmentsOptimized: ${error.message}`);
      throw error;
    }
  }

  // NEW: Hybrid method that uses fast dashboard data + appointment info
  async getOpportunitiesWithAppointmentsHybrid(userId: string) {
    const credentials = this.getGhlCredentials();
    if (!credentials) {
      this.logger.warn('GHL credentials not configured - returning empty response');
      return { opportunities: [], total: 0 };
    }

    try {
      // Get user details
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      this.logger.log(`🚀 HYBRID APPROACH: Starting fast dashboard + appointments for user: ${user.name}`);

      // STEP 1: Get fast dashboard data (this is what makes dashboard load in 10 seconds)
      const dashboardData = await this.getAiVsManualOpportunities(userId);
      
      // Combine all opportunities from dashboard
      const allOpportunities = [
        ...(dashboardData.ai?.opportunities || []),
        ...(dashboardData.manual?.opportunities || [])
      ];

      this.logger.log(`📊 Fast dashboard data loaded: ${allOpportunities.length} opportunities`);

      // STEP 2: Get unique contact IDs for appointment fetching
      const contactIds = [...new Set(allOpportunities
        .map(opp => opp.contactId || opp.contact?.id)
        .filter(id => id))];

      this.logger.log(`📋 Processing ${contactIds.length} unique contacts for appointments`);

      // STEP 3: Process opportunities with contact notes appointment detection (no GHL API appointments)
      const opportunitiesWithAppointments = await this.processOpportunitiesWithHybridAppointments(
        allOpportunities,
        credentials
      );

      this.logger.log(`✅ Processed ${opportunitiesWithAppointments.length} opportunities with contact notes appointment detection`);

      // STEP 4: Classify opportunities (simplified since we're only using contact notes)
      const confirmedWithAppointments = opportunitiesWithAppointments.filter(opp => opp.hasAppointment);
      const noAppointments = opportunitiesWithAppointments.filter(opp => !opp.hasAppointment);
      
      const classifiedOpportunities = {
        confirmedWithAppointments,
        taggedButNoAppointment: [],
        multipleAppointments: [], // No multiple appointments since we only use contact notes
        noAppointments,
        all: opportunitiesWithAppointments
      };

      this.logger.log(`📈 HYBRID APPOINTMENT CLASSIFICATION:`);
      this.logger.log(`✅ Confirmed with appointments: ${classifiedOpportunities.confirmedWithAppointments.length}`);
      this.logger.log(`⚠️ Tagged but no appointment: ${classifiedOpportunities.taggedButNoAppointment.length}`);
      this.logger.log(`❓ Multiple appointments: ${classifiedOpportunities.multipleAppointments.length}`);
      this.logger.log(`❌ No appointments: ${classifiedOpportunities.noAppointments.length}`);

      return {
        opportunities: classifiedOpportunities.all,
        total: classifiedOpportunities.all.length,
        classification: {
          confirmedWithAppointments: classifiedOpportunities.confirmedWithAppointments.length,
          taggedButNoAppointment: classifiedOpportunities.taggedButNoAppointment.length,
          multipleAppointments: classifiedOpportunities.multipleAppointments.length,
          noAppointments: classifiedOpportunities.noAppointments.length,
        },
        user: {
          id: user.id,
          name: user.name,
          role: user.role,
        },
        performance: {
          method: 'hybrid',
          dashboardDataTime: 'fast',
          appointmentProcessingTime: 'optimized',
          totalContacts: contactIds.length,
          totalAppointments: opportunitiesWithAppointments.filter(opp => opp.hasAppointment).length
        }
      };
    } catch (error) {
      this.logger.error(`Error in getOpportunitiesWithAppointmentsHybrid: ${error.message}`);
      throw error;
    }
  }

  private async analyzeOpportunityAppointments(
    opportunities: any[],
    credentials: { accessToken: string; locationId: string }
  ): Promise<any[]> {
    const opportunitiesWithAnalysis: any[] = [];

    for (const opportunity of opportunities) {
      const contactId = opportunity.contactId || opportunity.contact?.id;
      
      if (!contactId) {
        opportunitiesWithAnalysis.push({
          ...opportunity,
          appointmentAnalysis: {
            status: 'NO_CONTACT_ID',
            reason: 'Opportunity has no contact ID',
            appointments: [],
            validAppointments: [],
            appointmentCount: 0,
            hasValidAppointment: false,
            confidence: 'LOW'
          }
        });
        continue;
      }

      try {
        // Get all appointments for this contact
        const allAppointments = await this.goHighLevelService.getAppointmentsByContactId(
          credentials.accessToken,
          credentials.locationId,
          contactId
        );

        // If no appointments found, try alternative method
        let contactAppointments = allAppointments;
        if (allAppointments.length === 0) {
          contactAppointments = await this.goHighLevelService.getAppointmentsByContactIdAlternative(
            credentials.accessToken,
            credentials.locationId,
            contactId
          );
        }

        // Analyze appointments for this opportunity
        const analysis = this.analyzeAppointmentsForOpportunity(
          contactAppointments,
          opportunity
        );

        opportunitiesWithAnalysis.push({
          ...opportunity,
          appointmentAnalysis: analysis
        });

        this.logger.log(`📋 Opportunity ${opportunity.id} (${opportunity.name}): ${analysis.status} - ${analysis.reason}`);

      } catch (error) {
        this.logger.warn(`❌ Failed to analyze appointments for opportunity ${opportunity.id}: ${error.message}`);
        opportunitiesWithAnalysis.push({
          ...opportunity,
          appointmentAnalysis: {
            status: 'ERROR',
            reason: `Failed to fetch appointments: ${error.message}`,
            appointments: [],
            validAppointments: [],
            appointmentCount: 0,
            hasValidAppointment: false,
            confidence: 'LOW'
          }
        });
      }
    }

    return opportunitiesWithAnalysis;
  }

  private analyzeAppointmentsForOpportunity(appointments: any[], opportunity: any) {
    const now = new Date();
    const opportunityCreatedAt = new Date(opportunity.createdAt || opportunity.created_at);
    
    // Filter for valid appointments (booked, confirmed, include past appointments for conversion tracking)
    const validAppointments = appointments.filter(appointment => {
      const title = appointment.title?.toLowerCase() || '';
      const status = appointment.status?.toLowerCase() || appointment.appoinmentStatus?.toLowerCase() || '';
      const startTime = appointment.startTime || appointment.start_time;
      
      // Skip irrelevant appointments
      if (title.includes('busy') || 
          status.includes('unavailable') || 
          title.length === 0) {
        return false;
      }

      // For dashboard conversion tracking, include all appointments (past and future)
      // Only skip if the appointment date is clearly invalid or very old (more than 2 years ago)
      if (startTime) {
        const appointmentDate = new Date(startTime);
        const twoYearsAgo = new Date(now.getTime() - (2 * 365 * 24 * 60 * 60 * 1000));
        if (appointmentDate < twoYearsAgo) {
          return false; // Skip very old appointments (more than 2 years ago)
        }
      }

      // Check for valid status (include cancelled appointments for historical tracking)
      const isValidStatus = status.includes('booked') || 
                           status.includes('confirmed') || 
                           status.includes('scheduled') ||
                           status.includes('cancelled'); // Include cancelled for historical tracking

      return isValidStatus;
    });

    // Sort valid appointments by date (closest to opportunity creation first)
    validAppointments.sort((a, b) => {
      const aDate = new Date(a.startTime || a.start_time);
      const bDate = new Date(b.startTime || b.start_time);
      const aDiff = Math.abs(aDate.getTime() - opportunityCreatedAt.getTime());
      const bDiff = Math.abs(bDate.getTime() - opportunityCreatedAt.getTime());
      return aDiff - bDiff;
    });

    // Determine status and confidence
    let status = 'NO_APPOINTMENTS';
    let reason = 'No valid appointments found';
    let confidence = 'LOW';

    if (validAppointments.length === 0) {
      if (appointments.length > 0) {
        status = 'INVALID_APPOINTMENTS';
        reason = `Found ${appointments.length} appointments but none are valid (past dates, cancelled, or invalid status)`;
      } else {
        status = 'NO_APPOINTMENTS';
        reason = 'No appointments found for this contact';
      }
    } else if (validAppointments.length === 1) {
      status = 'CONFIRMED_WITH_APPOINTMENT';
      reason = `Found 1 valid appointment: ${validAppointments[0].title}`;
      confidence = 'HIGH';
    } else {
      status = 'MULTIPLE_APPOINTMENTS';
      reason = `Found ${validAppointments.length} valid appointments - unclear which one is for this opportunity`;
      confidence = 'MEDIUM';
    }

    return {
      status,
      reason,
      appointments: appointments,
      validAppointments: validAppointments,
      appointmentCount: appointments.length,
      validAppointmentCount: validAppointments.length,
      hasValidAppointment: validAppointments.length > 0,
      confidence,
      primaryAppointment: validAppointments[0] || null
    };
  }

  private classifyOpportunities(opportunitiesWithAnalysis: any[]) {
    const confirmedWithAppointments: any[] = [];
    const taggedButNoAppointment: any[] = [];
    const multipleAppointments: any[] = [];
    const noAppointments: any[] = [];

    for (const opportunity of opportunitiesWithAnalysis) {
      const analysis = opportunity.appointmentAnalysis;
      
      switch (analysis.status) {
        case 'CONFIRMED_WITH_APPOINTMENT':
          confirmedWithAppointments.push({
            ...opportunity,
            hasAppointment: true,
            appointmentDetails: analysis.primaryAppointment ? {
              id: analysis.primaryAppointment.id,
              title: analysis.primaryAppointment.title,
              date: this.formatDateForAPI(analysis.primaryAppointment.startTime || analysis.primaryAppointment.start_time),
              status: analysis.primaryAppointment.status,
              confidence: analysis.confidence
            } : null,
            appointmentCount: analysis.validAppointmentCount,
            classification: 'CONFIRMED',
            type: opportunity.pipelineStageId === '8904bbe1-53a3-468e-94e4-f13cb04a4947' ? 'ai' : 'manual'
          });
          break;

        case 'MULTIPLE_APPOINTMENTS':
          multipleAppointments.push({
            ...opportunity,
            hasAppointment: true,
            appointmentDetails: analysis.primaryAppointment ? {
              id: analysis.primaryAppointment.id,
              title: analysis.primaryAppointment.title,
              date: this.formatDateForAPI(analysis.primaryAppointment.startTime || analysis.primaryAppointment.start_time),
              status: analysis.primaryAppointment.status,
              confidence: analysis.confidence,
              note: `Multiple appointments found (${analysis.validAppointmentCount})`
            } : null,
            appointmentCount: analysis.validAppointmentCount,
            classification: 'MULTIPLE',
            type: opportunity.pipelineStageId === '8904bbe1-53a3-468e-94e4-f13cb04a4947' ? 'ai' : 'manual'
          });
          break;

        case 'NO_APPOINTMENTS':
        case 'INVALID_APPOINTMENTS':
        case 'NO_CONTACT_ID':
        case 'ERROR':
          noAppointments.push({
            ...opportunity,
            hasAppointment: false,
            appointmentDetails: null,
            appointmentCount: 0,
            classification: 'NO_APPOINTMENT',
            reason: analysis.reason,
            type: opportunity.pipelineStageId === '8904bbe1-53a3-468e-94e4-f13cb04a4947' ? 'ai' : 'manual'
          });
          break;
      }
    }

    // Combine all opportunities
    const all = [
      ...confirmedWithAppointments,
      ...multipleAppointments,
      ...noAppointments
    ];

    return {
      confirmedWithAppointments,
      taggedButNoAppointment,
      multipleAppointments,
      noAppointments,
      all
    };
  }

  private extractTokenData(accessToken: string) {
    try {
      const payload = JSON.parse(Buffer.from(accessToken.split('.')[1], 'base64').toString());
      return {
        locationId: payload.location_id || payload.locationId,
        userId: payload.sub || payload.userId,
        companyId: payload.companyId
      };
    } catch (error) {
      this.logger.error('Error extracting token data:', error);
      throw new Error('Invalid access token format');
    }
  }

  private async enhanceOpportunitiesWithAppointmentInfo(
    opportunities: any[],
    credentials: { accessToken: string; locationId: string }
  ): Promise<any[]> {
    // Process opportunities in parallel for better performance (especially for admins)
    const enhancedPromises = opportunities.map(async (opportunity) => {
      const contactId = opportunity.contactId || opportunity.contact?.id;
      
      let appointmentInfo: any = {
        hasAppointment: false,
        appointmentCount: 0,
        appointmentDetails: null,
        classification: 'NO_APPOINTMENT',
        confidence: 'LOW',
        reason: 'No contact ID available'
      };

      if (contactId) {
        try {
          // ONLY USE CONTACT NOTES: Skip GHL API appointments to avoid "Multiple" classification
          try {
            const contactNotes = await this.goHighLevelService.getContactNotes(
              credentials.accessToken,
              contactId
            );
              
            const appointmentFromNotes = this.extractAppointmentDetailsFromNotes(contactNotes);
              
            if (appointmentFromNotes) {
              appointmentInfo = {
                hasAppointment: true,
                appointmentCount: 1,
                appointmentDetails: appointmentFromNotes,
                classification: 'CONFIRMED',
                confidence: 'MEDIUM',
                reason: 'Appointment found in contact notes',
                appointmentSource: 'contact_notes'
              };
            } else {
              appointmentInfo = {
                hasAppointment: false,
                appointmentCount: 0,
                appointmentDetails: null,
                classification: 'NO_APPOINTMENT',
                confidence: 'LOW',
                reason: 'No appointment details found in contact notes'
              };
            }
          } catch (notesError) {
            this.logger.warn(`❌ Failed to get contact notes for ${opportunity.name}: ${notesError.message}`);
            appointmentInfo = {
              hasAppointment: false,
              appointmentCount: 0,
              appointmentDetails: null,
              classification: 'NO_APPOINTMENT',
              confidence: 'LOW',
              reason: 'Failed to get contact notes'
            };
          }

        } catch (error) {
          this.logger.warn(`❌ Failed to get appointments for opportunity ${opportunity.id}: ${error.message}`);
          appointmentInfo = {
            hasAppointment: false,
            appointmentCount: 0,
            appointmentDetails: null,
            classification: 'ERROR',
            confidence: 'LOW',
            reason: `Failed to fetch appointments: ${error.message}`
          };
        }
      }

      return {
        ...opportunity,
        ...appointmentInfo,
        // Add type field based on pipeline stage
        type: opportunity.pipelineStageId === '8904bbe1-53a3-468e-94e4-f13cb04a4947' ? 'ai' : 'manual'
      };
    });

    // Wait for all parallel requests to complete
    return Promise.all(enhancedPromises);
  }

  private quickAppointmentAnalysis(appointments: any[], opportunity: any) {
    const now = new Date();
    
    // Filter for valid appointments (booked, confirmed, future dates)
    const validAppointments = appointments.filter(appointment => {
      const title = appointment.title?.toLowerCase() || '';
      const status = appointment.status?.toLowerCase() || appointment.appoinmentStatus?.toLowerCase() || '';
      const startTime = appointment.startTime || appointment.start_time;
      
      // Skip irrelevant appointments
      if (title.includes('busy') || 
          status.includes('unavailable') || 
          status.includes('cancelled') ||
          title.length === 0) {
        return false;
      }

      // Check if appointment is in the future
      if (startTime) {
        const appointmentDate = new Date(startTime);
        if (appointmentDate <= now) {
          return false; // Skip past appointments
        }
      }

      // Check for valid status
      const isValidStatus = status.includes('booked') || 
                           status.includes('confirmed') || 
                           status.includes('scheduled');

      return isValidStatus;
    });

    // Determine status
    let classification = 'NO_APPOINTMENT';
    let confidence = 'LOW';
    let reason = 'No valid appointments found';

    if (validAppointments.length === 0) {
      if (appointments.length > 0) {
        classification = 'INVALID_APPOINTMENTS';
        reason = `Found ${appointments.length} appointments but none are valid`;
      } else {
        classification = 'NO_APPOINTMENTS';
        reason = 'No appointments found for this contact';
      }
    } else if (validAppointments.length === 1) {
      classification = 'CONFIRMED';
      reason = `Found 1 valid appointment: ${validAppointments[0].title}`;
      confidence = 'HIGH';
    } else {
      classification = 'MULTIPLE';
      reason = `Found ${validAppointments.length} valid appointments`;
      confidence = 'MEDIUM';
    }

    return {
      hasAppointment: validAppointments.length > 0,
      appointmentCount: appointments.length,
      validAppointmentCount: validAppointments.length,
      appointmentDetails: validAppointments[0] ? {
        id: validAppointments[0].id,
        title: validAppointments[0].title,
        date: this.formatDateForAPI(validAppointments[0].startTime || validAppointments[0].start_time),
        status: validAppointments[0].status,
        confidence: confidence
      } : null,
      classification,
      confidence,
      reason
    };
  }

  /**
   * Fetch contact address and notes for a single opportunity (on-demand)
   * This is called when viewing opportunity details
   */
  async getOpportunityDetails(opportunityId: string, userId: string): Promise<{ 
    contactAddress: string | null; 
    contactPostcode: string | null; 
    address: string | null;
    contactFirstName: string | null;
    contactLastName: string | null;
    contactCity: string | null;
    contactState: string | null;
    contactAddressLine2: string | null;
    notes: string | null;
    customFields: any[] | null;
    appointmentDetails: any | null;
  }> {
    this.logger.log(`🔍 Fetching address for opportunity: ${opportunityId}`);
    
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      if (!credentials) {
        throw new Error('GHL credentials not configured');
      }

      // Try to get the opportunity directly first (more efficient)
      this.logger.log(`🔍 Attempting direct opportunity fetch: ${opportunityId}`);
      let opportunity: any = null;
      let fetchMethod = 'direct';
      
      try {
        opportunity = await this.goHighLevelService.getOpportunityById(credentials.accessToken, opportunityId);
        this.logger.log(`✅ Found opportunity directly: ${opportunity?.name} (ID: ${opportunity?.id})`);
      } catch (directError: any) {
        // Only fallback for specific errors, not all errors
        if (directError.response?.status === 404 || directError.message?.includes('not found')) {
          this.logger.warn(`⚠️ Opportunity not found via direct fetch, trying search method`);
          fetchMethod = 'search';
          
          // Fallback: search through all opportunities (but with better caching)
          this.logger.log(`🔍 Searching through all opportunities for: ${opportunityId}`);
          const allOpportunities = await this.goHighLevelService.getOpportunitiesWithStageNames(
            credentials.accessToken,
            credentials.locationId
          );
          
          this.logger.log(`🔍 Found ${allOpportunities.length} total opportunities`);
          
          opportunity = allOpportunities.find((opp: any) => opp.id === opportunityId);
        } else {
          // For other errors (rate limiting, network issues), throw the error
          this.logger.error(`❌ Direct fetch failed with error: ${directError.message}`);
          throw directError;
        }
      }
      
      if (!opportunity) {
        this.logger.warn(`⚠️ Opportunity not found: ${opportunityId} (tried ${fetchMethod} method) - returning empty data`);
        return {
          contactAddress: null,
          contactPostcode: null,
          address: null,
          contactFirstName: null,
          contactLastName: null,
          contactCity: null,
          contactState: null,
          contactAddressLine2: null,
          notes: null,
          customFields: null,
          appointmentDetails: null
        };
      }

      this.logger.log(`🔍 Found opportunity via ${fetchMethod} method: ${opportunity.name} (ID: ${opportunity.id})`);
      this.logger.log(`🔍 Opportunity structure:`, JSON.stringify(opportunity, null, 2));

      const contactId = opportunity.contactId || opportunity.contact?.id;
      this.logger.log(`🔍 Contact ID extracted: ${contactId}`);
      this.logger.log(`🔍 opportunity.contactId: ${opportunity.contactId}`);
      this.logger.log(`🔍 opportunity.contact?.id: ${opportunity.contact?.id}`);
      
      if (!contactId) {
        this.logger.warn(`❌ No contact ID found for opportunity: ${opportunityId}`);
        return {
          contactAddress: null,
          contactPostcode: null,
          address: null,
          contactFirstName: null,
          contactLastName: null,
          contactCity: null,
          contactState: null,
          contactAddressLine2: null,
          notes: null,
          customFields: null,
          appointmentDetails: null
        };
      }

      // Fetch contact details from GoHighLevel
      this.logger.log(`🔍 Making API call to /v1/contacts/${contactId}`);
      const contactDetails = await this.goHighLevelService.getContactById(credentials.accessToken, contactId);
      
      this.logger.log(`🔍 Contact details API response:`, JSON.stringify(contactDetails, null, 2));
      
      if (contactDetails && contactDetails.contact) {
        const contact = contactDetails.contact;
        this.logger.log(`🔍 Contact object structure:`, JSON.stringify(contact, null, 2));
        
        // Extract contact information from contact
        const firstName = contact.firstName || null;
        const lastName = contact.lastName || null;
        const address1 = contact.address1 || contact.addresses?.[0]?.address1;
        const address2 = contact.address2 || contact.addresses?.[0]?.address2;
        const city = contact.city || contact.addresses?.[0]?.city;
        const state = contact.state || contact.addresses?.[0]?.state;
        const postalCode = contact.postalCode || contact.addresses?.[0]?.postalCode;
        
        this.logger.log(`🔍 Extracted contact fields:`);
        this.logger.log(`   - firstName: "${firstName}"`);
        this.logger.log(`   - lastName: "${lastName}"`);
        this.logger.log(`   - address1: "${address1}"`);
        this.logger.log(`   - address2: "${address2}"`);
        this.logger.log(`   - city: "${city}"`);
        this.logger.log(`   - state: "${state}"`);
        this.logger.log(`   - postalCode: "${postalCode}"`);
        
        // Create full address string
        const fullAddress = [address1, city, state].filter(Boolean).join(', ');
        this.logger.log(`🔍 Full address constructed: "${fullAddress}"`);
        
        // Extract custom fields and notes
        const rawCustomFields = contact.customField || [];
        const notes = contact.notes || null;
        
        // Filter and deduplicate custom fields
        const customFields = this.filterAndDeduplicateCustomFields(rawCustomFields);
        
        // Extract appointment details from contact notes
        let appointmentDetails = null;
        try {
          const contactNotes = await this.goHighLevelService.getContactNotes(
            credentials.accessToken,
            contactId
          );
          appointmentDetails = this.extractAppointmentDetailsFromNotes(contactNotes);
          this.logger.log(`🔍 Appointment details extracted: ${appointmentDetails ? 'found' : 'none'}`);
        } catch (appointmentError) {
          this.logger.warn(`❌ Failed to get appointment details for contact ${contactId}: ${appointmentError.message}`);
        }
        
        this.logger.log(`🔍 Raw custom fields: ${rawCustomFields.length} fields`);
        this.logger.log(`🔍 Filtered custom fields: ${customFields.length} fields`);
        this.logger.log(`🔍 Extracted notes: ${notes ? 'present' : 'none'}`);
        
        this.logger.log(`✅ Details fetched for opportunity ${opportunityId}: ${firstName} ${lastName}, ${fullAddress || 'No address'}, ${customFields.length} custom fields`);
        
        return {
          contactAddress: fullAddress || null,
          contactPostcode: postalCode || null,
          address: fullAddress || null,
          contactFirstName: firstName,
          contactLastName: lastName,
          contactCity: city || null,
          contactState: state || null,
          contactAddressLine2: address2 || null,
          notes: notes,
          customFields: customFields,
          appointmentDetails: appointmentDetails
        };
      } else {
        this.logger.warn(`❌ No contact details found for contact ID: ${contactId}`);
        this.logger.log(`🔍 Contact details response:`, contactDetails);
        return {
          contactAddress: null,
          contactPostcode: null,
          address: null,
          contactFirstName: null,
          contactLastName: null,
          contactCity: null,
          contactState: null,
          contactAddressLine2: null,
          notes: null,
          customFields: null,
          appointmentDetails: null
        };
      }
      
    } catch (error) {
      this.logger.error(`❌ Error fetching address for opportunity ${opportunityId}: ${error.message}`);
      this.logger.error(`❌ Error stack:`, error.stack);
      throw error;
    }
  }

  /**
   * Fetch customer details including email for welcome email functionality
   */
  async getCustomerDetails(opportunityId: string, userId: string): Promise<{
    name: string | null;
    email: string | null;
    address: string | null;
    postcode: string | null;
    phone: string | null;
  }> {
    this.logger.log(`🔍 Fetching customer details for opportunity: ${opportunityId}`);
    
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      if (!credentials) {
        throw new Error('GHL credentials not configured');
      }

      // Get the opportunity details
      this.logger.log(`🔍 Attempting to fetch opportunity: ${opportunityId}`);
      let opportunity: any = null;
      
      try {
        opportunity = await this.goHighLevelService.getOpportunityById(credentials.accessToken, opportunityId);
        this.logger.log(`✅ Found opportunity: ${opportunity?.name} (ID: ${opportunity?.id})`);
      } catch (directError: any) {
        if (directError.response?.status === 404 || directError.message?.includes('not found')) {
          this.logger.warn(`⚠️ Opportunity not found via direct fetch, trying search method`);
          
          const allOpportunities = await this.goHighLevelService.getOpportunitiesWithStageNames(
            credentials.accessToken,
            credentials.locationId
          );
          
          opportunity = allOpportunities.find((opp: any) => opp.id === opportunityId);
        } else {
          this.logger.error(`❌ Direct fetch failed: ${directError.message}`);
          throw directError;
        }
      }
      
      if (!opportunity) {
        this.logger.warn(`⚠️ Opportunity not found: ${opportunityId}`);
        return {
          name: null,
          email: null,
          address: null,
          postcode: null,
          phone: null
        };
      }

      this.logger.log(`🔍 Found opportunity: ${opportunity.name}`);

      const contactId = opportunity.contactId || opportunity.contact?.id;
      this.logger.log(`🔍 Contact ID: ${contactId}`);
      
      if (!contactId) {
        this.logger.warn(`❌ No contact ID found for opportunity: ${opportunityId}`);
        return {
          name: opportunity.name || null,
          email: null,
          address: null,
          postcode: null,
          phone: null
        };
      }

      // Fetch contact details
      this.logger.log(`🔍 Fetching contact details for: ${contactId}`);
      const contactDetails = await this.goHighLevelService.getContactById(credentials.accessToken, contactId);
      
      if (contactDetails && contactDetails.contact) {
        const contact = contactDetails.contact;
        
        // Extract customer name with priority order
        let customerName = '';
        if (contact.firstName && contact.lastName) {
          customerName = `${contact.firstName.trim()} ${contact.lastName.trim()}`;
        } else if (contact.name && contact.name.trim() !== '') {
          customerName = contact.name.trim();
        } else if (opportunity.name && opportunity.name.trim() !== '') {
          customerName = opportunity.name.trim();
        } else if (contact.email) {
          const emailPrefix = contact.email.split('@')[0];
          customerName = emailPrefix.replace(/[._-]/g, ' ').replace(/\b\w/g, (l: string) => l.toUpperCase());
        } else {
          customerName = `Customer ${opportunityId.slice(-6)}`;
        }

        // Clean up name (remove postcode patterns, etc.)
        if (customerName) {
          const postcodePattern = /^[A-Z]{1,2}\d{1,2}\s?\d[A-Z]{2},\s*/i;
          customerName = customerName.replace(postcodePattern, '');
          
          const emailPattern = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,},\s*/;
          customerName = customerName.replace(emailPattern, '');
          
          customerName = customerName.replace(/^,\s*/, '').replace(/\s*,$/, '').trim();
        }

        // Extract address information
        const primaryAddress = contact.address1 || contact.addresses?.[0]?.address1 || null;
        const city = contact.city || contact.addresses?.[0]?.city || null;
        const postalCode = contact.postalCode || contact.addresses?.[0]?.postalCode || null;
        
        let fullAddress: string | null = null;
        if (primaryAddress) {
          fullAddress = primaryAddress;
          if (city) {
            fullAddress = fullAddress + `, ${city}`;
          }
        }

        this.logger.log(`✅ Customer details extracted:`, {
          name: customerName,
          email: contact.email || null,
          address: fullAddress,
          postcode: postalCode,
          phone: contact.phone || null
        });
        
        return {
          name: customerName || null,
          email: contact.email || null,
          address: fullAddress,
          postcode: postalCode,
          phone: contact.phone || null
        };
      } else {
        this.logger.warn(`❌ No contact details found for contact ID: ${contactId}`);
        return {
          name: opportunity.name || null,
          email: null,
          address: null,
          postcode: null,
          phone: null
        };
      }
      
    } catch (error) {
      this.logger.error(`❌ Error fetching customer details for opportunity ${opportunityId}: ${error.message}`);
      this.logger.error(`❌ Error stack:`, error.stack);
      throw error;
    }
  }

  /**
   * Enhance opportunities with contact location data - EFFICIENT VERSION
   * Uses caching and proper rate limiting to avoid API quota issues
   */
  private async enhanceOpportunitiesWithContactDataEfficient(opportunities: any[], accessToken: string): Promise<any[]> {
    this.logger.log(`Enhancing ${opportunities.length} opportunities with contact location data (EFFICIENT)`);
    
    const enhancedOpportunities: any[] = [];
    const contactCache = new Map<string, any>(); // Cache contact data
    const batchSize = 1; // Process one at a time to avoid rate limiting
    const delayBetweenCalls = 1000; // 1 second delay between API calls
    
    // First pass: extract postcodes from opportunity names
    const opportunitiesWithPostcodes = opportunities.map(opp => {
      let contactPostcode: string | null = null;
      
      // Extract postcode from opportunity name (e.g., "N12 9JA, Lisa Jones")
      if (opp.name) {
        const postcodeMatch = opp.name.match(/[A-Z]{1,2}\d{1,2}\s?\d[A-Z]{2}/i);
        if (postcodeMatch) {
          contactPostcode = postcodeMatch[0].toUpperCase();
        }
      }
      
      return {
        ...opp,
        contactPostcode,
        contactAddress: null,
        address: null
      };
    });
    
    // Second pass: fetch contact details for addresses
    for (let i = 0; i < opportunitiesWithPostcodes.length; i++) {
      const opportunity = opportunitiesWithPostcodes[i];
      const contactId = opportunity.contactId || opportunity.contact?.id;
      
      if (contactId && !contactCache.has(contactId)) {
        try {
          this.logger.log(`🔍 Fetching contact details for ${opportunity.name} (${i + 1}/${opportunities.length})`);
          
          // Fetch contact details from GoHighLevel
          const contactDetails = await this.goHighLevelService.getContactById(accessToken, contactId);
          
          if (contactDetails && contactDetails.contact) {
            contactCache.set(contactId, contactDetails.contact);
          }
          
          // Add delay to avoid rate limiting
          if (i < opportunitiesWithPostcodes.length - 1) {
            await new Promise(resolve => setTimeout(resolve, delayBetweenCalls));
          }
          
        } catch (error) {
          this.logger.warn(`❌ Failed to fetch contact ${contactId}: ${error.message}`);
        }
      }
    }
    
    // Third pass: enhance opportunities with contact data
    for (const opportunity of opportunitiesWithPostcodes) {
      const contactId = opportunity.contactId || opportunity.contact?.id;
      const cachedContact = contactCache.get(contactId);
      
      if (cachedContact) {
        // Extract address information from cached contact
        const address1 = cachedContact.address1 || cachedContact.addresses?.[0]?.address1;
        const city = cachedContact.city || cachedContact.addresses?.[0]?.city;
        const state = cachedContact.state || cachedContact.addresses?.[0]?.state;
        const postalCode = cachedContact.postalCode || cachedContact.addresses?.[0]?.postalCode;
        
        // Create full address string
        const fullAddress = [address1, city, state].filter(Boolean).join(', ');
        
        enhancedOpportunities.push({
          ...opportunity,
          contactAddress: fullAddress || null,
          contactPostcode: postalCode || opportunity.contactPostcode,
          address: fullAddress || null
        });
      } else {
        // Use postcode as fallback if no contact details
        enhancedOpportunities.push({
          ...opportunity,
          contactAddress: opportunity.contactPostcode ? `Address: ${opportunity.contactPostcode}` : null,
          address: opportunity.contactPostcode || null
        });
      }
    }
    
    this.logger.log(`✅ Successfully enhanced ${enhancedOpportunities.length} opportunities with contact location data`);
    this.logger.log(`📊 Contact cache hits: ${contactCache.size} unique contacts fetched`);
    
    return enhancedOpportunities;
  }

  /**
   * Enhance opportunities with contact location data - OPTIMIZED VERSION
   */
  private async enhanceOpportunitiesWithContactData(opportunities: any[], accessToken: string): Promise<any[]> {
    this.logger.log(`Enhancing ${opportunities.length} opportunities with contact location data (OPTIMIZED)`);
    
    const enhancedOpportunities: any[] = [];
    let enhancedCount = 0;
    let noContactIdCount = 0;
    let errorCount = 0;
    
    // Process opportunities in batches to avoid overwhelming the API
    const batchSize = 2; // Reduced batch size to avoid rate limiting
    const batches: any[][] = [];
    
    for (let i = 0; i < opportunities.length; i += batchSize) {
      batches.push(opportunities.slice(i, i + batchSize));
    }
    
    this.logger.log(`Processing ${opportunities.length} opportunities in ${batches.length} batches of ${batchSize}`);
    
    for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
      const batch = batches[batchIndex];
      this.logger.log(`Processing batch ${batchIndex + 1}/${batches.length} with ${batch.length} opportunities`);
      
      // Process batch sequentially to avoid rate limiting
      for (const opportunity of batch) {
        try {
          // Get contact ID from opportunity
          const contactId = opportunity.contactId || opportunity.contact?.id;
          
          if (contactId) {
            // Fetch contact details from GoHighLevel
            const contactDetails = await this.goHighLevelService.getContactById(accessToken, contactId);
            
            if (contactDetails && contactDetails.contact) {
              // Extract address information from contact
              const contact = contactDetails.contact;
              const address1 = contact.address1 || contact.addresses?.[0]?.address1;
              const city = contact.city || contact.addresses?.[0]?.city;
              const state = contact.state || contact.addresses?.[0]?.state;
              const postalCode = contact.postalCode || contact.addresses?.[0]?.postalCode;
              
              // Create full address string
              const fullAddress = [address1, city, state].filter(Boolean).join(', ');
              
              // Enhance opportunity with location data
              enhancedOpportunities.push({
                ...opportunity,
                contactAddress: fullAddress || null,
                contactPostcode: postalCode || null,
                address: fullAddress || null // Fallback for general address field
              });
            } else {
              // If contact details not found, keep original opportunity
              enhancedOpportunities.push({
                ...opportunity,
                contactAddress: null,
                contactPostcode: null,
                address: null
              });
            }
          } else {
            // If no contact ID, keep original opportunity
            enhancedOpportunities.push({
              ...opportunity,
              contactAddress: null,
              contactPostcode: null,
              address: null
            });
          }
          
          // Add delay between each API call to avoid rate limiting
          await new Promise(resolve => setTimeout(resolve, 200));
          
        } catch (error) {
          this.logger.warn(`❌ Failed to enhance opportunity ${opportunity.id} with contact data: ${error.message}`);
          // Keep original opportunity if enhancement fails
          enhancedOpportunities.push({
            ...opportunity,
            contactAddress: null,
            contactPostcode: null,
            address: null
          });
        }
      }
      
      // Add longer delay between batches to avoid rate limiting
      if (batchIndex < batches.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }
    
    // Count results
    enhancedCount = enhancedOpportunities.filter(opp => opp.contactAddress || opp.contactPostcode).length;
    noContactIdCount = enhancedOpportunities.filter(opp => !opp.contactId && !opp.contact?.id).length;
    errorCount = opportunities.length - enhancedCount - noContactIdCount;
    
    this.logger.log(`✅ Successfully enhanced ${enhancedOpportunities.length} opportunities with contact location data`);
    this.logger.log(`📊 Enhancement stats: ${enhancedCount} enhanced, ${noContactIdCount} no contact ID, ${errorCount} errors`);
    
    // Log a sample of enhanced opportunities
    if (enhancedCount > 0) {
      const sampleEnhanced = enhancedOpportunities.find(opp => opp.contactAddress || opp.contactPostcode);
      if (sampleEnhanced) {
        this.logger.log(`📋 Sample enhanced opportunity: ${sampleEnhanced.id} - contactAddress: "${sampleEnhanced.contactAddress}", contactPostcode: "${sampleEnhanced.contactPostcode}"`);
      }
    }
    
    return enhancedOpportunities;
  }

  /**
   * UNIFIED APPROACH: Get opportunities with "confirmed" or "appt confirmed" tags from specific stages
   * This method:
   * 1. Gets opportunities from all 4 stages: AI Bot Home Survey Booked, Manual Home Survey Booked, 
   *    Manual Online Quotes, and AI Bot Online Quotes
   * 2. Filters opportunities that have the "confirmed" or "appt confirmed" tag
   * Returns only confirmed opportunities from these specific stages
   */
  async getOpportunitiesWithAppointmentsUnified(userId: string) {
    const credentials = this.getGhlCredentials();
    if (!credentials) {
      this.logger.warn('GHL credentials not configured - returning empty response');
      return { opportunities: [], total: 0 };
    }

    try {
      // Get user details
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      this.logger.log(`🎯 CONFIRMED OPPORTUNITIES FETCH: Starting for user: ${user.name} (${user.role})`);

      // Get opportunities from all 4 stages: AI Bot and Manual Home Survey Booked stages, plus Online Quotes stages
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      const additionalStageId = '69b91b0e-b23d-4d7e-9a15-6bcbf0d05b9b'; // Manual Online Quotes
      const additionalStageId2 = '1bb0bbae-97ca-42ca-b10e-f32097fb189d'; // AI Bot Online Quotes
      
      this.logger.log(`📋 Fetching opportunities from 4 stages:`);
      this.logger.log(`   1. AI Bot Home Survey Booked: ${aiStageId}`);
      this.logger.log(`   2. Manual Home Survey Booked: ${manualStageId}`);
      this.logger.log(`   3. Manual Online Quotes: ${additionalStageId}`);
      this.logger.log(`   4. AI Bot Online Quotes: ${additionalStageId2}`);
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId, additionalStageId, additionalStageId2]
      );

      // Log breakdown by stage BEFORE any filtering
      const aiBeforeFilter = allOpportunities.filter(opp => opp.pipelineStageId === aiStageId);
      const manualBeforeFilter = allOpportunities.filter(opp => opp.pipelineStageId === manualStageId);
      const onlineQuotesManualBeforeFilter = allOpportunities.filter(opp => opp.pipelineStageId === additionalStageId);
      const onlineQuotesAiBeforeFilter = allOpportunities.filter(opp => opp.pipelineStageId === additionalStageId2);
      
      this.logger.log(`📊 Found ${allOpportunities.length} total opportunities from all 4 stages:`);
      this.logger.log(`   - AI Bot Home Survey Booked: ${aiBeforeFilter.length} opportunities`);
      this.logger.log(`   - Manual Home Survey Booked: ${manualBeforeFilter.length} opportunities`);
      this.logger.log(`   - Manual Online Quotes: ${onlineQuotesManualBeforeFilter.length} opportunities`);
      this.logger.log(`   - AI Bot Online Quotes: ${onlineQuotesAiBeforeFilter.length} opportunities`);

      // Filter opportunities based on user role
      const filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        allOpportunities,
        user,
        credentials
      );

      // Log breakdown by stage AFTER user role filtering
      const aiAfterUserFilter = filteredOpportunities.filter(opp => opp.pipelineStageId === aiStageId);
      const manualAfterUserFilter = filteredOpportunities.filter(opp => opp.pipelineStageId === manualStageId);
      const onlineQuotesManualAfterUserFilter = filteredOpportunities.filter(opp => opp.pipelineStageId === additionalStageId);
      const onlineQuotesAiAfterUserFilter = filteredOpportunities.filter(opp => opp.pipelineStageId === additionalStageId2);
      
      this.logger.log(`👤 After user role filtering: ${filteredOpportunities.length} opportunities for user ${user.name}`);
      this.logger.log(`   - AI Bot Home Survey Booked: ${aiAfterUserFilter.length} opportunities`);
      this.logger.log(`   - Manual Home Survey Booked: ${manualAfterUserFilter.length} opportunities`);
      this.logger.log(`   - Manual Online Quotes: ${onlineQuotesManualAfterUserFilter.length} opportunities`);
      this.logger.log(`   - AI Bot Online Quotes: ${onlineQuotesAiAfterUserFilter.length} opportunities`);

      // TAG-BASED APPROACH: Filter opportunities that have "confirmed" or "appt confirmed" tags
      this.logger.log(`🏷️  Filtering opportunities by "confirmed" or "appt confirmed" tags...`);
      const confirmedOpportunities = filteredOpportunities.filter(opp => {
        const hasTag = this.checkForConfirmedTag(opp);
        if (!hasTag && (opp.pipelineStageId === additionalStageId || opp.pipelineStageId === additionalStageId2)) {
          // Log Online Quotes opportunities that don't have tags for debugging
          const tags = opp.contact?.tags || opp.tags || opp.tag || [];
          const tagNames = Array.isArray(tags) 
            ? tags.map(tag => typeof tag === 'string' ? tag : tag.name || tag.title || '').filter(Boolean)
            : [tags].filter(Boolean);
          this.logger.log(`⚠️  Online Quotes opportunity "${opp.name}" missing confirmed tag. Tags: [${tagNames.join(', ')}]`);
        }
        return hasTag;
      });

      // Log breakdown by stage AFTER tag filtering
      const aiAfterTagFilter = confirmedOpportunities.filter(opp => opp.pipelineStageId === aiStageId);
      const manualAfterTagFilter = confirmedOpportunities.filter(opp => opp.pipelineStageId === manualStageId);
      const onlineQuotesManualAfterTagFilter = confirmedOpportunities.filter(opp => opp.pipelineStageId === additionalStageId);
      const onlineQuotesAiAfterTagFilter = confirmedOpportunities.filter(opp => opp.pipelineStageId === additionalStageId2);
      
      this.logger.log(`✅ After tag filtering: ${confirmedOpportunities.length} opportunities with "confirmed" or "appt confirmed" tags`);
      this.logger.log(`   - AI Bot Home Survey Booked: ${aiAfterTagFilter.length} opportunities`);
      this.logger.log(`   - Manual Home Survey Booked: ${manualAfterTagFilter.length} opportunities`);
      this.logger.log(`   - Manual Online Quotes: ${onlineQuotesManualAfterTagFilter.length} opportunities`);
      this.logger.log(`   - AI Bot Online Quotes: ${onlineQuotesAiAfterTagFilter.length} opportunities`);

      // ENHANCE OPPORTUNITIES WITH APPOINTMENT INFO
      const opportunitiesWithAppointmentInfo = await this.enhanceOpportunitiesWithAppointmentInfo(
        confirmedOpportunities,
        credentials
      );

      // Log breakdown by stage AFTER enhancement
      const aiOpportunities = opportunitiesWithAppointmentInfo.filter(opp => opp.pipelineStageId === aiStageId);
      const manualOpportunities = opportunitiesWithAppointmentInfo.filter(opp => opp.pipelineStageId === manualStageId);
      const additionalOpportunities = opportunitiesWithAppointmentInfo.filter(opp => opp.pipelineStageId === additionalStageId);
      const additionalOpportunities2 = opportunitiesWithAppointmentInfo.filter(opp => opp.pipelineStageId === additionalStageId2);
      
      this.logger.log(`📅 Final result: ${opportunitiesWithAppointmentInfo.length} confirmed opportunities with appointment info`);
      this.logger.log(`   - AI Bot Home Survey Booked: ${aiOpportunities.length} opportunities`);
      this.logger.log(`   - Manual Home Survey Booked: ${manualOpportunities.length} opportunities`);
      this.logger.log(`   - Manual Online Quotes: ${additionalOpportunities.length} opportunities`);
      this.logger.log(`   - AI Bot Online Quotes: ${additionalOpportunities2.length} opportunities`);

      return {
        opportunities: opportunitiesWithAppointmentInfo,
        total: opportunitiesWithAppointmentInfo.length,
        user: {
          id: user.id,
          name: user.name,
          role: user.role,
        },
        method: 'confirmed-tag-with-appointment-extraction',
        breakdown: {
          aiHomeSurveyBooked: aiOpportunities.length,
          manualHomeSurveyBooked: manualOpportunities.length,
          manualOnlineQuotes: additionalOpportunities.length,
          aiOnlineQuotes: additionalOpportunities2.length,
        }
      };
    } catch (error) {
      this.logger.error(`Error in getOpportunitiesWithAppointmentsUnified: ${error.message}`);
      throw error;
    }
  }

  /**
   * HYBRID APPROACH: Process opportunities using both manual (tags) and automatic (appointments) methods
   */
  private async processOpportunitiesWithHybridAppointments(
    opportunities: any[],
    credentials: any
  ): Promise<any[]> {
    this.logger.log(`🔍 Processing ${opportunities.length} opportunities with contact notes appointment detection`);

    const opportunitiesWithAppointments: any[] = [];
    const batchSize = 5; // Process in small batches to avoid rate limiting
    const delayBetweenBatches = 1000; // 1 second delay

    for (let i = 0; i < opportunities.length; i += batchSize) {
      const batch = opportunities.slice(i, i + batchSize);
      
      const batchPromises = batch.map(async (opportunity) => {
        try {
          const contactId = opportunity.contactId || opportunity.contact?.id;
          if (!contactId) {
            this.logger.warn(`❌ No contact ID for opportunity ${opportunity.name}`);
            return null;
          }

          // ONLY METHOD: Check for appointments in contact notes (no GHL API appointments)
          const hasAppointmentTag = this.checkForConfirmedTag(opportunity);
          
          if (hasAppointmentTag) {
            this.logger.log(`✅ Opportunity ${opportunity.name} has confirmed tag, checking contact notes`);
            
            // Fetch contact notes from the separate notes endpoint
            const contactNotes = await this.goHighLevelService.getContactNotes(
              credentials.accessToken,
              contactId
            );

            // Extract appointment details from contact notes
            const appointmentDetails = this.extractAppointmentDetailsFromNotes(contactNotes);

            if (appointmentDetails && appointmentDetails.isValidDate) {
              this.logger.log(`✅ Found appointment details for ${opportunity.name}: ${appointmentDetails.rawText}`);
              return {
                ...opportunity,
                type: opportunity.pipelineStageId === '8904bbe1-53a3-468e-94e4-f13cb04a4947' ? 'ai' : 'manual',
                hasAppointment: true,
                appointmentDetails: {
                  ...appointmentDetails,
                  appointmentType: 'manual'
                },
                appointmentSource: 'contact_notes',
                classification: 'CONFIRMED'
              };
            } else if (appointmentDetails) {
              // Tag exists but no valid date in notes - still count as appointment
              this.logger.log(`✅ Found appointment tag for ${opportunity.name} but no valid date in notes`);
              return {
                ...opportunity,
                type: opportunity.pipelineStageId === '8904bbe1-53a3-468e-94e4-f13cb04a4947' ? 'ai' : 'manual',
                hasAppointment: true,
                appointmentDetails: {
                  date: null,
                  rawText: appointmentDetails.rawText || 'Appointment booked (details in CRM)',
                  notes: appointmentDetails.notes || '',
                  extractedFrom: 'appointment_tag_only',
                  appointmentType: 'manual',
                  isValidDate: false
                },
                appointmentSource: 'contact_notes_tag_only',
                classification: 'TAGGED'
              };
            }
          }

          return null;

        } catch (error) {
          this.logger.error(`❌ Error processing opportunity ${opportunity.name}: ${error.message}`);
          return null;
        }
      });

      const batchResults = await Promise.all(batchPromises);
      const validResults = batchResults.filter(result => result !== null);
      opportunitiesWithAppointments.push(...validResults);

      this.logger.log(`📦 Processed batch ${Math.floor(i / batchSize) + 1}, found ${validResults.length} valid opportunities`);

      // Add delay between batches to avoid rate limiting
      if (i + batchSize < opportunities.length) {
        await new Promise(resolve => setTimeout(resolve, delayBetweenBatches));
      }
    }

    return opportunitiesWithAppointments;
  }

  /**
   * Process opportunities by checking tags for "appt booked" and getting appointment details from contact notes
   */
  private async processOpportunitiesWithTagBasedAppointments(
    opportunities: any[],
    credentials: any
  ): Promise<any[]> {
    this.logger.log(`🔍 Processing ${opportunities.length} opportunities for tag-based appointments`);

    const opportunitiesWithAppointments: any[] = [];
    const batchSize = 5; // Process in small batches to avoid rate limiting
    const delayBetweenBatches = 1000; // 1 second delay

    for (let i = 0; i < opportunities.length; i += batchSize) {
      const batch = opportunities.slice(i, i + batchSize);
      
      const batchPromises = batch.map(async (opportunity) => {
        try {
          // Check if opportunity has "confirmed" tag
          const hasAppointmentTag = this.checkForConfirmedTag(opportunity);
          
          if (!hasAppointmentTag) {
            this.logger.log(`❌ Opportunity ${opportunity.name} does not have confirmed tag`);
            return null;
          }

          this.logger.log(`✅ Opportunity ${opportunity.name} has confirmed tag, fetching contact details`);

          // Get contact details to extract appointment info from notes
          const contactId = opportunity.contactId || opportunity.contact?.id;
          if (!contactId) {
            this.logger.warn(`❌ No contact ID for opportunity ${opportunity.name}`);
            return null;
          }

          const contactDetails = await this.goHighLevelService.getContactById(
            credentials.accessToken,
            contactId
          );

          if (!contactDetails?.contact) {
            this.logger.warn(`❌ No contact details found for ${contactId}`);
            return null;
          }

          // Fetch contact notes from the separate notes endpoint
          const contactNotes = await this.goHighLevelService.getContactNotes(
            credentials.accessToken,
            contactId
          );

          // Extract appointment details from contact notes
          const appointmentDetails = this.extractAppointmentDetailsFromNotes(contactNotes);

          // Even if we can't extract specific appointment details, 
          // we still return the opportunity since it has the appointment tag
          const defaultAppointmentDetails = {
            date: null,
            rawText: 'Appointment booked (details in CRM)',
            notes: contactNotes.join(' ') || '',
            extractedFrom: appointmentDetails ? 'contact_notes' : 'appointment_tag_only'
          };

          // Return enhanced opportunity
          return {
            ...opportunity,
            type: opportunity.pipelineStageId === '8904bbe1-53a3-468e-94e4-f13cb04a4947' ? 'ai' : 'manual',
            hasAppointment: true,
            appointmentDetails: appointmentDetails || defaultAppointmentDetails,
            classification: appointmentDetails ? 'CONFIRMED' : 'TAGGED'
          };

        } catch (error) {
          this.logger.error(`❌ Error processing opportunity ${opportunity.name}: ${error.message}`);
          return null;
        }
      });

      const batchResults = await Promise.all(batchPromises);
      const validResults = batchResults.filter(result => result !== null);
      opportunitiesWithAppointments.push(...validResults);

      this.logger.log(`📦 Processed batch ${Math.floor(i / batchSize) + 1}, found ${validResults.length} valid opportunities`);

      // Add delay between batches to avoid rate limiting
      if (i + batchSize < opportunities.length) {
        await new Promise(resolve => setTimeout(resolve, delayBetweenBatches));
      }
    }

    return opportunitiesWithAppointments;
  }


  /**
   * Check if opportunity has "won" tag
   */
  private checkForWonTag(opportunity: any): boolean {
    // First check if the opportunity status is explicitly "won"
    const status = opportunity.status;
    const hasWonStatus = status && status.toLowerCase() === 'won';
    
    if (hasWonStatus) {
      this.logger.log(`✅ Found won status: "${status}" for opportunity: ${opportunity.name}`);
    }
    
    // Check various possible tag fields - prioritize contact.tags as per GoHighLevel API
    const tags = opportunity.contact?.tags || opportunity.tags || opportunity.tag || [];
    let hasWonTag = false;
    
    if (Array.isArray(tags)) {
      hasWonTag = tags.some(tag => {
        const tagText = typeof tag === 'string' ? tag : tag.name || tag.title || '';
        const lowerTagText = tagText.toLowerCase();
        const isWonTag = lowerTagText.includes('won') || 
                        lowerTagText.includes('sold') ||
                        lowerTagText.includes('closed won') ||
                        lowerTagText.includes('deal closed');
        
        if (isWonTag) {
          this.logger.log(`✅ Found won tag: "${tagText}" for opportunity: ${opportunity.name}`);
        }
        
        return isWonTag;
      });
    } else if (typeof tags === 'string') {
      const lowerTagText = tags.toLowerCase();
      hasWonTag = lowerTagText.includes('won') || 
                  lowerTagText.includes('sold') ||
                  lowerTagText.includes('closed won') ||
                  lowerTagText.includes('deal closed');
      
      if (hasWonTag) {
        this.logger.log(`✅ Found won tag in string: "${tags}" for opportunity: ${opportunity.name}`);
      }
    }

    // Return true if either the status is "won" OR there's a won tag
    const isWon = hasWonStatus || hasWonTag;
    
    if (isWon) {
      this.logger.log(`✅ Opportunity ${opportunity.name} is marked as won - Status: ${status}, Has Won Tag: ${hasWonTag}`);
    }
    
    return isWon;
  }

  private checkForConfirmedTag(opportunity: any): boolean {
    // Check various possible tag fields - prioritize contact.tags as per GoHighLevel API
    const tags = opportunity.contact?.tags || opportunity.tags || opportunity.tag || [];
    
    if (Array.isArray(tags)) {
      // Log all tags for debugging
      const tagNames = tags.map(tag => typeof tag === 'string' ? tag : tag.name || tag.title || '').filter(Boolean);
      this.logger.log(`🔍 Checking tags for "${opportunity.name}": [${tagNames.join(', ')}]`);
      
      const hasConfirmedTag = tags.some(tag => {
        const tagText = typeof tag === 'string' ? tag : tag.name || tag.title || '';
        const lowerTagText = tagText.toLowerCase().trim();
        // Check for exact matches or contains "confirmed" or "appt confirmed"
        const isConfirmedTag = lowerTagText === 'confirmed' || 
                              lowerTagText === 'appt confirmed' ||
                              lowerTagText.includes('appt confirmed') ||
                              (lowerTagText.includes('confirmed') && lowerTagText.includes('appt'));
        
        if (isConfirmedTag) {
          this.logger.log(`✅ Found confirmed tag: "${tagText}" for opportunity: ${opportunity.name}`);
        }
        
        return isConfirmedTag;
      });
      
      if (!hasConfirmedTag) {
        this.logger.log(`❌ No confirmed tag found for "${opportunity.name}" - tags checked: [${tagNames.join(', ')}]`);
      }
      
      return hasConfirmedTag;
    }

    if (typeof tags === 'string') {
      const lowerTagText = tags.toLowerCase().trim();
      // Check for exact matches or contains "confirmed" or "appt confirmed"
      const hasConfirmedTag = lowerTagText === 'confirmed' || 
                              lowerTagText === 'appt confirmed' ||
                              lowerTagText.includes('appt confirmed') ||
                              (lowerTagText.includes('confirmed') && lowerTagText.includes('appt'));
      
      this.logger.log(`🔍 Checking tag string for "${opportunity.name}": "${tags}"`);
      
      if (hasConfirmedTag) {
        this.logger.log(`✅ Found confirmed tag in string: "${tags}" for opportunity: ${opportunity.name}`);
      } else {
        this.logger.log(`❌ No confirmed tag found in string for "${opportunity.name}": "${tags}"`);
      }
      
      return hasConfirmedTag;
    }

    this.logger.log(`❌ No tags found for "${opportunity.name}" - tags field type: ${typeof tags}, value: ${JSON.stringify(tags)}`);
    return false;
  }

  /**
   * Format date for API response - preserves local timezone instead of converting to UTC
   */
  private formatDateForAPI(dateInput: any): string {
    if (!dateInput) return '';
    
    try {
      let dateToFormat = dateInput;
      
      // If it's a string and not in ISO format, try to convert it
      if (typeof dateToFormat === 'string' && !dateToFormat.includes('T')) {
        // Convert "2025-09-11 11:00:00" to "2025-09-11T11:00:00"
        dateToFormat = dateToFormat.replace(' ', 'T');
      }
      
      const date = new Date(dateToFormat);
      if (isNaN(date.getTime())) {
        this.logger.warn(`❌ Invalid date format: ${dateInput}`);
        return dateInput; // Return original if can't parse
      }
      
      // Instead of toISOString() which converts to UTC, preserve the local time
      // Format as YYYY-MM-DDTHH:mm:ss (without timezone conversion)
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      const hours = String(date.getHours()).padStart(2, '0');
      const minutes = String(date.getMinutes()).padStart(2, '0');
      const seconds = String(date.getSeconds()).padStart(2, '0');
      
      return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
    } catch (error) {
      this.logger.warn(`❌ Date formatting error for ${dateInput}: ${error.message}`);
      return dateInput; // Return original if can't format
    }
  }

  /**
   * Extract appointment details from contact notes
   */
  private extractAppointmentDetailsFromNotes(notes: any[]): any | null {
    this.logger.log(`🔍 Extracting appointment details from contact notes:`, notes);
    this.logger.log(`🔍 Notes type:`, typeof notes);
    this.logger.log(`🔍 Notes length:`, Array.isArray(notes) ? notes.length : 'not an array');
    
    if (!notes || !Array.isArray(notes) || notes.length === 0) {
      this.logger.log(`❌ No notes found for contact`);
      return null;
    }

    // Extract text content from note objects (they have a 'body' property)
    const noteTexts = notes.map(note => {
      if (typeof note === 'string') {
        return note;
      } else if (note && typeof note === 'object' && note.body) {
        return note.body;
      } else {
        return '';
      }
    }).filter(text => text.trim() !== '');

    // Combine all note texts into a single string for pattern matching
    const combinedNotes = noteTexts.join(' ');
    this.logger.log(`🔍 Combined notes:`, combinedNotes);

    // Look for appointment patterns in notes
    // Common patterns: "Appointment: [date]", "Booked: [date]", "Survey: [date]", etc.
    const appointmentPatterns = [
      // PRIORITY 1: Full date-time format "Friday, September 12, 2025 10:00 AM"
      /(monday|tuesday|wednesday|thursday|friday|saturday|sunday)[,\s]+(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s+[AP]M/i,
      // PRIORITY 2: Date-time without day name "September 12, 2025 10:00 AM"
      /(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s+[AP]M/i,
      // PRIORITY 3: Date without time "September 12, 2025"
      /(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},\s+\d{4}/i,
      // PRIORITY 4: More general patterns
      /appointment[:\s]+([^,\n]+)/i,
      /booked[:\s]+([^,\n]+)/i,
      /survey[:\s]+([^,\n]+)/i,
      /visit[:\s]+([^,\n]+)/i,
      /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i, // Date pattern
      /(\d{1,2}\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{2,4})/i, // Date with month name
      /(monday|tuesday|wednesday|thursday|friday|saturday|sunday)[\s,]+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i, // Day + date
      /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})[\s,]+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)/i, // Date + day
      // LAST RESORT: Day name only (this should be last to avoid false matches)
      /^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$/i,
    ];

    for (let i = 0; i < appointmentPatterns.length; i++) {
      const pattern = appointmentPatterns[i];
      const match = combinedNotes.match(pattern);
      if (match) {
        // For priority patterns 1-3, use the full match (match[0]) to get the complete date string
        // For other patterns, use match[1] if available, otherwise match[0]
        const appointmentText = (i < 3) ? match[0] : (match[1] || match[0]);
        
        this.logger.log(`✅ Found appointment pattern (priority ${i + 1}): "${appointmentText}"`);
        
        // Try to parse the date with multiple strategies
        let appointmentDate: Date | null = null;
        try {
          // Strategy 1: Handle the specific format "Friday, September 12, 2025 10:00 AM" (Priority 1)
          if (i === 0) {
            // appointmentText now contains the full date string like "Friday, September 12, 2025 10:00 AM"
            const fullDateTimeMatch = appointmentText.match(/(\w+),\s+(\w+)\s+(\d{1,2}),\s+(\d{4})\s+(\d{1,2}):(\d{2})\s+(AM|PM)/i);
            if (fullDateTimeMatch) {
              const [, dayName, monthName, day, year, hour, minute, ampm] = fullDateTimeMatch;
              const monthMap = {
                'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5,
                'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11
              };
              
              const monthIndex = monthMap[monthName.toLowerCase()];
              if (monthIndex !== undefined) {
                let hour24 = parseInt(hour);
                if (ampm.toUpperCase() === 'PM' && hour24 !== 12) {
                  hour24 += 12;
                } else if (ampm.toUpperCase() === 'AM' && hour24 === 12) {
                  hour24 = 0;
                }
                
                appointmentDate = new Date(parseInt(year), monthIndex, parseInt(day), hour24, parseInt(minute), 0);
                this.logger.log(`✅ Parsed full date-time (Priority 1): ${appointmentDate.toISOString()}`);
              }
            }
          }
          
          // Strategy 2: Handle "September 12, 2025 10:00 AM" (Priority 2)
          if (!appointmentDate && i === 1) {
            // appointmentText now contains the full date string like "September 12, 2025 10:00 AM"
            const dateTimeMatch = appointmentText.match(/(\w+)\s+(\d{1,2}),\s+(\d{4})\s+(\d{1,2}):(\d{2})\s+(AM|PM)/i);
            if (dateTimeMatch) {
              const [, monthName, day, year, hour, minute, ampm] = dateTimeMatch;
              const monthMap = {
                'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5,
                'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11
              };
              
              const monthIndex = monthMap[monthName.toLowerCase()];
              if (monthIndex !== undefined) {
                let hour24 = parseInt(hour);
                if (ampm.toUpperCase() === 'PM' && hour24 !== 12) {
                  hour24 += 12;
                } else if (ampm.toUpperCase() === 'AM' && hour24 === 12) {
                  hour24 = 0;
                }
                
                appointmentDate = new Date(parseInt(year), monthIndex, parseInt(day), hour24, parseInt(minute), 0);
                this.logger.log(`✅ Parsed date-time without day (Priority 2): ${appointmentDate.toISOString()}`);
              }
            }
          }
          
          // Strategy 3: Handle "September 12, 2025" (Priority 3)
          if (!appointmentDate && i === 2) {
            // appointmentText now contains the full date string like "September 12, 2025"
            const dateMatch = appointmentText.match(/(\w+)\s+(\d{1,2}),\s+(\d{4})/i);
            if (dateMatch) {
              const [, monthName, day, year] = dateMatch;
              const monthMap = {
                'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5,
                'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11
              };
              
              const monthIndex = monthMap[monthName.toLowerCase()];
              if (monthIndex !== undefined) {
                appointmentDate = new Date(parseInt(year), monthIndex, parseInt(day), 10, 0, 0); // Default to 10 AM
                this.logger.log(`✅ Parsed date without time (Priority 3): ${appointmentDate.toISOString()}`);
              }
            }
          }
          
          // Strategy 4: Direct parsing for other formats (Priority 4+)
          if (!appointmentDate && i >= 3) {
            appointmentDate = new Date(appointmentText);
            if (isNaN(appointmentDate.getTime())) {
              appointmentDate = null;
            } else {
              this.logger.log(`✅ Direct parsing successful (Priority ${i + 1}): ${appointmentDate.toISOString()}`);
            }
          }
          
          // Strategy 5: If direct parsing fails, try to clean up the text
          if (!appointmentDate && i >= 3) {
            // Remove common prefixes and clean up the text
            let cleanText = appointmentText
              .replace(/^(appointment|booked|survey|visit)[:\s]+/i, '')
              .replace(/^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)[,\s]+/i, '')
              .trim();
            
            appointmentDate = new Date(cleanText);
            if (isNaN(appointmentDate.getTime())) {
              appointmentDate = null;
            } else {
              this.logger.log(`✅ Cleaned text parsing successful (Priority ${i + 1}): ${appointmentDate.toISOString()}`);
            }
          }
          
          // Strategy 6: Try to extract just the date part
          if (!appointmentDate && i >= 3) {
            const dateMatch = appointmentText.match(/(\w+ \d{1,2}, \d{4})/);
            if (dateMatch) {
              appointmentDate = new Date(dateMatch[1]);
              if (isNaN(appointmentDate.getTime())) {
                appointmentDate = null;
              } else {
                this.logger.log(`✅ Date part extraction successful (Priority ${i + 1}): ${appointmentDate.toISOString()}`);
              }
            }
          }
          
          // Strategy 7: Try parsing with time
          if (!appointmentDate && i >= 3) {
            const dateTimeMatch = appointmentText.match(/(\w+ \d{1,2}, \d{4} \d{1,2}:\d{2} [AP]M)/);
            if (dateTimeMatch) {
              appointmentDate = new Date(dateTimeMatch[1]);
              if (isNaN(appointmentDate.getTime())) {
                appointmentDate = null;
              } else {
                this.logger.log(`✅ Date-time extraction successful (Priority ${i + 1}): ${appointmentDate.toISOString()}`);
              }
            }
          }
          
          // Strategy 8: Day name fallback (LAST RESORT - Priority 10)
          if (!appointmentDate && i === 9) {
            const dayName = appointmentText.trim().toLowerCase();
            const dayMap = {
              'sunday': 0, 'monday': 1, 'tuesday': 2, 'wednesday': 3,
              'thursday': 4, 'friday': 5, 'saturday': 6
            };
            
            const targetDay = dayMap[dayName];
            if (targetDay !== undefined) {
              const today = new Date();
              const daysUntilTarget = (targetDay - today.getDay() + 7) % 7;
              const nextOccurrence = new Date(today);
              nextOccurrence.setDate(today.getDate() + (daysUntilTarget === 0 ? 7 : daysUntilTarget));
              nextOccurrence.setHours(10, 0, 0, 0); // Default to 10 AM
              appointmentDate = nextOccurrence;
              this.logger.log(`✅ Day name fallback successful (Priority 10): ${appointmentDate.toISOString()}`);
            }
          }
          
        } catch (error) {
          this.logger.warn(`❌ Date parsing failed for "${appointmentText}": ${error.message}`);
          appointmentDate = null;
        }

        const result = {
          date: appointmentDate ? this.formatDateForAPI(appointmentDate) : null,
          rawText: appointmentText,
          notes: combinedNotes,
          extractedFrom: 'contact_notes',
          parsedDate: appointmentDate ? appointmentDate.toISOString() : null,
          isValidDate: !!appointmentDate
        };
        
        this.logger.log(`✅ Extracted appointment details:`, result);
        this.logger.log(`📅 Parsed date: ${appointmentDate ? appointmentDate.toISOString() : 'Failed to parse'}`);
        
        // If we found a valid date, return it immediately
        if (appointmentDate) {
          return result;
        }
        
        // If no valid date found with this pattern, continue to next pattern
        this.logger.log(`❌ No valid date found with pattern ${i + 1}, trying next pattern...`);
      }
    }

    this.logger.log(`❌ No appointment patterns found in notes`);
    return null;
  }

  private filterAndDeduplicateCustomFields(rawCustomFields: any[]): any[] {
    if (!rawCustomFields || rawCustomFields.length === 0) {
      return [];
    }

    // Create a map to track unique field IDs and their values
    const fieldMap = new Map<string, any>();
    
    for (const field of rawCustomFields) {
      if (!field || !field.id || !field.value) {
        continue; // Skip invalid fields
      }

      const fieldId = field.id;
      const fieldValue = field.value;

      // Skip empty or null values
      if (fieldValue === '' || fieldValue === 'null' || fieldValue === null || fieldValue === undefined) {
        continue;
      }

      // If we haven't seen this field ID before, or if the new value is more descriptive
      if (!fieldMap.has(fieldId) || this.isMoreDescriptiveValue(fieldValue, fieldMap.get(fieldId).value)) {
        fieldMap.set(fieldId, {
          id: fieldId,
          value: fieldValue
        });
      }
    }

    // Convert map back to array and sort by field ID for consistency
    const filteredFields = Array.from(fieldMap.values()).sort((a, b) => a.id.localeCompare(b.id));
    
    this.logger.log(`🔍 Filtered ${rawCustomFields.length} raw fields down to ${filteredFields.length} unique fields`);
    
    return filteredFields;
  }

  private isMoreDescriptiveValue(newValue: string, existingValue: string): boolean {
    // Prefer longer, more descriptive values
    if (newValue.length > existingValue.length) {
      return true;
    }
    
    // Prefer values that contain more meaningful information
    const meaningfulWords = ['house', 'detached', 'semi', 'owner', 'employed', 'solar', 'battery', 'budget', 'bedroom'];
    const newValueLower = newValue.toLowerCase();
    const existingValueLower = existingValue.toLowerCase();
    
    const newValueScore = meaningfulWords.reduce((score, word) => 
      score + (newValueLower.includes(word) ? 1 : 0), 0);
    const existingValueScore = meaningfulWords.reduce((score, word) => 
      score + (existingValueLower.includes(word) ? 1 : 0), 0);
    
    return newValueScore > existingValueScore;
  }

  /**
   * Simple method to check ALL opportunities for status="won" - for debugging
   */
  async getAllWonOpportunities(userId: string) {
    this.logger.log(`🔍 DEBUG: Getting ALL won opportunities for user: ${userId}`);
    
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found: ${userId}`);
        return { success: false, error: 'User not found' };
      }

      const credentials = this.getGhlCredentials();
      if (!credentials) {
        this.logger.warn('GHL credentials not configured');
        return { success: false, error: 'GHL credentials not configured' };
      }

      // Get ALL opportunities from the stages (same as working endpoints)
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      this.logger.log(`🔍 DEBUG: Found ${allOpportunities.length} total opportunities from stages`);

      // Filter by user if they have a GHL ID
      let userOpportunities = allOpportunities;
      if (user.ghlUserId) {
        userOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
          allOpportunities,
          user,
          credentials
        );
        this.logger.log(`🔍 DEBUG: After user filtering: ${userOpportunities.length} opportunities`);
      }

      // Check ALL opportunities for status="won" (more flexible matching)
      const wonOpportunities = userOpportunities.filter(opp => {
        const status = opp.status;
        const isWon = status && (
          status.toLowerCase() === 'won' ||
          status.toLowerCase().includes('won') ||
          status.toLowerCase() === 'closed won' ||
          status.toLowerCase() === 'sold'
        );
        
        if (isWon) {
          this.logger.log(`✅ FOUND WON OPPORTUNITY: ${opp.name} - Status: ${status}`);
        }
        
        return isWon;
      });

      this.logger.log(`🔍 DEBUG: Found ${wonOpportunities.length} opportunities with status="won"`);

      // Also check for "won" tags
      const wonByTag = userOpportunities.filter(opp => {
        const tags = opp.contact?.tags || opp.tags || opp.tag || [];
        let hasWonTag = false;
        
        if (Array.isArray(tags)) {
          hasWonTag = tags.some(tag => {
            const tagText = typeof tag === 'string' ? tag : tag.name || tag.title || '';
            const lowerTagText = tagText.toLowerCase();
            return lowerTagText.includes('won') || 
                   lowerTagText.includes('sold') ||
                   lowerTagText.includes('closed won') ||
                   lowerTagText === 'won';
          });
        } else if (typeof tags === 'string') {
          const lowerTags = tags.toLowerCase();
          hasWonTag = lowerTags.includes('won') || 
                      lowerTags.includes('sold') ||
                      lowerTags.includes('closed won') ||
                      lowerTags === 'won';
        }
        
        if (hasWonTag) {
          this.logger.log(`✅ FOUND WON BY TAG: ${opp.name} - Tags: ${JSON.stringify(tags)}`);
        }
        
        return hasWonTag;
      });

      this.logger.log(`🔍 DEBUG: Found ${wonByTag.length} opportunities with "won" tag`);

      // Log some examples
      if (wonOpportunities.length > 0) {
        this.logger.log(`🔍 DEBUG: Won by status examples:`, wonOpportunities.slice(0, 3).map(opp => ({
          name: opp.name,
          status: opp.status,
          stageId: opp.pipelineStageId,
          value: opp.monetaryValue
        })));
      }

      if (wonByTag.length > 0) {
        this.logger.log(`🔍 DEBUG: Won by tag examples:`, wonByTag.slice(0, 3).map(opp => ({
          name: opp.name,
          status: opp.status,
          tags: opp.contact?.tags || opp.tags || opp.tag,
          stageId: opp.pipelineStageId,
          value: opp.monetaryValue
        })));
      }

      return {
        success: true,
        data: {
          totalOpportunities: userOpportunities.length,
          wonByStatus: wonOpportunities.length,
          wonByTag: wonByTag.length,
          wonByStatusExamples: wonOpportunities.slice(0, 5),
          wonByTagExamples: wonByTag.slice(0, 5)
        }
      };

    } catch (error) {
      this.logger.error('Error in getAllWonOpportunities:', error.stack);
      return { success: false, error: error.message };
    }
  }

  /**
   * Debug method to show all opportunities for a user (both assigned and owned)
   */
  async debugAllUserOpportunities(userId: string) {
    this.logger.log(`🔍 DEBUG: Getting ALL opportunities for user: ${userId}`);
    
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found: ${userId}`);
        return { success: false, error: 'User not found' };
      }

      const credentials = this.getGhlCredentials();
      if (!credentials) {
        this.logger.warn('GHL credentials not configured');
        return { success: false, error: 'GHL credentials not configured' };
      }

      // Get opportunities from AI and Manual stages
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      this.logger.log(`🔍 DEBUG: Found ${allOpportunities.length} total opportunities from stages`);

      // Get unique user IDs from opportunities
      const userIds = [...new Set(allOpportunities
        .map(opp => opp.assignedTo)
        .filter(id => id))];

      // Fetch user details from GHL API
      let userMap = new Map<string, any>();
      if (userIds.length > 0) {
        userMap = await this.goHighLevelService.getUsersByIds(credentials.accessToken, userIds);
      }

      // Check for opportunities assigned to user
      const assignedOpportunities = allOpportunities.filter(opp => {
        const assignedToId = opp.assignedTo;
        if (assignedToId && userMap) {
          const assignedUser = userMap.get(assignedToId);
          if (assignedUser) {
            const assignedUserName = assignedUser.name || assignedUser.firstName + ' ' + assignedUser.lastName;
            return assignedUserName.toLowerCase().includes(user.name?.toLowerCase() || '') ||
                   (user.name?.toLowerCase() || '').includes(assignedUserName.toLowerCase());
          }
        }
        return false;
      });

      // Check for opportunities owned by user (check various ownership fields)
      const ownedOpportunities = allOpportunities.filter(opp => {
        // Check various ownership fields
        const ownerFields = [
          opp.ownerId,
          opp.createdBy,
          opp.userId,
          opp.contact?.ownerId,
          opp.contact?.createdBy
        ];
        
        return ownerFields.some(field => {
          if (typeof field === 'string') {
            return field.toLowerCase().includes(user.name?.toLowerCase() || '') ||
                   (user.name?.toLowerCase() || '').includes(field.toLowerCase());
          }
          return false;
        });
      });

      // Check for opportunities with user name in any field
      const nameBasedOpportunities = allOpportunities.filter(opp => {
        const searchFields = [
          opp.name,
          opp.assignedTo,
          opp.teamMember,
          opp.assignee,
          opp.assignedToName,
          opp.ownerId,
          opp.createdBy,
          opp.userId,
          opp.contact?.firstName,
          opp.contact?.lastName,
          opp.contact?.ownerId,
          opp.contact?.createdBy
        ];
        
        return searchFields.some(field => {
          if (typeof field === 'string') {
            return field.toLowerCase().includes(user.name?.toLowerCase() || '') ||
                   (user.name?.toLowerCase() || '').includes(field.toLowerCase());
          }
          return false;
        });
      });

      // Log sample opportunities to see their structure
      const sampleOpportunities = allOpportunities.slice(0, 3);
      this.logger.log(`🔍 DEBUG: Sample opportunities structure:`, JSON.stringify(sampleOpportunities, null, 2));

      return {
        success: true,
        data: {
          totalOpportunities: allOpportunities.length,
          assignedOpportunities: assignedOpportunities.length,
          ownedOpportunities: ownedOpportunities.length,
          nameBasedOpportunities: nameBasedOpportunities.length,
          assignedList: assignedOpportunities.slice(0, 10),
          ownedList: ownedOpportunities.slice(0, 10),
          nameBasedList: nameBasedOpportunities.slice(0, 10),
          user: {
            id: user.id,
            name: user.name,
            role: user.role
          }
        }
      };

    } catch (error) {
      this.logger.error('Error in debugAllUserOpportunities:', error.stack);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get opportunities with won status from Private Customers pipeline
   */
  async getOpportunitiesWithWon(userId: string) {
    this.logger.log(`🔍 Getting opportunities with won status for user: ${userId}`);
    
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found: ${userId}`);
        return { success: false, error: 'User not found' };
      }

      const credentials = this.getGhlCredentials();
      if (!credentials) {
        this.logger.warn('GHL credentials not configured');
        return { success: false, error: 'GHL credentials not configured' };
      }

      // Get opportunities from AI and Manual stages (same as working endpoints)
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      const allOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      this.logger.log(`🔍 Found ${allOpportunities.length} total opportunities from stages`);

      // Filter by user if they have a GHL ID
      let userOpportunities = allOpportunities;
      if (user.ghlUserId) {
        userOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
          allOpportunities,
          user,
          credentials
        );
        this.logger.log(`🔍 After user filtering: ${userOpportunities.length} opportunities for user ${user.name}`);
      }

      // Log all opportunities for debugging (no filtering by won status as requested)
      this.logger.log(`🔍 Found ${userOpportunities.length} total opportunities for user ${user.name} in Private Customers pipeline`);
      
      // Log first few opportunities for debugging
      userOpportunities.slice(0, 5).forEach((opp, index) => {
        this.logger.log(`🔍 Opportunity ${index + 1}: ${opp.name} - Status: ${opp.status || 'No Status'} - Value: £${opp.value || 0}`);
      });

      // Calculate total value of all opportunities
      const totalValue = userOpportunities.reduce((sum, opp) => sum + (opp.value || 0), 0);

      return {
        success: true,
        data: {
          totalOpportunities: userOpportunities.length,
          totalValue,
          opportunitiesList: userOpportunities.slice(0, 10), // First 10 opportunities
          user: {
            id: user.id,
            name: user.name,
            role: user.role
          }
        }
      };

    } catch (error) {
      this.logger.error('Error getting opportunities with won status:', error.stack);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get sales performance statistics for reps dashboard
   * This replaces the AI vs Manual opportunities view with sales-focused metrics
   */
  async getSalesPerformanceStats(userId: string, month?: string, year?: string) {
    this.logger.log(`📊 Fetching sales performance stats for user: ${userId}, month: ${month}, year: ${year}`);
    
    try {
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      const credentials = this.getGhlCredentials();
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty sales performance response');
        return {
          appointments: { count: 0, value: 0 },
          sales: { count: 0, value: 0 },
          conversionRate: 0,
          monthlyBreakdown: []
        };
      }

      // Get current date for default month/year
      const now = new Date();
      const targetMonth = month || (now.getMonth() + 1).toString();
      const targetYear = year || now.getFullYear().toString();

      this.logger.log(`📅 Analyzing data for ${targetMonth}/${targetYear}`);

      // Private Customers pipeline ID
      const privateCustomersPipelineId = 'FxPA8fVU11VnudThxhFy';

      // Get all opportunities from Private Customers pipeline
      const pipelineOpportunities = await this.goHighLevelService.getOpportunitiesByPipeline(
        credentials.accessToken,
        privateCustomersPipelineId
      );

      if (!pipelineOpportunities || !pipelineOpportunities.opportunities) {
        this.logger.warn('No opportunities found in Private Customers pipeline');
        return {
          appointments: { count: 0, value: 0 },
          sales: { count: 0, value: 0 },
          conversionRate: 0,
          monthlyBreakdown: []
        };
      }

      // Filter opportunities by the logged-in user (assignedTo field)
      let userOpportunities;
      
      if (user.ghlUserId) {
        // User has a GHL ID, filter by it
        userOpportunities = pipelineOpportunities.opportunities.filter(opp => 
          opp.assignedTo === user.ghlUserId
        );
        this.logger.log(`👤 Found ${userOpportunities.length} opportunities assigned to user ${user.name} (GHL ID: ${user.ghlUserId})`);
      } else {
        // User doesn't have a GHL ID, for now return all opportunities for testing
        // TODO: This should be replaced with proper GHL ID assignment
        userOpportunities = pipelineOpportunities.opportunities;
        this.logger.warn(`⚠️  User ${user.name} (${user.id}) has no ghlUserId assigned. Returning all opportunities for testing.`);
        this.logger.warn(`   Please assign a GHL user ID to this user for proper filtering.`);
      }

      // Get all opportunities that have "confirmed" or "appt confirmed" tags (regardless of stage or date)
      const allAppointmentOpportunities = userOpportunities.filter(opp => 
        this.checkForConfirmedTag(opp)
      );

      this.logger.log(`📅 Found ${allAppointmentOpportunities.length} opportunities with "confirmed" or "appt confirmed" tags`);

      // Use ALL appointment opportunities (no date filtering for dashboard)
      const appointments = allAppointmentOpportunities;

      // Find all won opportunities (don't require appointment tag as it might be removed when won)
      const wonOpportunities = userOpportunities.filter(opp => 
        this.checkForWonTag(opp) && 
        opp.monetaryValue && opp.monetaryValue > 0
      );

      // Calculate values
      const appointmentsValue = appointments.reduce((sum, opp) => sum + (opp.monetaryValue || 0), 0);
      const salesValue = wonOpportunities.reduce((sum, opp) => sum + (opp.monetaryValue || 0), 0);
      const conversionRate = appointments.length > 0 ? (wonOpportunities.length / appointments.length) * 100 : 0;

      // Log detailed information about opportunities for debugging
      const statusCounts = appointments.reduce((acc, opp) => {
        acc[opp.status] = (acc[opp.status] || 0) + 1;
        return acc;
      }, {});
      
      this.logger.log(`📊 Opportunity Status Breakdown:`, statusCounts);
      this.logger.log(`📊 Sales Performance Stats (Tag-based):
        - Appointments (appt booked): ${appointments.length} (value: $${appointmentsValue})
        - Sales Won (won tag or status): ${wonOpportunities.length} (value: $${salesValue})
        - Conversion Rate: ${conversionRate.toFixed(2)}%`);
      
      // Log some examples of sales opportunities for debugging
      if (wonOpportunities.length > 0) {
        this.logger.log(`📊 Sales Won examples:`, wonOpportunities.slice(0, 3).map(opp => ({
          name: opp.name,
          status: opp.status,
          stageId: opp.pipelineStageId,
          value: opp.monetaryValue
        })));
      }

      // Get monthly breakdown for the last 6 months
      const monthlyBreakdown = await this.getMonthlyBreakdown(userOpportunities, 6);

      const result = {
        appointments: {
          count: appointments.length,
          value: appointmentsValue
        },
        sales: {
          count: wonOpportunities.length,
          value: salesValue
        },
        conversionRate: Math.round(conversionRate * 100) / 100, // Round to 2 decimal places
        monthlyBreakdown,
        period: {
          month: targetMonth,
          year: targetYear
        },
        user: {
          id: user.id,
          name: user.name,
          role: user.role
        }
      };

      this.logger.log(`✅ Sales performance stats calculated:`, {
        appointments: result.appointments.count,
        sales: result.sales.count,
        conversionRate: result.conversionRate,
        salesValue: result.sales.value
      });

      return result;
    } catch (error) {
      this.logger.error(`Error in getSalesPerformanceStats: ${error.message}`);
      throw error;
    }
  }

  /**
   * Filter opportunities by month and year
   */
  private filterOpportunitiesByMonth(opportunities: any[], month: number, year: number): any[] {
    return opportunities.filter(opp => {
      const oppDate = new Date(opp.createdAt || opp.created_at);
      return oppDate.getMonth() + 1 === month && oppDate.getFullYear() === year;
    });
  }


  private filterWonOpportunitiesByMonth(wonOpportunities: any[], month: number, year: number): any[] {
    return wonOpportunities.filter(opp => {
      // Use lastStatusChangeAt if available (when status changed to 'won'), otherwise use updatedAt
      const wonDate = new Date(opp.lastStatusChangeAt || opp.updatedAt || opp.createdAt);
      return wonDate.getMonth() + 1 === month && wonDate.getFullYear() === year;
    });
  }

  /**
   * Get monthly breakdown for the last N months
   */
  private async getMonthlyBreakdown(opportunities: any[], months: number): Promise<any[]> {
    const breakdown: any[] = [];
    const now = new Date();

    for (let i = months - 1; i >= 0; i--) {
      const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const month = date.getMonth() + 1;
      const year = date.getFullYear();
      const monthName = date.toLocaleString('default', { month: 'short' });

      // Get all opportunities that have "confirmed" or "appt confirmed" tags (same logic as main method)
      const allAppointmentOpportunities = opportunities.filter(opp => 
        this.checkForConfirmedTag(opp)
      );

      // Filter appointment opportunities by the month they were created
      const monthlyAppointments = this.filterOpportunitiesByMonth(
        allAppointmentOpportunities,
        month,
        year
      );
      
      // Find won opportunities for this month (don't require appointment tag as it might be removed when won)
      const monthlyWonOpportunities = opportunities.filter(opp => 
        this.checkForWonTag(opp) && 
        opp.monetaryValue && opp.monetaryValue > 0
      );
      
      const conversionRate = monthlyAppointments.length > 0 ? (monthlyWonOpportunities.length / monthlyAppointments.length) * 100 : 0;

      breakdown.push({
        month: monthName,
        year: year,
        monthNumber: month,
        appointments: monthlyAppointments.length,
        sales: monthlyWonOpportunities.length,
        conversionRate: Math.round(conversionRate * 100) / 100,
        salesValue: monthlyWonOpportunities.reduce((sum, opp) => sum + (opp.monetaryValue || 0), 0)
      });
    }

    return breakdown;
  }

  /**
   * Get all pipelines from GoHighLevel
   */
  async getPipelines(userId: string): Promise<any> {
    try {
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        return { success: false, error: 'GHL credentials not configured' };
      }

      this.logger.log('Fetching pipelines from GoHighLevel');
      
      const pipelines = await this.goHighLevelService.getPipelines(credentials.accessToken, credentials.locationId);
      
      this.logger.log(`Successfully fetched ${pipelines?.length || 0} pipelines`);
      
      return {
        success: true,
        data: { pipelines: pipelines },
        count: pipelines?.length || 0
      };
    } catch (error) {
      this.logger.error('Error fetching pipelines:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get opportunities by pipeline ID (filtered by user)
   */
  async getOpportunitiesByPipeline(userId: string, pipelineId: string): Promise<any> {
    try {
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        return { success: false, error: 'GHL credentials not configured' };
      }

      // Get user details
      const user = await this.userService.findById(userId);
      if (!user) {
        this.logger.error(`User not found: ${userId}`);
        return { success: false, error: 'User not found' };
      }

      this.logger.log(`Fetching opportunities for pipeline: ${pipelineId} for user: ${user.name}`);
      this.logger.log(`User GHL ID: ${user.ghlUserId || 'NOT ASSIGNED'}`);
      
      // For efficiency, if user has no GHL ID, try to find it first
      if (!user.ghlUserId && user.name) {
        this.logger.log(`🔍 User ${user.name} has no GHL ID, attempting to find it...`);
        try {
          const ghlUser = await this.goHighLevelService.findUserByName(
            credentials.accessToken,
            credentials.locationId,
            user.name
          );
          if (ghlUser) {
            user.ghlUserId = ghlUser.id;
            this.logger.log(`✅ Found GHL user ID for ${user.name}: ${ghlUser.id}`);
            
            // Update the user's GHL user ID in the database
            await this.userService.update(user.id, { ghlUserId: ghlUser.id });
            this.logger.log(`💾 Updated user ${user.name} with GHL user ID: ${ghlUser.id}`);
          } else {
            this.logger.warn(`❌ Could not find GHL user for: ${user.name}`);
          }
        } catch (error) {
          this.logger.error(`❌ Error finding GHL user for ${user.name}: ${error.message}`);
        }
      }
      
      const opportunities = await this.goHighLevelService.getOpportunitiesByPipeline(
        credentials.accessToken, 
        pipelineId,
        user.ghlUserId || undefined // Pass the user's GHL ID for smart filtering
      );
      
      this.logger.log(`Successfully fetched ${opportunities?.opportunities?.length || 0} opportunities for pipeline ${pipelineId}`);
      
      // The GoHighLevel service now handles filtering, so we can use the results directly
      let userOpportunities = opportunities.opportunities;
      
      if (user.ghlUserId && opportunities.meta?.filtered) {
        // The GoHighLevel service already filtered the opportunities
        this.logger.log(`👤 GoHighLevel service returned ${userOpportunities.length} pre-filtered opportunities for user ${user.name} (GHL ID: ${user.ghlUserId})`);
        this.logger.log(`🎯 Smart filtering was applied - early termination used for efficiency`);
      } else if (user.ghlUserId) {
        // Fallback: filter manually if the GoHighLevel service didn't filter
        userOpportunities = opportunities.opportunities.filter(opp => 
          opp.assignedTo === user.ghlUserId
        );
        this.logger.log(`👤 Manually filtered to ${userOpportunities.length} opportunities for user ${user.name} (GHL ID: ${user.ghlUserId})`);
      } else {
        // User doesn't have a GHL ID, for now return all opportunities for testing
        this.logger.warn(`⚠️  User ${user.name} (${user.id}) has no ghlUserId assigned. Returning all opportunities for testing.`);
        this.logger.warn(`   Please assign a GHL user ID to this user for proper filtering.`);
      }
      
      // Log opportunity status breakdown for ratio calculation
      if (userOpportunities.length > 0) {
        const statusBreakdown = userOpportunities.reduce((acc, opp) => {
          const status = opp.status || opp.pipelineStageId || 'unknown';
          acc[status] = (acc[status] || 0) + 1;
          return acc;
        }, {});
        
        this.logger.log(`📊 Opportunity Status Breakdown for ${user.name}:`);
        Object.entries(statusBreakdown).forEach(([status, count]) => {
          this.logger.log(`   - ${status}: ${count} opportunities`);
        });
        
        // Debug: Log sample tags from opportunities
        const sampleOpportunities = userOpportunities.slice(0, 5);
        this.logger.log(`🔍 Sample opportunity tags for debugging:`);
        sampleOpportunities.forEach((opp, index) => {
          const oppWithTags = opp as any;
          const contactTags = oppWithTags.contact?.tags || [];
          const opportunityTags = oppWithTags.tags || [];
          this.logger.log(`   Opportunity ${index + 1}: ${opp.name}`);
          this.logger.log(`     - Contact Tags: ${JSON.stringify(contactTags)}`);
          this.logger.log(`     - Opportunity Tags: ${JSON.stringify(opportunityTags)}`);
        });
        
        // Calculate potential booking vs won ratio
        // Look for "appt booked" tag (same logic as DashboardScreen)
        const bookedCount = userOpportunities.filter(opp => {
          const oppWithTags = opp as any;
          // Check contact.tags first as per GoHighLevel API structure
          const tags = oppWithTags.contact?.tags || oppWithTags.tags || [];
          
          // Check if opportunity has "appt booked" tag
          const hasApptBookedTag = Array.isArray(tags) 
            ? tags.some(tag => {
                const tagText = typeof tag === 'string' ? tag : tag.name || tag.title || '';
                return tagText === 'appt booked';
              })
            : false;
          
          return hasApptBookedTag;
        }).length;
        
        const wonCount = userOpportunities.filter(opp => 
          opp.status?.toLowerCase().includes('won') || 
          opp.status?.toLowerCase().includes('closed') ||
          opp.pipelineStageId?.includes('won') ||
          opp.pipelineStageId?.includes('closed')
        ).length;
        
        this.logger.log(`🎯 Ratio Analysis for ${user.name}:`);
        this.logger.log(`   - Booked/Survey opportunities: ${bookedCount}`);
        this.logger.log(`   - Won/Closed opportunities: ${wonCount}`);
        if (bookedCount > 0) {
          const winRate = ((wonCount / bookedCount) * 100).toFixed(1);
          this.logger.log(`   - Win Rate: ${winRate}% (${wonCount}/${bookedCount})`);
        }
      }
      
      // Calculate ratio metrics for the response
      let ratioMetrics: any = null;
      if (userOpportunities.length > 0) {
        const bookedCount = userOpportunities.filter(opp => {
          const oppWithTags = opp as any;
          // Check contact.tags first as per GoHighLevel API structure
          const tags = oppWithTags.contact?.tags || oppWithTags.tags || [];
          
          // Check if opportunity has "appt booked" tag
          const hasApptBookedTag = Array.isArray(tags) 
            ? tags.some(tag => {
                const tagText = typeof tag === 'string' ? tag : tag.name || tag.title || '';
                return tagText === 'appt booked';
              })
            : false;
          
          return hasApptBookedTag;
        }).length;
        
        const wonCount = userOpportunities.filter(opp => 
          opp.status?.toLowerCase().includes('won') || 
          opp.status?.toLowerCase().includes('closed') ||
          opp.pipelineStageId?.includes('won') ||
          opp.pipelineStageId?.includes('closed')
        ).length;
        
        ratioMetrics = {
          bookedCount,
          wonCount,
          winRate: bookedCount > 0 ? ((wonCount / bookedCount) * 100).toFixed(1) : '0.0',
          totalOpportunities: userOpportunities.length
        };
      }

      return {
        success: true,
        data: {
          opportunities: userOpportunities,
          meta: {
            total: userOpportunities.length,
            pipelineId: pipelineId,
            user: {
              id: user.id,
              name: user.name,
              ghlUserId: user.ghlUserId
            },
            ratioMetrics: ratioMetrics
          }
        },
        pipelineId: pipelineId,
        count: userOpportunities.length
      };
    } catch (error) {
      this.logger.error(`Error fetching opportunities for pipeline ${pipelineId}:`, error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get opportunities by stage progression (e.g., "Installation Survey Booked")
   */
  async getOpportunitiesByStageProgression(userId: string, stageName: string): Promise<any> {
    try {
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        return { success: false, error: 'GHL credentials not configured' };
      }

      this.logger.log(`Fetching opportunities for stage: ${stageName}`);
      
      // First get all pipelines to find the stage
      const pipelinesResponse = await this.goHighLevelService.getPipelines(credentials.accessToken, credentials.locationId);
      
      if (!pipelinesResponse || pipelinesResponse.length === 0) {
        return { success: false, error: 'No pipelines found' };
      }

      // Find the stage ID for the given stage name
      // Prioritize "Private Customers" pipeline first
      let targetStageId = null;
      let targetPipelineId = null;
      
      // First, try to find the stage in "Private Customers" pipeline
      const privateCustomersPipeline = pipelinesResponse.find(p => p.name === "Private Customers");
      if (privateCustomersPipeline) {
        for (const stage of privateCustomersPipeline.stages) {
          if (stage.name.toLowerCase().includes(stageName.toLowerCase()) || 
              stageName.toLowerCase().includes(stage.name.toLowerCase())) {
            targetStageId = stage.id;
            targetPipelineId = privateCustomersPipeline.id;
            this.logger.log(`Found stage "${stage.name}" with ID: ${targetStageId} in Private Customers pipeline`);
            break;
          }
        }
      }
      
      // If not found in Private Customers, search all pipelines
      if (!targetStageId) {
        for (const pipeline of pipelinesResponse) {
          for (const stage of pipeline.stages) {
            if (stage.name.toLowerCase().includes(stageName.toLowerCase()) || 
                stageName.toLowerCase().includes(stage.name.toLowerCase())) {
              targetStageId = stage.id;
              targetPipelineId = pipeline.id;
              this.logger.log(`Found stage "${stage.name}" with ID: ${targetStageId} in pipeline: ${pipeline.name}`);
              break;
            }
          }
          if (targetStageId) break;
        }
      }

      if (!targetStageId) {
        return { 
          success: false, 
          error: `Stage "${stageName}" not found in any pipeline`,
          availableStages: pipelinesResponse.flatMap(p => p.stages.map(s => s.name))
        };
      }

      // Get opportunities by specific stage ID (same method as working dashboard)
      const stageOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [targetStageId]
      );
      
      if (!stageOpportunities || stageOpportunities.length === 0) {
        return { success: false, error: 'No opportunities found' };
      }
      
      this.logger.log(`Found ${stageOpportunities.length} opportunities in stage "${stageName}" (ID: ${targetStageId})`);

      // Get user info for filtering
      const user = await this.userService.findById(userId);
      if (!user) {
        return { success: false, error: 'User not found' };
      }

      // For testing purposes, let's be less restrictive with filtering
      // First try with full filtering
      let filteredOpportunities = await this.filterOpportunitiesByUserRoleAndAppointments(
        stageOpportunities,
        user,
        credentials
      );
      
      // If no results after filtering, log the issue and return unfiltered results for testing
      if (filteredOpportunities.length === 0 && stageOpportunities.length > 0) {
        this.logger.warn(`⚠️  Filtering removed all ${stageOpportunities.length} opportunities. Returning unfiltered results for testing.`);
        this.logger.log(`User role: ${user.role}, User name: ${user.name}`);
        filteredOpportunities = stageOpportunities;
      }

      return {
        success: true,
        data: {
          opportunities: filteredOpportunities,
          stage: {
            id: targetStageId,
            name: stageName,
            pipelineId: targetPipelineId
          },
          totalCount: stageOpportunities.length,
          filteredCount: filteredOpportunities.length
        }
      };
    } catch (error) {
      this.logger.error(`Error fetching opportunities for stage "${stageName}":`, error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get opportunities by stage progression without user filtering (for testing)
   */
  async getOpportunitiesByStageProgressionUnfiltered(userId: string, stageName: string): Promise<any> {
    try {
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        return { success: false, error: 'GHL credentials not configured' };
      }

      this.logger.log(`Fetching UNFILTERED opportunities for stage: ${stageName}`);
      
      // First get all pipelines to find the stage
      const pipelinesResponse = await this.goHighLevelService.getPipelines(credentials.accessToken, credentials.locationId);
      
      if (!pipelinesResponse || pipelinesResponse.length === 0) {
        return { success: false, error: 'No pipelines found' };
      }

      // Find the stage ID for the given stage name
      // Prioritize "Private Customers" pipeline first
      let targetStageId = null;
      let targetPipelineId = null;
      
      // First, try to find the stage in "Private Customers" pipeline
      const privateCustomersPipeline = pipelinesResponse.find(p => p.name === "Private Customers");
      if (privateCustomersPipeline) {
        for (const stage of privateCustomersPipeline.stages) {
          if (stage.name.toLowerCase().includes(stageName.toLowerCase()) || 
              stageName.toLowerCase().includes(stage.name.toLowerCase())) {
            targetStageId = stage.id;
            targetPipelineId = privateCustomersPipeline.id;
            this.logger.log(`Found stage "${stage.name}" with ID: ${targetStageId} in Private Customers pipeline`);
            break;
          }
        }
      }
      
      // If not found in Private Customers, search all pipelines
      if (!targetStageId) {
        for (const pipeline of pipelinesResponse) {
          for (const stage of pipeline.stages) {
            if (stage.name.toLowerCase().includes(stageName.toLowerCase()) || 
                stageName.toLowerCase().includes(stage.name.toLowerCase())) {
              targetStageId = stage.id;
              targetPipelineId = pipeline.id;
              this.logger.log(`Found stage "${stage.name}" with ID: ${targetStageId} in pipeline: ${pipeline.name}`);
              break;
            }
          }
          if (targetStageId) break;
        }
      }

      if (!targetStageId) {
        return { 
          success: false, 
          error: `Stage "${stageName}" not found in any pipeline`,
          availableStages: pipelinesResponse.flatMap(p => p.stages.map(s => s.name))
        };
      }

      // Get opportunities by specific stage ID (same method as working dashboard)
      const stageOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [targetStageId]
      );
      
      if (!stageOpportunities || stageOpportunities.length === 0) {
        return { success: false, error: 'No opportunities found' };
      }
      
      this.logger.log(`Found ${stageOpportunities.length} UNFILTERED opportunities in stage "${stageName}" (ID: ${targetStageId})`);

      return {
        success: true,
        data: {
          opportunities: stageOpportunities,
          stage: {
            id: targetStageId,
            name: stageName,
            pipelineId: targetPipelineId
          },
          total: stageOpportunities.length,
          note: "These are unfiltered results for testing purposes"
        }
      };
    } catch (error) {
      this.logger.error('Error fetching unfiltered opportunities by stage progression:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Update opportunity status and stage in GoHighLevel
   */
  async updateOpportunityStatus(
    opportunityId: string,
    status: 'open' | 'won' | 'lost' | 'abandoned',
    stageId?: string
  ): Promise<{ success: boolean; error?: string; data?: any }> {
    try {
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        return { success: false, error: 'GHL credentials not configured' };
      }

      this.logger.log(`Updating opportunity ${opportunityId} status to ${status}${stageId ? ` and stage to ${stageId}` : ''}`);

      // Use the Private Customers pipeline ID
      const pipelineId = 'FxPA8fVU11VnudThxhFy';

      const result = await this.goHighLevelService.updateOpportunityStatus(
        credentials.accessToken,
        pipelineId,
        opportunityId,
        status,
        stageId
      );

      if (result.success) {
        this.logger.log(`Successfully updated opportunity ${opportunityId} in GoHighLevel`);
      } else {
        this.logger.error(`Failed to update opportunity ${opportunityId}: ${result.error}`);
      }

      return result;
    } catch (error) {
      this.logger.error(`Error updating opportunity status: ${error.message}`);
      return { success: false, error: error.message };
    }
  }

  /**
   * Move opportunity to "Signed Contract" stage when won
   */
  async moveOpportunityToSignedContract(opportunityId: string): Promise<{ success: boolean; error?: string; data?: any }> {
    try {
      // Stage ID for "Signed Contract" from the pipeline test screen
      const signedContractStageId = '09107d21-d594-4301-9d27-de95525bef11';
      
      this.logger.log(`Moving opportunity ${opportunityId} to Signed Contract stage`);
      
      return await this.updateOpportunityStatus(opportunityId, 'won', signedContractStageId);
    } catch (error) {
      this.logger.error(`Error moving opportunity to signed contract stage: ${error.message}`);
      return { success: false, error: error.message };
    }
  }

  /**
   * Sync opportunities from GHL to database
   * Saves/updates opportunities with all relevant data (customer, rep, status, etc.)
   */
  async syncOpportunitiesFromGHL(userId: string, limit?: number): Promise<{
    synced: number;
    updated: number;
    created: number;
    errors: number;
  }> {
    this.logger.log(`🔄 Starting opportunity sync from GHL for user: ${userId}`);
    
    try {
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        throw new Error('GHL credentials not configured');
      }

      // Only get opportunities from the specific stages (AI and Manual Home Survey Booked)
      // This is much faster than fetching all opportunities
      const aiStageId = '8904bbe1-53a3-468e-94e4-f13cb04a4947'; // (AI Bot) Home Survey Booked
      const manualStageId = '97cbf1b8-31c2-4486-9edc-5a3d5d0c198c'; // (Manual) Home Survey Booked
      
      this.logger.log(`📊 Fetching opportunities from 2 specific stages only (much faster)`);
      
      // Get opportunities from these specific stages only
      const ghlOpportunities = await this.goHighLevelService.getOpportunitiesByStages(
        credentials.accessToken,
        credentials.locationId,
        [aiStageId, manualStageId]
      );

      this.logger.log(`📊 Found ${ghlOpportunities.length} opportunities from specific stages`);

      // Limit if specified
      const opportunitiesToSync = limit 
        ? ghlOpportunities.slice(0, limit)
        : ghlOpportunities;

      this.logger.log(`📊 Syncing ${opportunitiesToSync.length} opportunities`);

      let synced = 0;
      let updated = 0;
      let created = 0;
      let errors = 0;

      // Process opportunities in batches to avoid overwhelming the database
      const batchSize = 50;
      for (let i = 0; i < opportunitiesToSync.length; i += batchSize) {
        const batch = opportunitiesToSync.slice(i, i + batchSize);
        
        await Promise.all(
          batch.map(async (ghlOpp: any) => {
            try {
              await this.upsertOpportunityFromGHL(ghlOpp, credentials);
              synced++;
            } catch (error) {
              this.logger.error(`Error syncing opportunity ${ghlOpp.id}:`, error.message);
              errors++;
            }
          })
        );

        // Small delay between batches
        if (i + batchSize < opportunitiesToSync.length) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      }

      this.logger.log(`✅ Sync complete: ${synced} synced, ${created} created, ${updated} updated, ${errors} errors`);
      
      return { synced, updated, created, errors };
    } catch (error) {
      this.logger.error(`Error syncing opportunities from GHL:`, error.stack);
      throw error;
    }
  }

  /**
   * Upsert a single opportunity from GHL data
   */
  private async upsertOpportunityFromGHL(
    ghlOpp: any,
    credentials: { accessToken: string; locationId: string }
  ): Promise<void> {
    try {
      // Extract customer information
      const contact = ghlOpp.contact || {};
      const customerName = this.extractCustomerName(ghlOpp, contact);
      const customerFirstName = contact.firstName || null;
      const customerLastName = contact.lastName || null;
      const customerEmail = contact.email || null;
      const customerPhone = contact.phone || null;
      
      // Extract address information
      const addressObj = contact.address || {};
      const customerAddress = addressObj.address1 || addressObj.address || null;
      const customerCity = addressObj.city || null;
      const customerState = addressObj.state || null;
      const customerPostcode = addressObj.postalCode || addressObj.zipCode || null;

      // Find assigned user by GHL user ID
      let assignedUserId: string | null = null;
      if (ghlOpp.assignedTo) {
        const assignedUser = await this.prisma.user.findFirst({
          where: { ghlUserId: ghlOpp.assignedTo },
          select: { id: true, name: true },
        });
        if (assignedUser) {
          assignedUserId = assignedUser.id;
        }
      }

      // Get assigned user name from GHL or our DB
      let assignedToName: string | null = null;
      if (ghlOpp.assignedTo) {
        if (assignedUserId) {
          const user = await this.prisma.user.findUnique({
            where: { id: assignedUserId },
            select: { name: true, ghlUserName: true },
          });
          assignedToName = user?.name || user?.ghlUserName || null;
        }
        // If not found in DB, try to get from GHL (would need additional API call)
      }

      // Extract tags
      const tags = Array.isArray(ghlOpp.tags) 
        ? ghlOpp.tags.map((tag: any) => typeof tag === 'string' ? tag : tag.name || tag.title || '')
        : [];

      // Get outcome from existing OpportunityOutcome if it exists
      const existingOutcome = await this.prisma.opportunityOutcome.findUnique({
        where: { ghlOpportunityId: ghlOpp.id },
        select: { outcome: true },
      });

      // Upsert the opportunity
      await this.prisma.opportunity.upsert({
        where: { ghlOpportunityId: ghlOpp.id },
        create: {
          ghlOpportunityId: ghlOpp.id,
          name: ghlOpp.name || 'Unnamed Opportunity',
          userId: assignedUserId,
          assignedToGhlId: ghlOpp.assignedTo || null,
          assignedToName: assignedToName,
          contactId: ghlOpp.contactId || contact.id || null,
          customerName: customerName,
          customerFirstName: customerFirstName,
          customerLastName: customerLastName,
          customerEmail: customerEmail,
          customerPhone: customerPhone,
          customerAddress: customerAddress,
          customerCity: customerCity,
          customerState: customerState,
          customerPostcode: customerPostcode,
          monetaryValue: ghlOpp.monetaryValue || null,
          status: ghlOpp.status || null,
          pipelineId: ghlOpp.pipelineId || null,
          pipelineName: ghlOpp.pipelineName || null,
          pipelineStageId: ghlOpp.pipelineStageId || null,
          stageName: ghlOpp.stageName || null,
          outcome: existingOutcome?.outcome || null,
          tags: tags,
          notes: ghlOpp.notes || null,
          source: ghlOpp.source || null,
          ghlData: ghlOpp as any, // Store full GHL data
          ghlCreatedAt: ghlOpp.createdAt ? new Date(ghlOpp.createdAt) : null,
          ghlUpdatedAt: ghlOpp.updatedAt ? new Date(ghlOpp.updatedAt) : null,
          lastSyncedAt: new Date(),
        },
        update: {
          name: ghlOpp.name || 'Unnamed Opportunity',
          userId: assignedUserId,
          assignedToGhlId: ghlOpp.assignedTo || null,
          assignedToName: assignedToName,
          contactId: ghlOpp.contactId || contact.id || null,
          customerName: customerName,
          customerFirstName: customerFirstName,
          customerLastName: customerLastName,
          customerEmail: customerEmail,
          customerPhone: customerPhone,
          customerAddress: customerAddress,
          customerCity: customerCity,
          customerState: customerState,
          customerPostcode: customerPostcode,
          monetaryValue: ghlOpp.monetaryValue || null,
          status: ghlOpp.status || null,
          pipelineId: ghlOpp.pipelineId || null,
          pipelineName: ghlOpp.pipelineName || null,
          pipelineStageId: ghlOpp.pipelineStageId || null,
          stageName: ghlOpp.stageName || null,
          outcome: existingOutcome?.outcome || null,
          tags: tags,
          notes: ghlOpp.notes || null,
          source: ghlOpp.source || null,
          ghlData: ghlOpp as any,
          ghlCreatedAt: ghlOpp.createdAt ? new Date(ghlOpp.createdAt) : null,
          ghlUpdatedAt: ghlOpp.updatedAt ? new Date(ghlOpp.updatedAt) : null,
          lastSyncedAt: new Date(),
        },
      });
    } catch (error) {
      this.logger.error(`Error upserting opportunity ${ghlOpp.id}:`, error);
      throw error;
    }
  }

  /**
   * Extract customer name from opportunity and contact data
   */
  private extractCustomerName(opportunity: any, contact: any): string | null {
    // Priority 1: First and last name combination
    if (contact?.firstName && contact?.lastName) {
      return `${contact.firstName.trim()} ${contact.lastName.trim()}`;
    }
    // Priority 2: Full name field
    if (contact?.name && contact.name.trim() !== '') {
      return contact.name.trim();
    }
    // Priority 3: Opportunity name/title
    if (opportunity?.name && opportunity.name.trim() !== '') {
      return opportunity.name.trim();
    }
    // Priority 4: Extract from email
    if (contact?.email) {
      const emailPrefix = contact.email.split('@')[0];
      return emailPrefix.replace(/[._-]/g, ' ').replace(/\b\w/g, (l: string) => l.toUpperCase());
    }
    return null;
  }

  /**
   * Get opportunities from database (with optional filters)
   */
  async getOpportunitiesFromDB(filters?: {
    userId?: string;
    outcome?: string;
    status?: string;
    startDate?: Date;
    endDate?: Date;
    limit?: number;
  }): Promise<any[]> {
    try {
      const where: any = {};

      if (filters?.userId) {
        where.userId = filters.userId;
      }

      if (filters?.outcome) {
        where.outcome = filters.outcome;
      }

      if (filters?.status) {
        where.status = filters.status;
      }

      if (filters?.startDate || filters?.endDate) {
        where.createdAt = {};
        if (filters.startDate) where.createdAt.gte = filters.startDate;
        if (filters.endDate) where.createdAt.lte = filters.endDate;
      }

      const opportunities = await this.prisma.opportunity.findMany({
        where,
        include: {
          user: {
            select: {
              id: true,
              name: true,
              email: true,
              ghlUserName: true,
            },
          },
        },
        orderBy: { lastSyncedAt: 'desc' },
        take: filters?.limit || 1000,
      });

      return opportunities;
    } catch (error) {
      this.logger.error('Error getting opportunities from DB:', error);
      throw error;
    }
  }
} 