import { Injectable, Logger } from '@nestjs/common';
import axios from 'axios';

@Injectable()
export class GoHighLevelService {
  private readonly baseUrl = 'https://rest.gohighlevel.com/v1';
  private readonly logger = new Logger(GoHighLevelService.name);
  
  // Simple in-memory cache for API responses
  private cache = new Map<string, { data: any; timestamp: number }>();
  private readonly CACHE_TTL = 5 * 60 * 1000; // 5 minutes
  
  // Track active requests to prevent duplicates
  private activeRequests: Set<string> | null = null;

  private getCacheKey(method: string, params: any): string {
    return `${method}:${JSON.stringify(params)}`;
  }

  private getCachedData(key: string): any | null {
    const cached = this.cache.get(key);
    if (cached && Date.now() - cached.timestamp < this.CACHE_TTL) {
      this.logger.log(`Cache hit for key: ${key}`);
      return cached.data;
    }
    return null;
  }

  private setCachedData(key: string, data: any): void {
    this.cache.set(key, { data, timestamp: Date.now() });
    this.logger.log(`Cache set for key: ${key}`);
  }

  private getHeaders(accessToken: string) {
    return {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    };
  }

  // Fetch all pipelines and stages for a location with retry logic
  async getPipelines(accessToken: string, locationId: string): Promise<any[]> {
    const url = `${this.baseUrl}/pipelines/`;
    const maxRetries = 3;
    const baseDelay = 2000; // 2 seconds
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        this.logger.log(`Fetching pipelines (attempt ${attempt}/${maxRetries})`);
        const response = await axios.get(url, {
          headers: this.getHeaders(accessToken),
          timeout: 15000,
        });
        const pipelines = (response.data as any).pipelines || [];
        this.logger.log(`Successfully fetched ${pipelines.length} pipelines`);
        return pipelines;
      } catch (error: any) {
        // Don't retry on authentication errors (401) - these won't succeed
        if (error.response?.status === 401) {
          this.logger.error(`Authentication error (401) - not retrying: ${error.message}`);
          return [];
        }
        
        if (error.response?.status === 429 && attempt < maxRetries) {
          const delay = baseDelay * attempt; // Exponential backoff
          this.logger.warn(`Rate limited (429), retrying in ${delay}ms (attempt ${attempt}/${maxRetries})`);
          await new Promise(resolve => setTimeout(resolve, delay));
          continue;
        }
        this.logger.error(`Error fetching pipelines (attempt ${attempt}): ${error.message}`);
        if (attempt === maxRetries) {
          return [];
        }
      }
    }
    return [];
  }

  // Fetch opportunities from specific stages in Private Customers pipeline
  async getOpportunitiesByStages(accessToken: string, locationId: string, stageIds: string[]): Promise<any[]> {
    // Check cache first
    const cacheKey = this.getCacheKey('getOpportunitiesByStages', { stageIds: stageIds.sort() });
    const cached = this.getCachedData(cacheKey);
    if (cached) {
      this.logger.log(`Cache hit for stages ${stageIds.join(', ')}, returning ${cached.length} opportunities`);
      return cached;
    }

    this.logger.log(`Starting to fetch opportunities for stages: ${stageIds.join(', ')}`);
    
    try {
      // Only fetch from "Private Customers" pipeline
      const pipelineId = 'FxPA8fVU11VnudThxhFy'; // Private Customers pipeline ID
      let allOpportunities: any[] = [];
      
      // Fetch opportunities for each stage
      for (const stageId of stageIds) {
        this.logger.log(`Fetching opportunities for stage: ${stageId}`);
        
        let stageOpportunities: any[] = [];
        let hasMore = true;
        let startAfterId = null;
        let startAfter = null;
        const limit = 100; // GHL API limit per request
        let iterationCount = 0;
        const maxIterations = 1000; // Safety limit to prevent infinite loops
        
        // Fetch all pages for this stage
        while (hasMore && stageOpportunities.length < 50000 && iterationCount < maxIterations) { // Increased limit to 50,000 to get all opportunities
          iterationCount++;
          const url = `${this.baseUrl}/pipelines/${pipelineId}/opportunities/`;
          const params: any = {
            limit,
            stageId,
          };
          
          if (startAfterId) {
            params.startAfterId = startAfterId;
          }
          
          if (startAfter) {
            params.startAfter = startAfter;
          }
          
          const response = await axios.get(url, {
            headers: this.getHeaders(accessToken),
            params,
            timeout: 15000,
          });
          
          const opportunities = (response.data as any).opportunities || [];
          const meta = (response.data as any).meta || {};
          
          this.logger.log(`Stage ${stageId}: Received ${opportunities.length} opportunities (total so far: ${stageOpportunities.length + opportunities.length})`);
          
          if (opportunities.length === 0) {
            hasMore = false;
          } else {
            stageOpportunities = [...stageOpportunities, ...opportunities];
            
            // Store previous pagination values to detect if we're stuck
            const previousStartAfterId = startAfterId;
            const previousStartAfter = startAfter;
            
            // Update pagination for next request
            startAfterId = meta.startAfterId;
            startAfter = meta.startAfter;
            
            // Check if pagination parameters haven't changed (indicating we're stuck)
            if (startAfterId === previousStartAfterId && startAfter === previousStartAfter && opportunities.length > 0) {
              this.logger.warn(`Stage ${stageId}: Pagination parameters unchanged, breaking loop to prevent infinite loop`);
              hasMore = false;
            }
            
            // If no next page URL, we've reached the end
            if (!meta.nextPageUrl) {
              hasMore = false;
            }
            
            // If we got fewer than the limit, we've reached the end
            if (opportunities.length < limit) {
              hasMore = false;
            }
            
            // Additional safety check: if we're getting very few opportunities per page consistently, stop
            if (opportunities.length <= 1 && stageOpportunities.length > 100) {
              this.logger.warn(`Stage ${stageId}: Getting very few opportunities per page (${opportunities.length}), stopping to prevent infinite loop`);
              hasMore = false;
            }
          }
        }
        
        if (iterationCount >= maxIterations) {
          this.logger.warn(`Stage ${stageId}: Reached maximum iterations (${maxIterations}), stopping to prevent infinite loop`);
        }
        
        this.logger.log(`Stage ${stageId}: Total ${stageOpportunities.length} opportunities`);
        allOpportunities = [...allOpportunities, ...stageOpportunities];
      }
      
      this.logger.log(`Total opportunities from all stages: ${allOpportunities.length}`);
      
      // Cache the result
      this.setCachedData(cacheKey, allOpportunities);
      
      return allOpportunities;
    } catch (error) {
      this.logger.error(`Error fetching opportunities by stages: ${error.message}`);
      return [];
    }
  }

  // Fetch all opportunities for a location (with optional filters)
  async getOpportunities(accessToken: string, locationId: string, filters: any = {}): Promise<any[]> {
    this.logger.log(`Starting to fetch opportunities for locationId: ${locationId}`);
    
    try {
      // First get all pipelines
      const pipelines = await this.getPipelines(accessToken, locationId);
      this.logger.log(`Found ${pipelines.length} pipelines`);
      
      let allOpportunities: any[] = [];
      
      // Fetch from ALL pipelines to ensure we find the opportunity
      this.logger.log(`Found ${pipelines.length} pipelines, searching all of them for opportunities`);
      
      for (let i = 0; i < pipelines.length; i++) {
        const pipeline = pipelines[i];
        this.logger.log(`Fetching opportunities from pipeline: ${pipeline.name} (${pipeline.id})`);
        
        // Add delay between pipelines to avoid rate limiting
        if (i > 0) {
          await new Promise(resolve => setTimeout(resolve, 1000)); // 1 second delay between pipelines
        }
      
        try {
          let pipelineOpportunities: any[] = [];
          let hasMore = true;
          let offset = 0;
          const limit = 100; // GHL API limit per request
          
          // Fetch all pages for this pipeline
          while (hasMore && pipelineOpportunities.length < 50000) { // Increased limit to 50,000 to get all opportunities
            const url = `${this.baseUrl}/pipelines/${pipeline.id}/opportunities/`;
            
            // Add delay between API calls to avoid rate limiting
            if (offset > 0) {
              await new Promise(resolve => setTimeout(resolve, 500)); // 500ms delay between pages
            }
            
            const response = await axios.get(url, {
              headers: this.getHeaders(accessToken),
              params: {
                limit,
                skip: offset,
                ...filters
              },
              timeout: 10000,
            });
            
            const opportunities = (response.data as any).opportunities || [];
            this.logger.log(`${pipeline.name} pipeline: Received ${opportunities.length} opportunities (offset: ${offset}, total so far: ${pipelineOpportunities.length + opportunities.length})`);
            
            if (opportunities.length === 0) {
              hasMore = false;
            } else {
              pipelineOpportunities = [...pipelineOpportunities, ...opportunities];
              offset += limit;
              
              // If we got fewer than the limit, we've reached the end
              if (opportunities.length < limit) {
                hasMore = false;
              }
            }
          }
          
          this.logger.log(`${pipeline.name} pipeline: Total ${pipelineOpportunities.length} opportunities (max 50,000)`);
          
          // Add pipeline info to each opportunity
          const opportunitiesWithPipeline = pipelineOpportunities.map((opp: any) => ({
            ...opp,
            pipelineName: pipeline.name,
            pipelineId: pipeline.id,
          }));
          
          allOpportunities = [...allOpportunities, ...opportunitiesWithPipeline];
        } catch (error) {
          this.logger.error(`Error fetching opportunities from ${pipeline.name} pipeline: ${error.message}`);
        }
      }
      
      this.logger.log(`Successfully fetched ${allOpportunities.length} total opportunities from all pipelines`);
      return allOpportunities;
    } catch (error) {
      this.logger.error(`Error fetching opportunities: ${error.message}`);
      this.logger.error(`Error details: ${JSON.stringify(error.response?.data || {})}`);
      throw error;
    }
  }

  // Map opportunities to their stage names using pipelines data
  async getOpportunitiesWithStageNames(accessToken: string, locationId: string, filters: any = {}): Promise<any[]> {
    const pipelines = await this.getPipelines(accessToken, locationId);
    const opportunities = await this.getOpportunities(accessToken, locationId, filters);
    
    // Build a map of stageId -> stageName for each pipeline
    const stageIdToName: Record<string, string> = {};
    for (const pipeline of pipelines) {
      for (const stage of pipeline.stages || []) {
        if (stage && stage.id && stage.name) {
          stageIdToName[stage.id] = stage.name;
        }
      }
    }
    
    // Attach stageName to each opportunity
    return opportunities.map(opp => ({
      ...opp,
      stageName: stageIdToName[opp.pipelineStageId] || opp.stageName || '',
    }));
  }

  // Fetch a specific opportunity by ID
  async getOpportunityById(accessToken: string, opportunityId: string, locationId?: string): Promise<any> {
    // Use provided locationId or default to the hardcoded one
    const targetLocationId = locationId || '03HhMPsHSZJsAwM77yp6';
    
    // First get all pipelines, then search for the opportunity
    const pipelines = await this.getPipelines(accessToken, targetLocationId);
    
    for (const pipeline of pipelines) {
      try {
        const url = `${this.baseUrl}/pipelines/${pipeline.id}/opportunities/${opportunityId}`;
        const response = await axios.get(url, {
          headers: this.getHeaders(accessToken),
          timeout: 10000,
        });
        return response.data;
      } catch (error) {
        // Continue to next pipeline if not found
        continue;
      }
    }
    
    throw new Error('Opportunity not found');
  }

  // Fetch notes for a specific opportunity
  async getOpportunityNotes(accessToken: string, opportunityId: string): Promise<string[]> {
    try {
      const opportunity = await this.getOpportunityById(accessToken, opportunityId);
      return opportunity.notes || [];
    } catch (error) {
      this.logger.error(`Error fetching opportunity notes: ${error.message}`);
      return [];
    }
  }

  // Alternative method to get opportunity notes
  async getOpportunityNotesAlternative(accessToken: string, opportunityId: string): Promise<string[]> {
    try {
      const opportunity = await this.getOpportunityById(accessToken, opportunityId);
      return opportunity.notes || [];
    } catch (error) {
      this.logger.error(`Error fetching opportunity notes (alternative): ${error.message}`);
      return [];
    }
  }

  // Get opportunity with notes from search
  async getOpportunityWithNotesFromSearch(accessToken: string, opportunityId: string): Promise<string[]> {
    try {
      const opportunities = await this.getOpportunities(accessToken, '03HhMPsHSZJsAwM77yp6');
      const opportunity = opportunities.find((opp: any) => opp.id === opportunityId);
      
      if (opportunity) {
        return opportunity.notes || [];
      }
      
      return [];
    } catch (error) {
      this.logger.error(`Error fetching opportunity with notes from search: ${error.message}`);
      return [];
    }
  }

  // Fetch contact details by ID
  async getContactById(accessToken: string, contactId: string): Promise<any> {
    const url = `${this.baseUrl}/contacts/${contactId}`;
    try {
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        timeout: 10000,
      });
      return response.data;
    } catch (error) {
      this.logger.error(`Error fetching contact by ID: ${error.message}`);
      return null;
    }
  }

  // Fetch contact notes
  async getContactNotes(accessToken: string, contactId: string): Promise<string[]> {
    const url = `${this.baseUrl}/contacts/${contactId}/notes`;
    try {
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        timeout: 10000,
      });
      return (response.data as any).notes || [];
    } catch (error) {
      this.logger.error(`Error fetching contact notes: ${error.message}`);
      return [];
    }
  }

  // Fetch opportunity activities
  async getOpportunityActivities(accessToken: string, opportunityId: string): Promise<string[]> {
    const url = `${this.baseUrl}/opportunities/${opportunityId}/activities`;
    try {
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        timeout: 10000,
      });
      return (response.data as any).activities || [];
    } catch (error) {
      this.logger.error(`Error fetching opportunity activities: ${error.message}`);
      return [];
    }
  }

  // Fetch opportunity communications
  async getOpportunityCommunications(accessToken: string, opportunityId: string): Promise<string[]> {
    const url = `${this.baseUrl}/opportunities/${opportunityId}/communications`;
    try {
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        timeout: 10000,
      });
      return (response.data as any).communications || [];
    } catch (error) {
      this.logger.error(`Error fetching opportunity communications: ${error.message}`);
      return [];
    }
  }

  // Fetch ALL appointments from GoHighLevel (no date filtering, no user/calendar filtering)
  async getAllAppointments(accessToken: string, locationId: string): Promise<any[]> {
    const url = `${this.baseUrl}/appointments/`;
    
    try {
      // Get all appointments without date filtering - just get everything
      const params = {
        includeAll: true, // Include contact and more data
        limit: 10000 // High limit to get all appointments
      };

      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        params: params,
        timeout: 30000, // Longer timeout for large requests
      });
      
      this.logger.log(`All appointments API response status: ${response.status}`);
      this.logger.log(`Fetched ${(response.data as any)?.appointments?.length || 0} total appointments`);
      return (response.data as any).appointments || [];
    } catch (error) {
      this.logger.error(`Error fetching all appointments: ${error.message}`);
      if (error.response) {
        this.logger.error(`All appointments API error status: ${error.response.status}`);
        this.logger.error(`All appointments API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Fetch appointments from GoHighLevel
  async getAppointments(accessToken: string, locationId: string): Promise<any[]> {
    const url = `${this.baseUrl}/appointments/`;
    
    // Calculate date range (last 60 days to next 90 days) - expanded to catch more appointments
    const now = new Date();
    const startDate = new Date(now.getTime() - (60 * 24 * 60 * 60 * 1000)); // 60 days ago
    const endDate = new Date(now.getTime() + (90 * 24 * 60 * 60 * 1000)); // 90 days from now
    
    try {
      // Try different parameter combinations
      const params = {
        startDate: Math.floor(startDate.getTime()),
        endDate: Math.floor(endDate.getTime()),
        includeAll: true, // Include contact and more data
        limit: 1000 // Add limit to get more appointments
      };

      // First try with calendarId (using locationId as calendarId)
      try {
        const response = await axios.get(url, {
          headers: this.getHeaders(accessToken),
          params: { ...params, calendarId: locationId },
          timeout: 10000,
        });
        this.logger.log(`Appointments API response status: ${response.status}`);
        this.logger.log(`Appointments API response data keys: ${Object.keys(response.data || {}).join(', ')}`);
        return (response.data as any).appointments || [];
      } catch (calendarError) {
        this.logger.warn(`CalendarId approach failed, trying userId: ${calendarError.message}`);
        
        // Try with userId (using locationId as userId)
        try {
          const response = await axios.get(url, {
            headers: this.getHeaders(accessToken),
            params: { ...params, userId: locationId },
            timeout: 10000,
          });
          this.logger.log(`Appointments API response status: ${response.status}`);
          this.logger.log(`Appointments API response data keys: ${Object.keys(response.data || {}).join(', ')}`);
          return (response.data as any).appointments || [];
        } catch (userError) {
          this.logger.warn(`UserId approach failed, trying teamId: ${userError.message}`);
          
          // Try with teamId (using locationId as teamId)
          const response = await axios.get(url, {
            headers: this.getHeaders(accessToken),
            params: { ...params, teamId: locationId },
            timeout: 10000,
          });
          this.logger.log(`Appointments API response status: ${response.status}`);
          this.logger.log(`Appointments API response data keys: ${Object.keys(response.data || {}).join(', ')}`);
          return (response.data as any).appointments || [];
        }
      }
    } catch (error) {
      this.logger.error(`Error fetching appointments: ${error.message}`);
      if (error.response) {
        this.logger.error(`Appointments API error status: ${error.response.status}`);
        this.logger.error(`Appointments API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Get all calendars for a location
  async getCalendars(accessToken: string, locationId: string): Promise<any[]> {
    const url = `${this.baseUrl}/calendars/`;
    
    try {
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        params: {
          locationId: locationId,
          limit: 1000 // Get all calendars
        },
        timeout: 10000,
      });
      
      this.logger.log(`Successfully fetched ${(response.data as any)?.calendars?.length || 0} calendars from GHL`);
      return (response.data as any)?.calendars || [];
    } catch (error) {
      this.logger.error(`Error fetching calendars: ${error.message}`);
      if (error.response) {
        this.logger.error(`Calendars API error status: ${error.response.status}`);
        this.logger.error(`Calendars API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Fetch appointments for a specific calendar ID
  async getAppointmentsByCalendarId(accessToken: string, locationId: string, calendarId: string): Promise<any[]> {
    const url = `${this.baseUrl}/appointments/`;
    
    // Calculate date range (last 60 days to next 90 days)
    const now = new Date();
    const startDate = new Date(now.getTime() - (60 * 24 * 60 * 60 * 1000)); // 60 days ago
    const endDate = new Date(now.getTime() + (90 * 24 * 60 * 60 * 1000)); // 90 days from now
    
    try {
      const params = {
        startDate: Math.floor(startDate.getTime()),
        endDate: Math.floor(endDate.getTime()),
        calendarId: calendarId, // Use the specific calendar ID
        includeAll: true, // Include contact and more data
        limit: 1000 // Add limit to get more appointments
      };

      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        params: params,
        timeout: 10000,
      });
      
      this.logger.log(`Appointments by calendarId API response status: ${response.status}`);
      this.logger.log(`Appointments by calendarId API response data keys: ${Object.keys(response.data || {}).join(', ')}`);
      return (response.data as any).appointments || [];
    } catch (error) {
      this.logger.error(`Error fetching appointments by calendarId: ${error.message}`);
      if (error.response) {
        this.logger.error(`Appointments by calendarId API error status: ${error.response.status}`);
        this.logger.error(`Appointments by calendarId API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Fetch appointments for a specific user ID
  async getAppointmentsByUserId(accessToken: string, locationId: string, userId: string): Promise<any[]> {
    const url = `${this.baseUrl}/appointments/`;
    
    // Calculate date range (last 60 days to next 90 days) - expanded to catch more appointments
    const now = new Date();
    const startDate = new Date(now.getTime() - (60 * 24 * 60 * 60 * 1000)); // 60 days ago
    const endDate = new Date(now.getTime() + (90 * 24 * 60 * 60 * 1000)); // 90 days from now
    
    try {
      const params = {
        startDate: Math.floor(startDate.getTime()),
        endDate: Math.floor(endDate.getTime()),
        userId: userId, // Use the specific user ID
        includeAll: true, // Include contact and more data
        limit: 1000 // Add limit to get more appointments
      };

      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        params: params,
        timeout: 10000,
      });
      
      this.logger.log(`Appointments by userId API response status: ${response.status}`);
      this.logger.log(`Appointments by userId API response data keys: ${Object.keys(response.data || {}).join(', ')}`);
      return (response.data as any).appointments || [];
    } catch (error) {
      this.logger.error(`Error fetching appointments by userId: ${error.message}`);
      if (error.response) {
        this.logger.error(`Appointments by userId API error status: ${error.response.status}`);
        this.logger.error(`
          : ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Get appointment slots (available times)
  async getAppointmentSlots(accessToken: string, calendarId: string, startDate: Date, endDate: Date, timezone: string = 'Europe/London'): Promise<any[]> {
    const url = `${this.baseUrl}/appointments/slots`;
    
    try {
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        params: {
          calendarId,
          startDate: Math.floor(startDate.getTime()),
          endDate: Math.floor(endDate.getTime()),
          timezone,
        },
        timeout: 10000,
      });
      
      this.logger.log(`Appointment slots API response status: ${response.status}`);
      return (response.data as any)._dates_?.slots || [];
    } catch (error) {
      this.logger.error(`Error fetching appointment slots: ${error.message}`);
      if (error.response) {
        this.logger.error(`Appointment slots API error status: ${error.response.status}`);
        this.logger.error(`Appointment slots API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Fetch appointments assigned to a specific team member
  async getAppointmentsByTeamMember(accessToken: string, locationId: string, teamMemberName: string): Promise<any[]> {
    try {
      const allAppointments = await this.getAppointments(accessToken, locationId);
      this.logger.log(`Fetched ${allAppointments.length} total appointments`);
      
      // Filter appointments by team member name
      const filteredAppointments = allAppointments.filter((appointment: any) => {
        const assignedTo = appointment.userId || appointment.assignedTo || appointment.teamMember || '';
        return assignedTo.toLowerCase().includes(teamMemberName.toLowerCase());
      });
      
      this.logger.log(`Found ${filteredAppointments.length} appointments assigned to ${teamMemberName}`);
      return filteredAppointments;
    } catch (error) {
      this.logger.error(`Error fetching appointments by team member: ${error.message}`);
      return [];
    }
  }

  // Fetch appointments for a specific contact ID
  async getAppointmentsByContactId(accessToken: string, locationId: string, contactId: string): Promise<any[]> {
    const cacheKey = this.getCacheKey('getAppointmentsByContactId', { contactId });
    const cached = this.getCachedData(cacheKey);
    if (cached) {
      return cached;
    }

    const url = `${this.baseUrl}/contacts/${contactId}/appointments/`;
    
    // Calculate date range (last 60 days to next 90 days)
    const now = new Date();
    const startDate = new Date(now.getTime() - (60 * 24 * 60 * 60 * 1000)); // 60 days ago
    const endDate = new Date(now.getTime() + (90 * 24 * 60 * 60 * 1000)); // 90 days from now
    
    try {
      const params = {
        startDate: Math.floor(startDate.getTime()),
        endDate: Math.floor(endDate.getTime()),
        includeAll: true, // Include contact and more data
        limit: 1000 // Add limit to get more appointments
      };

      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        params: params,
        timeout: 10000,
      });
      
      this.logger.log(`Appointments by contactId API response status: ${response.status}`);
      this.logger.log(`Appointments by contactId API response data keys: ${Object.keys(response.data || {}).join(', ')}`);
      
      // The API returns "events" instead of "appointments"
      const events = (response.data as any).events || [];
      this.logger.log(`Found ${events.length} events for contact ${contactId}`);
      
      // Log the first event structure for debugging
      if (events.length > 0) {
        this.logger.log(`First event structure: ${JSON.stringify(events[0])}`);
      }
      
      // Cache the result
      this.setCachedData(cacheKey, events);
      
      return events;
    } catch (error) {
      this.logger.error(`Error fetching appointments by contactId: ${error.message}`);
      if (error.response) {
        this.logger.error(`Appointments by contactId API error status: ${error.response.status}`);
        this.logger.error(`Appointments by contactId API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Alternative method: Get all appointments and filter by contactId
  async getAppointmentsByContactIdAlternative(accessToken: string, locationId: string, contactId: string): Promise<any[]> {
    try {
      // Get all appointments first
      const allAppointments = await this.getAppointments(accessToken, locationId);
      this.logger.log(`Fetched ${allAppointments.length} total appointments for filtering by contactId`);
      
      // Filter appointments by contactId
      const contactAppointments = allAppointments.filter((appointment: any) => {
        const appointmentContactId = appointment.contactId || appointment.contact?.id;
        return appointmentContactId === contactId;
      });
      
      this.logger.log(`Found ${contactAppointments.length} appointments for contactId ${contactId}`);
      
      // Log the first appointment structure for debugging
      if (contactAppointments.length > 0) {
        this.logger.log(`First appointment structure: ${JSON.stringify(contactAppointments[0])}`);
      }
      
      return contactAppointments;
    } catch (error) {
      this.logger.error(`Error in getAppointmentsByContactIdAlternative: ${error.message}`);
      return [];
    }
  }

  // Get user details by ID from GHL API
  async getUserById(accessToken: string, userId: string): Promise<any> {
    try {
      const url = `${this.baseUrl}/users/${userId}`;
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        timeout: 10000,
      });
      
      this.logger.log(`Successfully fetched user details for ID: ${userId}`);
      return response.data;
    } catch (error) {
      this.logger.error(`Error fetching user details for ID ${userId}: ${error.message}`);
      return null;
    }
  }

  // Get multiple users by IDs (batch request)
  async getUsersByIds(accessToken: string, userIds: string[]): Promise<Map<string, any>> {
    const userMap = new Map<string, any>();
    
    try {
      // Process in batches of 10 to avoid overwhelming the API
      const batchSize = 10;
      for (let i = 0; i < userIds.length; i += batchSize) {
        const batch = userIds.slice(i, i + batchSize);
        
        const promises = batch.map(async (userId) => {
          try {
            const user = await this.getUserById(accessToken, userId);
            if (user) {
              userMap.set(userId, user);
            }
          } catch (error) {
            this.logger.warn(`Failed to fetch user ${userId}: ${error.message}`);
          }
        });
        
        await Promise.all(promises);
        
        // Small delay between batches to be respectful to the API
        if (i + batchSize < userIds.length) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      }
      
      this.logger.log(`Successfully fetched ${userMap.size} users out of ${userIds.length} requested`);
      return userMap;
    } catch (error) {
      this.logger.error(`Error in batch user fetch: ${error.message}`);
      return userMap;
    }
  }

  // Get all users from GHL API
  async getAllUsers(accessToken: string, locationId: string): Promise<any[]> {
    try {
      const url = `${this.baseUrl}/users/`;
      const response = await axios.get(url, {
        headers: this.getHeaders(accessToken),
        params: {
          locationId: locationId,
          limit: 1000 // Get all users
        },
        timeout: 10000,
      });
      
      this.logger.log(`Successfully fetched ${(response.data as any)?.users?.length || 0} users from GHL`);
      return (response.data as any)?.users || [];
    } catch (error) {
      this.logger.error(`Error fetching all users: ${error.message}`);
      if (error.response) {
        this.logger.error(`Users API error status: ${error.response.status}`);
        this.logger.error(`Users API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  // Find user by name in GHL with improved precision
  async findUserByName(accessToken: string, locationId: string, userName: string): Promise<any> {
    try {
      const allUsers = await this.getAllUsers(accessToken, locationId);
      
      // First, try exact mappings for known problematic names
      const exactMappings = this.getExactUserMappings();
      if (exactMappings[userName]) {
        const mappedUser = allUsers.find((u: any) => u.id === exactMappings[userName]);
        if (mappedUser) {
          this.logger.log(`Found user via exact mapping: ${mappedUser.firstName} ${mappedUser.lastName} (ID: ${mappedUser.id})`);
          return mappedUser;
        }
      }
      
      // Try to find user by name with more precise matching
      const user = allUsers.find((u: any) => {
        const fullName = `${u.firstName || ''} ${u.lastName || ''}`.trim();
        const normalizedFullName = fullName.toLowerCase().replace(/\s+/g, ' ');
        const normalizedUserName = userName.toLowerCase().replace(/\s+/g, ' ');
        
        // Exact match first
        if (normalizedFullName === normalizedUserName) {
          return true;
        }
        
        // Check if the search name is contained in the full name (but be more careful)
        if (normalizedFullName.includes(normalizedUserName)) {
          // Additional validation: ensure it's not just a partial match that could be confused
          const nameParts = normalizedUserName.split(' ');
          if (nameParts.length >= 2) {
            // For multi-part names, ensure both first and last name parts match
            const firstName = u.firstName?.toLowerCase() || '';
            const lastName = u.lastName?.toLowerCase() || '';
            return nameParts.every(part => 
              firstName.includes(part) || lastName.includes(part)
            );
          }
          return true;
        }
        
        // Check individual name parts
        const firstName = u.firstName?.toLowerCase() || '';
        const lastName = u.lastName?.toLowerCase() || '';
        const nameParts = normalizedUserName.split(' ');
        
        return nameParts.every(part => 
          firstName.includes(part) || lastName.includes(part)
        );
      });
      
      if (user) {
        this.logger.log(`Found user in GHL: ${user.firstName} ${user.lastName} (ID: ${user.id})`);
        return user;
      } else {
        this.logger.warn(`User not found in GHL: ${userName}`);
        return null;
      }
    } catch (error) {
      this.logger.error(`Error finding user by name: ${error.message}`);
      return null;
    }
  }

  // Get exact user mappings for problematic names
  private getExactUserMappings(): Record<string, string> {
    return {
      'Robert Koch': 'er2hzTmwr4zgFoBpaaS1', // rob.koch -> Herrman RObert Koch
      'rob.koch': 'er2hzTmwr4zgFoBpaaS1',
      'Terrence Koch': '1H8Dos8NFvnEV3RzL4Y5', // terrence.koch -> Terrance koch  
      'terrence.koch': '1H8Dos8NFvnEV3RzL4Y5',
      // Add more mappings as needed
    };
  }

  // Get opportunities by pipeline ID with pagination and caching
  async getOpportunitiesByPipeline(accessToken: string, pipelineId: string, targetUserId?: string): Promise<any> {
    // Check cache first
    const cacheKey = this.getCacheKey('getOpportunitiesByPipeline', { pipelineId });
    const cached = this.getCachedData(cacheKey);
    if (cached) {
      this.logger.log(`Cache hit for pipeline ${pipelineId}, returning ${cached.opportunities?.length || 0} opportunities`);
      return cached;
    }

    // Check if there's already a request in progress for this pipeline
    const requestKey = `pipeline_${pipelineId}`;
    if (this.activeRequests && this.activeRequests.has(requestKey)) {
      this.logger.log(`Request already in progress for pipeline ${pipelineId}, waiting...`);
      // Wait for the existing request to complete
      return new Promise((resolve, reject) => {
        const checkInterval = setInterval(() => {
          if (!this.activeRequests || !this.activeRequests.has(requestKey)) {
            clearInterval(checkInterval);
            // Try to get from cache again
            const cachedResult = this.getCachedData(cacheKey);
            if (cachedResult) {
              resolve(cachedResult);
            } else {
              reject(new Error('Request completed but no cached result found'));
            }
          }
        }, 100);
      });
    }

    // Mark request as active
    if (!this.activeRequests) {
      this.activeRequests = new Set();
    }
    this.activeRequests.add(requestKey);

    const url = `${this.baseUrl}/pipelines/${pipelineId}/opportunities`;
    const maxRetries = 3;
    const baseDelay = 2000; // 2 seconds
    
    let allOpportunities: any[] = [];
    let userOpportunities: any[] = [];
    let hasMore = true;
    let startAfterId = null;
    let startAfter = null;
    const limit = 100; // GHL API limit per request
    
    this.logger.log(`Starting to fetch opportunities for pipeline ${pipelineId} with pagination${targetUserId ? ` (filtering for user: ${targetUserId})` : ''}`);
    
    try {
      // Fetch all pages
      while (hasMore && allOpportunities.length < 50000) { // Safety limit of 50,000
        // Store previous pagination values to detect if we're stuck
        const previousStartAfterId = startAfterId;
        const previousStartAfter = startAfter;
        
        for (let attempt = 1; attempt <= maxRetries; attempt++) {
          try {
            const params: any = {
              limit,
            };
            
            if (startAfterId) {
              params.startAfterId = startAfterId;
            }
            
            if (startAfter) {
              params.startAfter = startAfter;
            }
            
            this.logger.log(`Fetching opportunities for pipeline ${pipelineId} (page ${Math.floor(allOpportunities.length / limit) + 1}, attempt ${attempt}/${maxRetries})`);
            
            const response = await axios.get(url, {
              headers: this.getHeaders(accessToken),
              params,
              timeout: 15000,
            });

            if (response.status === 200 && response.data) {
              const opportunities = (response.data as any).opportunities || [];
              const meta = (response.data as any).meta || {};
              
              this.logger.log(`Pipeline ${pipelineId}: Received ${opportunities.length} opportunities (total so far: ${allOpportunities.length + opportunities.length})`);
              
              if (opportunities.length === 0) {
                hasMore = false;
              } else {
                allOpportunities = [...allOpportunities, ...opportunities];
                
                // If we're filtering for a specific user, check for user opportunities
                if (targetUserId) {
                  const pageUserOpportunities = opportunities.filter(opp => opp.assignedTo === targetUserId);
                  userOpportunities = [...userOpportunities, ...pageUserOpportunities];
                  
                  this.logger.log(`Page ${Math.floor(allOpportunities.length / limit)}: Found ${pageUserOpportunities.length} user opportunities (total user opportunities: ${userOpportunities.length})`);
                  
                  // Smart early termination: Stop when we're not finding many user opportunities per page
                  // This indicates we're either at the end of the user's opportunities or in a sparse section
                  const currentPage = Math.floor(allOpportunities.length / limit);
                  const userOpportunitiesPerPage = pageUserOpportunities.length;
                  
                  // If we've found a reasonable number of user opportunities and the current page has very few,
                  // we can stop (this handles users with 300+ opportunities efficiently)
                  if (userOpportunities.length >= 250 && userOpportunitiesPerPage <= 2 && currentPage >= 15) {
                    this.logger.log(`🚀 Smart early termination: Found ${userOpportunities.length} user opportunities after ${allOpportunities.length} total opportunities. Current page only has ${userOpportunitiesPerPage} user opportunities, stopping.`);
                    hasMore = false;
                    break;
                  }
                  
                  // Fallback: If we've found 300+ user opportunities and we're on page 20+, we can stop
                  if (userOpportunities.length >= 300 && allOpportunities.length >= 2000) {
                    this.logger.log(`🚀 Early termination: Found ${userOpportunities.length} user opportunities after ${allOpportunities.length} total opportunities. Stopping pagination.`);
                    hasMore = false;
                    break;
                  }
                }
                
                // Update pagination for next request
                startAfterId = meta.startAfterId;
                startAfter = meta.startAfter;
                
                // Check if pagination parameters haven't changed (indicating we're stuck)
                if (startAfterId === previousStartAfterId && startAfter === previousStartAfter && opportunities.length > 0) {
                  this.logger.warn(`Pipeline ${pipelineId}: Pagination parameters unchanged, breaking loop to prevent infinite loop`);
                  hasMore = false;
                  break;
                }
                
                // If no next page URL, we've reached the end
                if (!meta.nextPageUrl) {
                  hasMore = false;
                }
                
                // If we got fewer than the limit, we've reached the end
                if (opportunities.length < limit) {
                  hasMore = false;
                }
                
                // Additional safety check: if we're getting very few opportunities per page consistently, stop
                if (opportunities.length <= 1 && allOpportunities.length > 100) {
                  this.logger.warn(`Pipeline ${pipelineId}: Getting very few opportunities per page (${opportunities.length}), stopping to prevent infinite loop`);
                  hasMore = false;
                  break;
                }
              }
              
              // Break out of retry loop on success
              break;
            } else {
              throw new Error(`Unexpected response status: ${response.status}`);
            }
          } catch (error: any) {
            this.logger.error(`Attempt ${attempt} failed for pipeline ${pipelineId}:`, error.message);
            
            // Don't retry on authentication errors (401) - these won't succeed
            if (error.response?.status === 401) {
              this.logger.error(`Authentication error (401) - not retrying for pipeline ${pipelineId}`);
              throw error;
            }
            
            if (attempt === maxRetries) {
              this.logger.error(`All attempts failed for pipeline ${pipelineId}`);
              throw error;
            }
            
            // Exponential backoff
            const delay = baseDelay * Math.pow(2, attempt - 1);
            this.logger.log(`Retrying in ${delay}ms...`);
            await new Promise(resolve => setTimeout(resolve, delay));
          }
        }
        
        // Small delay between pages to be respectful to the API
        if (hasMore) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      }
    
      this.logger.log(`Successfully fetched ${allOpportunities.length} total opportunities for pipeline ${pipelineId}`);
      
      // Return the data in the same format as the original API response
      const result = {
        opportunities: targetUserId ? userOpportunities : allOpportunities,
        meta: {
          total: targetUserId ? userOpportunities.length : allOpportunities.length,
          pipelineId: pipelineId,
          filtered: !!targetUserId,
          targetUserId: targetUserId
        }
      };
      
      if (targetUserId) {
        this.logger.log(`🎯 Filtered results: ${userOpportunities.length} opportunities for user ${targetUserId} out of ${allOpportunities.length} total opportunities`);
      }
      
      // Cache the result
      this.setCachedData(cacheKey, result);
      
      // Remove from active requests
      if (this.activeRequests) {
        this.activeRequests.delete(requestKey);
      }
      
      return result;
    } catch (error) {
      // Clean up active request in case of any error
      if (this.activeRequests) {
        this.activeRequests.delete(requestKey);
      }
      throw error;
    }
  }

  /**
   * Update opportunity status and stage in GoHighLevel
   */
  async updateOpportunityStatus(
    accessToken: string, 
    pipelineId: string, 
    opportunityId: string, 
    status: 'open' | 'won' | 'lost' | 'abandoned',
    stageId?: string
  ): Promise<{ success: boolean; error?: string; data?: any }> {
    try {
      this.logger.log(`Updating opportunity ${opportunityId} status to ${status}${stageId ? ` and stage to ${stageId}` : ''}`);
      
      const url = `${this.baseUrl}/pipelines/${pipelineId}/opportunities/${opportunityId}/status`;
      
      const requestBody: any = {
        status: status
      };
      
      // Add stageId if provided
      if (stageId) {
        requestBody.stageId = stageId;
      }
      
      const response = await axios.put(url, requestBody, {
        headers: this.getHeaders(accessToken),
        timeout: 10000,
      });
      
      this.logger.log(`Successfully updated opportunity ${opportunityId} status to ${status}`);
      
      return {
        success: true,
        data: response.data
      };
    } catch (error: any) {
      this.logger.error(`Error updating opportunity status: ${error.message}`);
      if (error.response) {
        this.logger.error(`API error status: ${error.response.status}`);
        this.logger.error(`API error data: ${JSON.stringify(error.response.data)}`);
      }
      
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Update opportunity with full data (alternative method)
   */
  async updateOpportunity(
    accessToken: string,
    pipelineId: string,
    opportunityId: string,
    updateData: {
      title?: string;
      status?: 'open' | 'won' | 'lost' | 'abandoned';
      stageId?: string;
      monetaryValue?: number;
      assignedTo?: string;
      contactId?: string;
      email?: string;
      name?: string;
      phone?: string;
      tags?: string[];
      companyName?: string;
    }
  ): Promise<{ success: boolean; error?: string; data?: any }> {
    try {
      this.logger.log(`Updating opportunity ${opportunityId} with data:`, updateData);
      
      const url = `${this.baseUrl}/pipelines/${pipelineId}/opportunities/${opportunityId}`;
      
      const response = await axios.put(url, updateData, {
        headers: this.getHeaders(accessToken),
        timeout: 10000,
      });
      
      this.logger.log(`Successfully updated opportunity ${opportunityId}`);
      
      return {
        success: true,
        data: response.data
      };
    } catch (error: any) {
      this.logger.error(`Error updating opportunity: ${error.message}`);
      if (error.response) {
        this.logger.error(`API error status: ${error.response.status}`);
        this.logger.error(`API error data: ${JSON.stringify(error.response.data)}`);
      }
      
      return {
        success: false,
        error: error.message
      };
    }
  }
}
