import { Injectable, Logger } from '@nestjs/common';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { ConfigService } from '@nestjs/config';
import { PrismaService } from '../prisma/prisma.service';

export interface SurveyorArea {
  name: string;
  location: string;
  areas: string[];
  maxTravelTime: string;
  ghlUserId?: string;
  ghlUserData?: any;
}

@Injectable()
export class DynamicSurveyorService {
  private readonly logger = new Logger(DynamicSurveyorService.name);
  private readonly loggedNotFoundSurveyors = new Set<string>();
  
  // No hardcoded surveyor areas - all data comes from database and GHL
  

  constructor(
    private readonly goHighLevelService: GoHighLevelService,
    private readonly configService: ConfigService,
    private readonly prisma: PrismaService
  ) {}

  private getGhlCredentials() {
    const accessToken = this.configService.get<string>('GOHIGHLEVEL_API_TOKEN');
    
    if (!accessToken) {
      this.logger.warn('GHL API token not configured - falling back to default surveyors');
      this.logger.warn('To add Jennifer Garrett and other GHL users, configure GOHIGHLEVEL_API_TOKEN');
      return null;
    }
    
    // Extract location ID from the JWT token (same as opportunities service)
    try {
      const tokenData = this.extractTokenData(accessToken);
      const locationId = tokenData.locationId;
      
      if (!locationId) {
        this.logger.warn('Location ID not found in GHL token - falling back to default surveyors');
        return null;
      }
      
      return { accessToken, locationId };
    } catch (error) {
      this.logger.warn('Error extracting location ID from GHL token - falling back to default surveyors');
      return null;
    }
  }

  private extractTokenData(token: string) {
    try {
      // JWT tokens have 3 parts separated by dots
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
      this.logger.error('Error decoding JWT token:', error);
      throw error;
    }
  }

  /**
   * Get all surveyors (fetches from database and enhances with GoHighLevel data)
   */
  async getAllSurveyors(): Promise<SurveyorArea[]> {
    try {
      // Get all users from database who have surveyor areas configured
      const dbUsers = await this.prisma.user.findMany({
        where: {
          role: 'SURVEYOR',
          surveyorAreas: {
            isEmpty: false
          }
        },
        select: {
          id: true,
          username: true,
          name: true,
          ghlUserId: true,
          ghlUserName: true,
          surveyorAreas: true,
          surveyorLocation: true,
          maxTravelTime: true
        }
      });

      this.logger.log(`Found ${dbUsers.length} surveyors in database: ${dbUsers.map(u => u.name || u.username).join(', ')}`);

      // Convert database users to SurveyorArea format
      const surveyors: SurveyorArea[] = dbUsers.map(user => ({
        name: user.name || user.username,
        location: user.surveyorLocation || 'TBD',
        areas: user.surveyorAreas || [],
        maxTravelTime: user.maxTravelTime || '9:30am to 6:30pm',
        ghlUserId: user.ghlUserId || undefined,
        ghlUserData: undefined // Will be populated if GHL is available
      }));

      // Try to enhance with GoHighLevel data if available
      try {
        const credentials = this.getGhlCredentials();
        if (credentials) {
          const ghlUsers = await this.goHighLevelService.getAllUsers(credentials.accessToken, credentials.locationId);
          const userMap = new Map(ghlUsers.map((user: any) => [user.id, user]));
          
          // Update surveyors with GHL data where possible
          for (const surveyor of surveyors) {
            if (surveyor.ghlUserId && userMap.has(surveyor.ghlUserId)) {
              surveyor.ghlUserData = userMap.get(surveyor.ghlUserId);
              this.logger.log(`✅ Enhanced surveyor ${surveyor.name} with GHL data`);
            }
          }
        }
      } catch (ghlError) {
        this.logger.warn('Could not fetch GHL data, using database data only:', ghlError.message);
      }

      return surveyors;
    } catch (error) {
      this.logger.error('Error getting surveyors from database:', error);
      return [];
    }
  }

  /**
   * Get surveyor by name (searches both default and GoHighLevel users)
   */
  async getSurveyorByName(name: string): Promise<SurveyorArea | undefined> {
    const surveyors = await this.getAllSurveyors();
    return surveyors.find(surveyor => this.namesMatch(surveyor.name, name));
  }

  /**
   * Get surveyor by GHL user ID
   */
  async getSurveyorByGhlUserId(ghlUserId: string): Promise<SurveyorArea | undefined> {
    const surveyors = await this.getAllSurveyors();
    return surveyors.find(surveyor => surveyor.ghlUserId === ghlUserId);
  }

  /**
   * Check if opportunity is assigned to a surveyor
   * PRIMARY METHOD: Match by GHL User ID (most reliable)
   * FALLBACK: Match by name only if GHL User ID is not available
   */
  async isOpportunityAssignedToSurveyor(
    opportunity: any, 
    surveyorName: string, 
    userMap?: Map<string, any>
  ): Promise<boolean> {
    const surveyor = await this.getSurveyorByName(surveyorName);
    if (!surveyor) {
      // Only log the error once per surveyor name to avoid spam
      if (!this.loggedNotFoundSurveyors.has(surveyorName)) {
        this.logger.log(`Surveyor not found for name: ${surveyorName}`);
        this.loggedNotFoundSurveyors.add(surveyorName);
      }
      return false;
    }

    this.logger.log(`Checking opportunity: ${opportunity.name} for surveyor: ${surveyor.name} (GHL ID: ${surveyor.ghlUserId})`);

    // PRIMARY: Check if opportunity is assigned to this surveyor by GHL User ID (most reliable)
    const assignedToId = opportunity.assignedTo;
    if (assignedToId && surveyor.ghlUserId) {
      if (assignedToId === surveyor.ghlUserId) {
        this.logger.log(`✅ Match found by GHL User ID: ${assignedToId} === ${surveyor.ghlUserId}`);
        return true;
      } else {
        this.logger.log(`❌ No match by GHL User ID: ${assignedToId} !== ${surveyor.ghlUserId}`);
        // If we have both IDs and they don't match, this is definitely not assigned to this surveyor
        return false;
      }
    }

    // FALLBACK: Only if GHL User ID is not available, check by name
    if (!assignedToId || !surveyor.ghlUserId) {
      this.logger.log(`⚠️ Missing GHL User ID data - falling back to name matching`);
      
      // Check if opportunity is assigned to this surveyor by name
      const assignedToName = opportunity.teamMember || opportunity.assignee || opportunity.assignedToName;
      if (assignedToName) {
        this.logger.log(`Checking assignedTo name: ${assignedToName} against surveyor: ${surveyor.name}`);
        
        if (this.namesMatch(assignedToName, surveyor.name)) {
          this.logger.log(`✅ Match found by assignment name: ${assignedToName}`);
          return true;
        } else {
          this.logger.log(`❌ No match by assignment name: ${assignedToName} !== ${surveyor.name}`);
        }
      } else {
        this.logger.log(`❌ No name fields found for assignment check`);
      }
    }

    this.logger.log(`❌ No match found for opportunity - returning false`);
    return false;
  }


  /**
   * Helper method to check if two names match (handles various formats)
   */
  private namesMatch(name1: string, name2: string): boolean {
    if (!name1 || !name2) return false;
    
    const name1Lower = name1.toLowerCase().trim();
    const name2Lower = name2.toLowerCase().trim();
    
    // Direct match
    if (name1Lower === name2Lower) return true;
    
    // Handle common name variations
    const nameVariations = {
      'robert': ['rob', 'bob', 'robert', 'hermann'],
      'rob': ['robert', 'bob', 'rob', 'hermann'],
      'bob': ['robert', 'rob', 'bob', 'hermann'],
      'hermann': ['robert', 'rob', 'bob', 'hermann'],
      'terrence': ['terry', 'terrence'],
      'terry': ['terrence', 'terry'],
      'michael': ['mike', 'mick', 'michael'],
      'mike': ['michael', 'mick', 'mike'],
      'mick': ['michael', 'mike', 'mick'],
      'james': ['jim', 'jimmy', 'james'],
      'jim': ['james', 'jimmy', 'jim'],
      'jimmy': ['james', 'jim', 'jimmy'],
      'william': ['bill', 'will', 'william'],
      'bill': ['william', 'will', 'bill'],
      'will': ['william', 'bill', 'will'],
      'richard': ['rick', 'dick', 'richard'],
      'rick': ['richard', 'dick', 'rick'],
      'dick': ['richard', 'rick', 'dick']
    };
    
    // Check for name variations
    for (const [baseName, variations] of Object.entries(nameVariations)) {
      if (name1Lower.includes(baseName) && variations.some(v => name2Lower.includes(v))) return true;
      if (name2Lower.includes(baseName) && variations.some(v => name1Lower.includes(v))) return true;
    }
    
    // Contains match
    if (name1Lower.includes(name2Lower) || name2Lower.includes(name1Lower)) return true;
    
    // Check reversed order (e.g., "John Smith" vs "Smith John")
    const name1Parts = name1Lower.split(' ').filter(p => p.length > 0);
    const name2Parts = name2Lower.split(' ').filter(p => p.length > 0);
    
    if (name1Parts.length === 2 && name2Parts.length === 2) {
      const reversedName1 = `${name1Parts[1]} ${name1Parts[0]}`;
      const reversedName2 = `${name2Parts[1]} ${name2Parts[0]}`;
      
      if (reversedName1 === name2Lower || reversedName2 === name1Lower) return true;
    }
    
    // Partial match (any word matches)
    for (const word1 of name1Parts) {
      for (const word2 of name2Parts) {
        if (word1.length > 2 && word2.length > 2 && word1 === word2) {
          return true;
        }
      }
    }
    
    return false;
  }

  /**
   * Get surveyors that need area configuration (users without configured areas)
   */
  async getSurveyorsNeedingConfiguration(): Promise<SurveyorArea[]> {
    const surveyors = await this.getAllSurveyors();
    return surveyors.filter(s => s.areas.length === 0 || s.areas.includes('TBD') || s.location === 'TBD');
  }
}
