import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import axios from 'axios';

export interface GHLUser {
  id: string;
  firstName: string;
  lastName: string;
  email: string;
}

export interface GHLUserLookupResult {
  found: boolean;
  ghlUser?: GHLUser;
  message: string;
}

@Injectable()
export class GHLUserLookupService {
  private readonly logger = new Logger(GHLUserLookupService.name);
  private readonly GHL_BASE_URL = 'https://services.leadconnectorhq.com';

  constructor(private readonly configService: ConfigService) {}

  /**
   * Get GHL API credentials from environment
   */
  private getGHLCredentials(): { accessToken: string; locationId: string } | null {
    const accessToken = this.configService.get<string>('GOHIGHLEVEL_API_TOKEN') || 
                       this.configService.get<string>('GHL_ACCESS_TOKEN');
    const locationId = this.configService.get<string>('GHL_LOCATION_ID');

    if (!accessToken || !locationId) {
      this.logger.warn('GHL credentials not configured. GHL user lookup will be skipped.');
      return null;
    }

    return { accessToken, locationId };
  }

  /**
   * Fetch all users from GHL API
   */
  async getAllGHLUsers(): Promise<GHLUser[]> {
    const credentials = this.getGHLCredentials();
    if (!credentials) {
      return [];
    }

    try {
      const url = `${this.GHL_BASE_URL}/users/`;
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${credentials.accessToken}`,
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          Version: '2021-07-28',
        },
        params: {
          locationId: credentials.locationId,
          limit: 1000
        },
        timeout: 10000,
      });
      
      const users = (response.data as any)?.users || [];
      this.logger.log(`Successfully fetched ${users.length} users from GHL`);
      return users;
    } catch (error) {
      this.logger.error(`Error fetching GHL users: ${error.message}`);
      if (error.response) {
        this.logger.error(`GHL API error status: ${error.response.status}`);
        this.logger.error(`GHL API error data: ${JSON.stringify(error.response.data)}`);
      }
      return [];
    }
  }

  /**
   * Find GHL user by name (case insensitive)
   */
  async findGHLUserByName(userName: string): Promise<GHLUserLookupResult> {
    if (!userName || userName.trim().length === 0) {
      return {
        found: false,
        message: 'User name is required for GHL lookup'
      };
    }

    const ghlUsers = await this.getAllGHLUsers();
    if (ghlUsers.length === 0) {
      return {
        found: false,
        message: 'No GHL users found or GHL API not accessible'
      };
    }

    // Try to find user by name (case insensitive)
    const normalizedUserName = userName.toLowerCase().trim();
    
    const ghlUser = ghlUsers.find((u: GHLUser) => {
      const fullName = `${u.firstName || ''} ${u.lastName || ''}`.trim().toLowerCase();
      const firstName = (u.firstName || '').toLowerCase();
      const lastName = (u.lastName || '').toLowerCase();
      
      return fullName.includes(normalizedUserName) ||
             firstName.includes(normalizedUserName) ||
             lastName.includes(normalizedUserName) ||
             normalizedUserName.includes(firstName) ||
             normalizedUserName.includes(lastName);
    });
    
    if (ghlUser) {
      this.logger.log(`Found GHL user: ${ghlUser.firstName} ${ghlUser.lastName} (ID: ${ghlUser.id}) for name: ${userName}`);
      return {
        found: true,
        ghlUser,
        message: `Found matching GHL user: ${ghlUser.firstName} ${ghlUser.lastName}`
      };
    } else {
      this.logger.log(`No GHL user found for name: ${userName}`);
      return {
        found: false,
        message: `No matching GHL user found for name: ${userName}`
      };
    }
  }

  /**
   * Find GHL user by email
   */
  async findGHLUserByEmail(email: string): Promise<GHLUserLookupResult> {
    if (!email || email.trim().length === 0) {
      return {
        found: false,
        message: 'Email is required for GHL lookup'
      };
    }

    const ghlUsers = await this.getAllGHLUsers();
    if (ghlUsers.length === 0) {
      return {
        found: false,
        message: 'No GHL users found or GHL API not accessible'
      };
    }

    const normalizedEmail = email.toLowerCase().trim();
    const ghlUser = ghlUsers.find((u: GHLUser) => 
      (u.email || '').toLowerCase() === normalizedEmail
    );
    
    if (ghlUser) {
      this.logger.log(`Found GHL user by email: ${ghlUser.firstName} ${ghlUser.lastName} (ID: ${ghlUser.id}) for email: ${email}`);
      return {
        found: true,
        ghlUser,
        message: `Found matching GHL user by email: ${ghlUser.firstName} ${ghlUser.lastName}`
      };
    } else {
      this.logger.log(`No GHL user found for email: ${email}`);
      return {
        found: false,
        message: `No matching GHL user found for email: ${email}`
      };
    }
  }

  /**
   * Try to find GHL user by name or email
   */
  async findGHLUser(userName: string, email?: string): Promise<GHLUserLookupResult> {
    // First try by name
    const nameResult = await this.findGHLUserByName(userName);
    if (nameResult.found) {
      return nameResult;
    }

    // If email provided, try by email
    if (email) {
      const emailResult = await this.findGHLUserByEmail(email);
      if (emailResult.found) {
        return emailResult;
      }
    }

    // Return the name result (which will have the "not found" message)
    return nameResult;
  }

  /**
   * Get all available GHL users for manual assignment reference
   */
  async getAvailableGHLUsers(): Promise<GHLUser[]> {
    return this.getAllGHLUsers();
  }
}

