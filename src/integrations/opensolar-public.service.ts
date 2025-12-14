import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import axios from 'axios';

export interface OpenSolarPublicProject {
  id: number;
  address?: string;
  display_name?: string;
  name?: string;
  created_at?: string;
  updated_at?: string;
}

export interface OpenSolarPublicDesign {
  id: string;
  name: string;
  system_type: 'solar' | 'battery' | 'hybrid';
  panels?: OpenSolarPublicPanel[];
  arrays?: OpenSolarPublicArray[];
  batteries?: OpenSolarPublicBattery[];
  inverters?: OpenSolarPublicInverter[];
  orientation?: {
    tilt?: number;
    azimuth?: number;
    face?: string;
  };
  shading?: {
    annual_loss?: number;
    monthly_loss?: number[];
  };
}

export interface OpenSolarPublicPanel {
  model: string;
  count: number;
  watt_per_module?: number;
  dc_size_kw?: number;
  manufacturer?: string;
}

export interface OpenSolarPublicArray {
  id: string;
  name?: string;
  panel_count: number;
  panel_model: string;
  orientation: {
    tilt?: number;
    azimuth?: number;
    face?: string;
  };
  shading?: {
    annual_loss?: number;
    monthly_loss?: number[];
  };
}

export interface OpenSolarPublicBattery {
  manufacturer: string;
  model: string;
  capacity?: number;
  voltage?: number;
}

export interface OpenSolarPublicInverter {
  manufacturer: string;
  model: string;
  type: 'solar' | 'battery' | 'hybrid';
  capacity?: number;
}

export interface CreateProjectDto {
  name: string;
  address: string;
  postcode: string;
  customer_name?: string;
  customer_email?: string;
  customer_phone?: string;
}

export interface CreateDesignDto {
  projectId: number;
  name: string;
  systemType: 'solar' | 'battery' | 'hybrid';
  panels?: {
    model: string;
    count: number;
    watt_per_module?: number;
    manufacturer?: string;
  }[];
  arrays?: {
    name?: string;
    panel_count: number;
    panel_model: string;
    orientation?: {
      tilt?: number;
      azimuth?: number;
      face?: string;
    };
    shading?: {
      annual_loss?: number;
      monthly_loss?: number[];
    };
  }[];
  batteries?: {
    manufacturer: string;
    model: string;
    capacity?: number;
    voltage?: number;
  }[];
  inverters?: {
    manufacturer: string;
    model: string;
    type: 'solar' | 'battery' | 'hybrid';
    capacity?: number;
  }[];
  orientation?: {
    tilt?: number;
    azimuth?: number;
    face?: string;
  };
  shading?: {
    annual_loss?: number;
    monthly_loss?: number[];
  };
}

// API Response interfaces
export interface OpenSolarAuthResponse {
  token: string;
  user?: {
    org_id?: number;
    organization_id?: number;
  };
  org_id?: number;
  organization_id?: number;
}

export interface OpenSolarUserResponse {
  org_id: number;
  [key: string]: any;
}

export interface PostcodesIOResponse {
  result: {
    latitude: number;
    longitude: number;
  };
}

export interface OpenStreetMapResponse {
  lat: string;
  lon: string;
  [key: string]: any;
}

@Injectable()
export class OpenSolarPublicService {
  private readonly logger = new Logger(OpenSolarPublicService.name);
  private readonly baseUrl = 'https://api.opensolar.com';
  private readonly username: string;
  private readonly password: string;
  private authToken: string | null = null;
  private orgId: number | null = null;

  constructor(private readonly configService: ConfigService) {
    this.username = this.configService.get<string>('opensolar.username') || 'ramzi@paldev.tech';
    this.password = this.configService.get<string>('opensolar.password') || 'pUH6WdNCC,ZUdKd';
  }

  /**
   * Authenticate with OpenSolar API
   */
  private async authenticate(): Promise<void> {
    if (this.authToken && this.orgId) {
      return; // Already authenticated
    }

    try {
      this.logger.log('🔐 Authenticating with OpenSolar...');
      this.logger.log(`🔍 Using base URL: ${this.baseUrl}`);
      this.logger.log(`🔍 Username: ${this.username}`);
      
      const response = await axios.post<OpenSolarAuthResponse>(`${this.baseUrl}/api-token-auth/`, {
        username: this.username,
        password: this.password
      });

      this.logger.log(`🔍 Authentication response status: ${response.status}`);
      this.logger.log(`🔍 Authentication response data:`, JSON.stringify(response.data, null, 2));
      this.logger.log(`🔍 Authentication response headers:`, JSON.stringify(response.headers, null, 2));

      if (response.data && response.data.token) {
        this.authToken = response.data.token;
        
        // Try to extract org_id from different possible locations
        let orgId: number | null = null;
        if (response.data.user?.org_id) {
          orgId = response.data.user.org_id;
          this.logger.log(`✅ Found org_id in user object: ${orgId}`);
        } else if (response.data.org_id) {
          orgId = response.data.org_id;
          this.logger.log(`✅ Found org_id in root: ${orgId}`);
        } else if (response.data.user?.organization_id) {
          orgId = response.data.user.organization_id;
          this.logger.log(`✅ Found organization_id in user object: ${orgId}`);
        } else if (response.data.organization_id) {
          orgId = response.data.organization_id;
          this.logger.log(`✅ Found organization_id in root: ${orgId}`);
        } else {
          this.logger.log(`⚠️ No organization ID found in authentication response, will try to get it separately`);
        }
        
        this.orgId = orgId;
        this.logger.log(`✅ OpenSolar authentication successful | token=${this.authToken?.substring(0, 10) || 'unknown'}... | orgId=${orgId || 'will fetch separately'}`);
        
        // Test the token by making a simple API call
        await this.testApiConnection();
      } else {
        throw new Error('Invalid authentication response - no token found');
      }
    } catch (error: any) {
      this.logger.error('❌ OpenSolar authentication failed:', error.message);
      if (error.response) {
        this.logger.error('❌ Response status:', error.response.status);
        this.logger.error('❌ Response data:', error.response.data);
        this.logger.error('❌ Response headers:', error.response.headers);
      }
      throw new Error(`OpenSolar authentication failed: ${error.message}`);
    }
  }
  
  /**
   * Test API connection with the authenticated token
   */
  private async testApiConnection(): Promise<void> {
    try {
      this.logger.log('🧪 Testing API connection...');
      
      // Test the connection by trying to list projects (more reliable than /api/user/)
      if (this.orgId) {
        try {
          const projectsResponse = await axios.get(`${this.baseUrl}/api/orgs/${this.orgId}/projects/`, {
            headers: {
              'Authorization': `Bearer ${this.authToken}`,
              'Content-Type': 'application/json'
            }
          });
          this.logger.log(`✅ API connection test successful: ${projectsResponse.status}`);
          this.logger.log(`✅ Projects endpoint accessible: ${projectsResponse.status}`);
          this.logger.log(`🔍 Available projects:`, projectsResponse.data);
        } catch (projectsError: any) {
          this.logger.log(`⚠️ Projects endpoint test failed: ${projectsError.response?.status || 'unknown error'}`);
          // If projects endpoint fails, try a simpler test
          this.logger.log('🔄 Trying alternative connection test...');
          
          // Try to access the org endpoint directly
          try {
            const orgResponse = await axios.get(`${this.baseUrl}/api/orgs/${this.orgId}/`, {
              headers: {
                'Authorization': `Bearer ${this.authToken}`,
                'Content-Type': 'application/json'
              }
            });
            this.logger.log(`✅ Alternative connection test successful: ${orgResponse.status}`);
            this.logger.log(`🔍 Org data:`, orgResponse.data);
          } catch (orgError: any) {
            this.logger.error(`❌ Alternative connection test also failed: ${orgError.response?.status || 'unknown error'}`);
            throw orgError;
          }
        }
      } else {
        this.logger.warn('⚠️ No org ID available for connection test');
      }
    } catch (error: any) {
      this.logger.error('❌ API connection test failed:', error.message);
      if (error.response) {
        this.logger.error('❌ Test response status:', error.response.status);
        this.logger.error('❌ Test response data:', error.response.data);
      }
    }
  }

  /**
   * Get organization ID if not available
   */
  private async getOrgId(): Promise<number> {
    if (this.orgId) {
      return this.orgId;
    }

    try {
      this.logger.log(`🔍 Attempting to get organization ID, auth token: ${this.authToken ? 'present' : 'missing'}`);
      
      // Try to get user info to extract org_id
      const response = await axios.get<OpenSolarUserResponse>(`${this.baseUrl}/api/user/`, {
        headers: {
          'Authorization': `Bearer ${this.authToken}`,
          'Content-Type': 'application/json'
        }
      });

      this.logger.log(`🔍 User API response status: ${response.status}`);
      this.logger.log(`🔍 User API response data:`, response.data);

      if (response.data && response.data.org_id) {
        const orgId = response.data.org_id;
        this.orgId = orgId;
        this.logger.log(`✅ Organization ID retrieved: ${orgId}`);
        return orgId;
      } else {
        this.logger.error('❌ No org_id found in user response');
        this.logger.error('❌ Available fields:', Object.keys(response.data || {}));
        throw new Error('Could not determine organization ID from user response');
      }
    } catch (error: any) {
      this.logger.error('❌ Error getting organization ID:', error.message);
      if (error.response) {
        this.logger.error('❌ Response status:', error.response.status);
        this.logger.error('❌ Response data:', error.response.data);
      }
      throw new Error(`Failed to get organization ID: ${error.message}`);
    }
  }

  /**
   * Create a new OpenSolar project
   */
  async createProject(projectData: CreateProjectDto): Promise<OpenSolarPublicProject> {
    try {
      await this.authenticate();
      const orgId = await this.getOrgId();
      
      this.logger.log(`🏗️ Creating OpenSolar project: ${projectData.name}`);
      this.logger.log(`🔍 Using org ID: ${orgId}`);
      this.logger.log(`🔍 Auth token present: ${this.authToken ? 'yes' : 'no'}`);
      
      // Geocode the address to get coordinates
      const coordinates = await this.geocodeAddress(`${projectData.address}, ${projectData.postcode}`);
      this.logger.log(`📍 Geocoded coordinates: ${coordinates.lat}, ${coordinates.lng}`);
      
      // Use the actual customer email from the opportunity data
      const customerEmail = projectData.customer_email;
      const customerName = projectData.customer_name || 'Guest User';
      
      this.logger.log(`🔍 Using customer email: ${customerEmail} for customer: ${customerName}`);

      const projectPayload = {
        title: projectData.name,
        address: `${projectData.address}, ${projectData.postcode}`,
        lat: coordinates.lat,
        lon: coordinates.lng,
        is_residential: true,
        notes: `Project created from CreativSolar app for ${customerName}`,
        contacts_new: [
          {
            first_name: customerName.split(' ')[0] || 'Guest',
            family_name: customerName.split(' ').slice(1).join(' ') || 'User',
            email: customerEmail,
            phone: projectData.customer_phone || '',
            gender: 0 // 0 = unset
          }
        ]
      };
      
      this.logger.log(`🔍 Project payload:`, projectPayload);
      
      // Try multiple possible endpoints based on OpenSolar API documentation
      let response;
      let endpointUsed = '';
      
      // Based on OpenSolar API documentation, try these endpoints in order
      const endpoints = [
        `${this.baseUrl}/api/orgs/${orgId}/projects/`,  // Primary endpoint from docs
        `${this.baseUrl}/api/projects/`,                 // Alternative endpoint
        `${this.baseUrl}/api/v1/orgs/${orgId}/projects/`, // Versioned endpoint
        `${this.baseUrl}/api/v1/projects/`,              // Versioned alternative
        `${this.baseUrl}/projects/`,                     // Simple endpoint
        `${this.baseUrl}/orgs/${orgId}/projects/`,       // Simple org endpoint
        // Try alternative base URLs that might work
        `https://app.opensolar.com/api/orgs/${orgId}/projects/`,
        `https://app.opensolar.com/api/projects/`,
        `https://app.opensolar.com/api/v1/orgs/${orgId}/projects/`
      ];
      
      for (const endpoint of endpoints) {
        try {
          this.logger.log(`🔍 Trying endpoint: ${endpoint}`);
          
          // Try with different payload structures based on OpenSolar API documentation
          const payloads = [
            projectPayload, // Original payload
            {
              ...projectPayload,
              name: projectPayload.title, // Try 'name' instead of 'title'
              title: undefined
            },
            {
              ...projectPayload,
              site_address: projectPayload.address, // Try 'site_address' instead of 'address'
              address: undefined
            },
            // Try minimal payload structure
            {
              name: projectPayload.title,
              address: projectPayload.address,
              lat: projectPayload.lat,
              lon: projectPayload.lon,
              is_residential: projectPayload.is_residential
            },
            // Try with different field names
            {
              title: projectPayload.title,
              site_address: projectPayload.address,
              latitude: projectPayload.lat,
              longitude: projectPayload.lon,
              residential: projectPayload.is_residential,
              notes: projectPayload.notes
            },
            // Try with contacts as separate array
            {
              title: projectPayload.title,
              address: projectPayload.address,
              lat: projectPayload.lat,
              lon: projectPayload.lon,
              is_residential: projectPayload.is_residential,
              notes: projectPayload.notes,
              contacts: projectPayload.contacts_new
            }
          ];
          
          for (const payload of payloads) {
            try {
              this.logger.log(`🔍 Trying payload variant:`, payload);
              response = await axios.post<OpenSolarPublicProject>(endpoint, payload, {
                headers: {
                  'Authorization': `Bearer ${this.authToken}`,
                  'Content-Type': 'application/json'
                }
              });
              endpointUsed = endpoint;
              this.logger.log(`✅ Success with endpoint: ${endpoint} and payload variant`);
              break;
                                  } catch (payloadError: any) {
              const status = payloadError.response?.status;
              this.logger.log(`⚠️ Payload variant failed: ${status || 'unknown error'}`);
              
              if (status === 400) {
                this.logger.log(`🔍 400 Bad Request - payload format issue, trying next variant`);
                this.logger.log(`🔍 Error details:`, payloadError.response?.data);
                this.logger.log(`🔍 Error headers:`, payloadError.response?.headers);
                continue; // Try next payload variant
              } else if (status === 401) {
                this.logger.log(`🔍 401 Unauthorized - authentication issue`);
                this.logger.log(`🔍 Error details:`, payloadError.response?.data);
                break; // Don't try more payloads for this endpoint
              } else if (status === 404) {
                this.logger.log(`🔍 404 Not Found - endpoint doesn't exist`);
                this.logger.log(`🔍 Error details:`, payloadError.response?.data);
                break; // Don't try more payloads for this endpoint
              } else {
                this.logger.log(`🔍 Unexpected error status: ${status}, trying next variant`);
                this.logger.log(`🔍 Error details:`, payloadError.response?.data);
                continue;
              }
            }
          }
          
          if (response) break; // If we got a successful response, stop trying endpoints
          
        } catch (error: any) {
          this.logger.log(`⚠️ Endpoint ${endpoint} failed: ${error.response?.status || 'unknown error'}`);
          if (error.response?.status === 401) {
            this.logger.log(`🔍 401 Unauthorized - might be wrong endpoint or token issue`);
          } else if (error.response?.status === 404) {
            this.logger.log(`🔍 404 Not Found - endpoint doesn't exist`);
          }
          
          // If this is the last endpoint, throw the error
          if (endpoint === endpoints[endpoints.length - 1]) {
            throw error;
          }
        }
      }

      this.logger.log(`🔍 Project creation response status: ${response.status}`);
      this.logger.log(`🔍 Project creation response data:`, response.data);

      if (response.data && response.data.id) {
        this.logger.log(`✅ OpenSolar project created successfully: ${response.data.id}`);
        return response.data;
      } else {
        this.logger.error('❌ Invalid project creation response - no ID found');
        this.logger.error('❌ Response data:', response.data);
        throw new Error('Invalid project creation response - no ID found');
      }
    } catch (error: any) {
      this.logger.error('❌ Error creating OpenSolar project:', error.message);
      if (error.response) {
        this.logger.error('❌ Response status:', error.response.status);
        this.logger.error('❌ Response data:', error.response.data);
        this.logger.error('❌ Response headers:', error.response.headers);
      }
      throw new Error(`Failed to create OpenSolar project: ${error.message}`);
    }
  }

  /**
   * Create a new design for an OpenSolar project
   */
  async createDesign(designData: CreateDesignDto): Promise<OpenSolarPublicDesign> {
    try {
      await this.authenticate();
      const orgId = await this.getOrgId();
      
      this.logger.log(`🎨 Creating design for project ${designData.projectId}: ${designData.name}`);
      
      // Create the system design
      const systemData = {
        name: designData.name,
        display_name: designData.name,
        system_type: designData.systemType,
        panels: designData.panels || [],
        arrays: designData.arrays || [],
        batteries: designData.batteries || [],
        inverters: designData.inverters || [],
        orientation: designData.orientation || {},
        shading: designData.shading || {},
        status: 'draft'
      };

      const response = await axios.post<OpenSolarPublicDesign>(
        `${this.baseUrl}/api/orgs/${orgId}/projects/${designData.projectId}/systems/`,
        systemData,
        {
          headers: {
            'Authorization': `Bearer ${this.authToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.data && response.data.id) {
        this.logger.log(`✅ OpenSolar design created successfully: ${response.data.id}`);
        return response.data;
      } else {
        throw new Error('Invalid design creation response');
      }
    } catch (error: any) {
      this.logger.error('❌ Error creating OpenSolar design:', error.message);
      throw new Error(`Failed to create OpenSolar design: ${error.message}`);
    }
  }

  /**
   * Get project details by ID
   */
  async getProject(projectId: number): Promise<OpenSolarPublicProject> {
    try {
      await this.authenticate();
      const orgId = await this.getOrgId();
      
      this.logger.log(`📋 Fetching OpenSolar project #${projectId}...`);
      
      const response = await axios.get<OpenSolarPublicProject>(
        `${this.baseUrl}/api/orgs/${orgId}/projects/${projectId}/`,
        {
          headers: {
            'Authorization': `Bearer ${this.authToken}`
          }
        }
      );

      if (response.data) {
        this.logger.log(`✅ OpenSolar project fetched successfully`);
        return response.data;
      } else {
        throw new Error('Invalid project response');
      }
    } catch (error: any) {
      this.logger.error(`❌ Error fetching OpenSolar project #${projectId}:`, error.message);
      throw new Error(`Failed to fetch OpenSolar project: ${error.message}`);
    }
  }

  /**
   * Get design details by project ID and design ID
   */
  async getDesign(projectId: number, designId: string): Promise<OpenSolarPublicDesign> {
    try {
      await this.authenticate();
      
      this.logger.log(`🎨 Fetching OpenSolar design #${designId} for project #${projectId}...`);
      
      const response = await axios.get<OpenSolarPublicDesign>(
        `${this.baseUrl}/api/orgs/${this.orgId}/projects/${projectId}/systems/${designId}/`,
        {
          headers: {
            'Authorization': `Bearer ${this.authToken}`
          }
        }
      );

      if (response.data) {
        this.logger.log(`✅ OpenSolar design fetched successfully`);
        return response.data;
      } else {
        throw new Error('Invalid design response');
      }
    } catch (error: any) {
      this.logger.error(`❌ Error fetching OpenSolar design #${designId}:`, error.message);
      throw new Error(`Failed to fetch OpenSolar design: ${error.message}`);
    }
  }

  /**
   * List all designs for a project
   */
  async listDesigns(projectId: number): Promise<OpenSolarPublicDesign[]> {
    try {
      await this.authenticate();
      
      this.logger.log(`📋 Fetching designs for OpenSolar project #${projectId}...`);
      
      const response = await axios.get<OpenSolarPublicDesign[]>(
        `${this.baseUrl}/api/orgs/${this.orgId}/projects/${projectId}/systems/`,
        {
          headers: {
            'Authorization': `Bearer ${this.authToken}`
          }
        }
      );

      if (response.data && Array.isArray(response.data)) {
        this.logger.log(`✅ Found ${response.data.length} designs for project #${projectId}`);
        return response.data;
      } else {
        this.logger.log(`✅ No designs found for project #${projectId}`);
        return [];
      }
    } catch (error: any) {
      this.logger.error(`❌ Error fetching designs for project #${projectId}:`, error.message);
      throw new Error(`Failed to fetch designs: ${error.message}`);
    }
  }

  /**
   * Update an existing design
   */
  async updateDesign(projectId: number, designId: string, designData: Partial<CreateDesignDto>): Promise<OpenSolarPublicDesign> {
    try {
      await this.authenticate();
      
      this.logger.log(`✏️ Updating OpenSolar design #${designId} for project #${projectId}...`);
      
      const response = await axios.patch<OpenSolarPublicDesign>(
        `${this.baseUrl}/api/orgs/${this.orgId}/projects/${projectId}/systems/${designId}/`,
        designData,
        {
          headers: {
            'Authorization': `Bearer ${this.authToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.data && response.data.id) {
        this.logger.log(`✅ OpenSolar design updated successfully: ${response.data.id}`);
        return response.data;
      } else {
        throw new Error('Invalid design update response');
      }
    } catch (error: any) {
      this.logger.error(`❌ Error updating OpenSolar design #${designId}:`, error.message);
      throw new Error(`Failed to update OpenSolar design: ${error.message}`);
    }
  }

  /**
   * Delete a design
   */
  async deleteDesign(projectId: number, designId: string): Promise<void> {
    try {
      await this.authenticate();
      
      this.logger.log(`🗑️ Deleting OpenSolar design #${designId} from project #${projectId}...`);
      
      await axios.delete(
        `${this.baseUrl}/api/orgs/${this.orgId}/projects/${projectId}/systems/${designId}/`,
        {
          headers: {
            'Authorization': `Bearer ${this.authToken}`
          }
        }
      );

      this.logger.log(`✅ OpenSolar design deleted successfully: ${designId}`);
    } catch (error: any) {
      this.logger.error(`❌ Error deleting OpenSolar design #${designId}:`, error.message);
      throw new Error(`Failed to delete OpenSolar design: ${error.message}`);
    }
  }

  /**
   * Get project URL for viewing in OpenSolar
   */
  getProjectUrl(projectId: number): string {
    return `https://app.opensolar.com/projects/${projectId}`;
  }

  /**
   * Get design URL for viewing in OpenSolar
   */
  getDesignUrl(projectId: number, designId: string): string {
    return `https://app.opensolar.com/projects/${projectId}/systems/${designId}`;
  }

  /**
   * Get authentication token for frontend use
   */
  async getAuthToken(): Promise<{ token: string; orgId: number }> {
    try {
      await this.authenticate();
      const orgId = await this.getOrgId();
      
      if (!this.authToken) {
        throw new Error('No authentication token available');
      }
      
      return {
        token: this.authToken,
        orgId: orgId
      };
    } catch (error: any) {
      this.logger.error('❌ Error getting auth token:', error.message);
      throw new Error(`Failed to get auth token: ${error.message}`);
    }
  }

  /**
   * Geocode an address to get coordinates
   */
  private async geocodeAddress(address: string): Promise<{ lat: number; lng: number }> {
    try {
      this.logger.log(`🌍 Geocoding address: ${address}`);
      
      // Try UK postcode lookup first (more accurate for UK addresses)
      if (address.includes(',')) {
        const postcode = address.split(',').pop()?.trim();
        if (postcode && this.isUKPostcode(postcode)) {
          this.logger.log(`🇬🇧 UK postcode detected: ${postcode}`);
          const coordinates = await this.geocodeUKPostcode(postcode);
          if (coordinates) {
            return coordinates;
          }
        }
      }
      
      // Fallback to OpenStreetMap geocoding
      return await this.geocodeWithOpenStreetMap(address);
    } catch (error: any) {
      this.logger.error('❌ Geocoding failed:', error.message);
      // Return default coordinates (London) as fallback
      return { lat: 51.5074, lng: -0.1278 };
    }
  }
  
  /**
   * Check if a string looks like a UK postcode
   */
  private isUKPostcode(postcode: string): boolean {
    const ukPostcodeRegex = /^[A-Z]{1,2}[0-9][A-Z0-9]? ?[0-9][A-Z]{2}$/i;
    return ukPostcodeRegex.test(postcode);
  }
  
  /**
   * Geocode UK postcode using postcodes.io API
   */
  private async geocodeUKPostcode(postcode: string): Promise<{ lat: number; lng: number } | null> {
    try {
      const response = await axios.get<PostcodesIOResponse>(`https://api.postcodes.io/postcodes/${encodeURIComponent(postcode)}`);
      
      if (response.data && response.data.result) {
        const { latitude, longitude } = response.data.result;
        this.logger.log(`✅ UK postcode geocoded: ${latitude}, ${longitude}`);
        return { lat: latitude, lng: longitude };
      }
    } catch (error: any) {
      this.logger.log(`⚠️ UK postcode geocoding failed: ${error.message}`);
    }
    return null;
  }
  
  /**
   * Geocode address using OpenStreetMap Nominatim API
   */
  private async geocodeWithOpenStreetMap(address: string): Promise<{ lat: number; lng: number }> {
    try {
      const encodedAddress = encodeURIComponent(address);
      const response = await axios.get<OpenStreetMapResponse[]>(
        `https://nominatim.openstreetmap.org/search?q=${encodedAddress}&format=json&limit=1&countrycodes=gb`,
        {
          headers: {
            'User-Agent': 'CreativSolar/1.0'
          }
        }
      );
      
      if (response.data && response.data.length > 0) {
        const result = response.data[0];
        const lat = parseFloat(result.lat);
        const lng = parseFloat(result.lon);
        
        this.logger.log(`✅ OpenStreetMap geocoded: ${lat}, ${lng}`);
        return { lat, lng };
      }
      
      throw new Error('No geocoding results found');
    } catch (error: any) {
      this.logger.error('❌ OpenStreetMap geocoding failed:', error.message);
      throw error;
    }
  }

  /**
   * Get authenticated OpenSolar URL for a project
   */
  getAuthenticatedProjectUrl(projectId: number, token: string): string {
    // OpenSolar web app doesn't support direct token passing in URL
    // Instead, we'll return the project URL and the user will need to be logged in
    return `https://app.opensolar.com/projects/${projectId}`;
  }
}
