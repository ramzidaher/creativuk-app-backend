/**
 * GHL API Latency Tracker
 * 
 * This script tracks the latency between when you make a change in the GoHighLevel
 * dashboard and when that change appears in the GHL API.
 * 
 * Usage:
 *   1. Run: npm run track:ghl-latency
 *   2. Enter your JWT token when prompted (or set JWT_TOKEN env var)
 *   3. The script will create a baseline snapshot
 *   4. Make your change in the GHL dashboard (add/update an opportunity)
 *   5. The script will detect when the change appears and report the latency
 * 
 * Environment Variables:
 *   - JWT_TOKEN: Your JWT authentication token (optional, will auto-login if not set)
 *   - USERNAME: Username for auto-login (default: ramzi.daher)
 *   - PASSWORD: Password for auto-login (default: Ramzi@2003)
 *   - API_BASE_URL: Base URL of your API (default: http://localhost:3000)
 *   - POLLING_INTERVAL: Polling interval in milliseconds (default: 3000 = 3 seconds)
 *   - MAX_WAIT_TIME: Maximum time to wait in milliseconds (default: 300000 = 5 minutes)
 * 
 * Example:
 *   API_BASE_URL=http://localhost:3000 npm run track:ghl-latency
 */

import axios from 'axios';
import * as readline from 'readline';

interface OpportunitySnapshot {
  id: string;
  contactId?: string;
  title?: string;
  appointmentDate?: string;
  appointmentTime?: string;
  tags?: string[];
  pipelineStageId?: string;
  updatedAt?: string;
  fullData?: any; // Store full opportunity data for detailed comparison
}

interface LatencyResult {
  detectedAt: Date;
  latencyMs: number;
  latencySeconds: number;
  latencyMinutes: number;
  changeType: 'new' | 'updated' | 'unknown';
  opportunityId: string;
  details: any;
}

class GHLLatencyTracker {
  private apiBaseUrl: string;
  private jwtToken: string;
  private pollingInterval: number; // milliseconds
  private maxWaitTime: number; // milliseconds
  private baselineSnapshot: Map<string, OpportunitySnapshot> = new Map();
  private startTime: Date | null = null;
  private isTracking: boolean = false;

  constructor(
    apiBaseUrl: string,
    jwtToken: string,
    pollingInterval: number = 3000, // 3 seconds default
    maxWaitTime: number = 300000 // 5 minutes default
  ) {
    this.apiBaseUrl = apiBaseUrl.replace(/\/$/, ''); // Remove trailing slash
    this.jwtToken = jwtToken;
    this.pollingInterval = pollingInterval;
    this.maxWaitTime = maxWaitTime;
  }

  /**
   * Fetch opportunities from the API endpoint
   */
  private async fetchOpportunities(): Promise<OpportunitySnapshot[]> {
    try {
      const response = await axios.get(
        `${this.apiBaseUrl}/api/opportunities/with-appointments-unified`,
        {
          headers: {
            Authorization: `Bearer ${this.jwtToken}`,
            'Content-Type': 'application/json',
          },
          timeout: 10000,
        }
      );

      const opportunities = response.data?.opportunities || [];
      return opportunities.map((opp: any) => ({
        id: opp.id,
        contactId: opp.contactId || opp.contact?.id,
        title: opp.title || opp.name,
        appointmentDate: opp.appointmentDate || opp.appointment?.date,
        appointmentTime: opp.appointmentTime || opp.appointment?.time,
        tags: opp.tags || [],
        pipelineStageId: opp.pipelineStageId || opp.stageId,
        updatedAt: opp.updatedAt || opp.dateUpdated,
        // Include full opportunity for detailed comparison
        fullData: opp,
      }));
    } catch (error: any) {
      if (error.response) {
        console.error(`❌ API Error: ${error.response.status} - ${error.response.statusText}`);
        if (error.response.data) {
          console.error('Response data:', JSON.stringify(error.response.data, null, 2));
        }
      } else {
        console.error(`❌ Network Error: ${error.message}`);
      }
      throw error;
    }
  }

  /**
   * Create a baseline snapshot of current opportunities
   */
  async createBaseline(): Promise<void> {
    console.log('📸 Creating baseline snapshot...');
    const opportunities = await this.fetchOpportunities();
    
    this.baselineSnapshot.clear();
    opportunities.forEach((opp) => {
      this.baselineSnapshot.set(opp.id, opp);
    });

    console.log(`✅ Baseline created: ${this.baselineSnapshot.size} opportunities`);
    console.log('📋 Opportunity IDs in baseline:');
    opportunities.forEach((opp) => {
      console.log(`   - ${opp.id}${opp.title ? ` (${opp.title})` : ''}`);
    });
  }

  /**
   * Compare current opportunities with baseline and detect changes
   */
  private detectChanges(current: OpportunitySnapshot[]): LatencyResult | null {
    const currentMap = new Map<string, OpportunitySnapshot>();
    current.forEach((opp) => currentMap.set(opp.id, opp));

    // Check for new opportunities
    for (const [id, currentOpp] of currentMap) {
      if (!this.baselineSnapshot.has(id)) {
        return {
          detectedAt: new Date(),
          latencyMs: this.startTime ? Date.now() - this.startTime.getTime() : 0,
          latencySeconds: this.startTime
            ? Math.round((Date.now() - this.startTime.getTime()) / 1000)
            : 0,
          latencyMinutes: this.startTime
            ? Math.round((Date.now() - this.startTime.getTime()) / 60000 * 100) / 100
            : 0,
          changeType: 'new',
          opportunityId: id,
          details: currentOpp.fullData || currentOpp,
        };
      }
    }

    // Check for updated opportunities (compare key fields)
    for (const [id, currentOpp] of currentMap) {
      const baselineOpp = this.baselineSnapshot.get(id);
      if (baselineOpp) {
        // Compare key fields that might change
        const fieldsChanged: string[] = [];

        if (currentOpp.appointmentDate !== baselineOpp.appointmentDate) {
          fieldsChanged.push('appointmentDate');
        }
        if (currentOpp.appointmentTime !== baselineOpp.appointmentTime) {
          fieldsChanged.push('appointmentTime');
        }
        if (JSON.stringify(currentOpp.tags) !== JSON.stringify(baselineOpp.tags)) {
          fieldsChanged.push('tags');
        }
        if (currentOpp.pipelineStageId !== baselineOpp.pipelineStageId) {
          fieldsChanged.push('pipelineStageId');
        }
        if (currentOpp.updatedAt !== baselineOpp.updatedAt) {
          fieldsChanged.push('updatedAt');
        }

        if (fieldsChanged.length > 0) {
          return {
            detectedAt: new Date(),
            latencyMs: this.startTime ? Date.now() - this.startTime.getTime() : 0,
            latencySeconds: this.startTime
              ? Math.round((Date.now() - this.startTime.getTime()) / 1000)
              : 0,
            latencyMinutes: this.startTime
              ? Math.round((Date.now() - this.startTime.getTime()) / 60000 * 100) / 100
              : 0,
            changeType: 'updated',
            opportunityId: id,
            details: {
              ...currentOpp.fullData,
              changedFields: fieldsChanged,
              baseline: baselineOpp,
              current: currentOpp,
            },
          };
        }
      }
    }

    return null;
  }

  /**
   * Start tracking and polling for changes
   */
  async startTracking(): Promise<LatencyResult | null> {
    if (this.isTracking) {
      console.log('⚠️  Already tracking. Stop current tracking first.');
      return null;
    }

    this.isTracking = true;
    this.startTime = new Date();

    console.log('\n🚀 Starting GHL API latency tracking...');
    console.log(`⏱️  Polling interval: ${this.pollingInterval / 1000} seconds`);
    console.log(`⏰ Maximum wait time: ${this.maxWaitTime / 1000} seconds`);
    console.log(`🕐 Started at: ${this.startTime.toISOString()}`);
    console.log('\n📝 Instructions:');
    console.log('   1. Make your change in the GoHighLevel dashboard NOW');
    console.log('   2. The script will detect when the change appears in the API');
    console.log('   3. Press Ctrl+C to stop tracking manually\n');

    const endTime = this.startTime.getTime() + this.maxWaitTime;
    let pollCount = 0;

    return new Promise((resolve, reject) => {
      const pollInterval = setInterval(async () => {
        try {
          pollCount++;
          const elapsed = this.startTime
            ? Math.round((Date.now() - this.startTime.getTime()) / 1000)
            : 0;

          // Check if max wait time exceeded
          if (Date.now() >= endTime) {
            clearInterval(pollInterval);
            this.isTracking = false;
            console.log(`\n⏰ Maximum wait time (${this.maxWaitTime / 1000}s) exceeded.`);
            console.log('❌ No changes detected.');
            resolve(null);
            return;
          }

          // Fetch current opportunities
          console.log(`\n🔍 Poll #${pollCount} (${elapsed}s elapsed)...`);
          const currentOpportunities = await this.fetchOpportunities();
          console.log(`   Found ${currentOpportunities.length} opportunities`);

          // Detect changes
          const change = this.detectChanges(currentOpportunities);
          if (change) {
            clearInterval(pollInterval);
            this.isTracking = false;
            this.reportResult(change);
            resolve(change);
            return;
          }

          console.log('   ✅ No changes detected yet...');
        } catch (error: any) {
          console.error(`   ❌ Error during poll: ${error.message}`);
          // Continue polling on error
        }
      }, this.pollingInterval);
    });
  }

  /**
   * Report the latency result
   */
  private reportResult(result: LatencyResult): void {
    console.log('\n' + '='.repeat(60));
    console.log('🎯 CHANGE DETECTED!');
    console.log('='.repeat(60));
    console.log(`📅 Detected at: ${result.detectedAt.toISOString()}`);
    console.log(`⏱️  Latency: ${result.latencyMs}ms`);
    console.log(`⏱️  Latency: ${result.latencySeconds} seconds`);
    console.log(`⏱️  Latency: ${result.latencyMinutes} minutes`);
    console.log(`🔄 Change type: ${result.changeType}`);
    console.log(`🆔 Opportunity ID: ${result.opportunityId}`);
    
    if (result.changeType === 'updated' && result.details.changedFields) {
      console.log(`📝 Changed fields: ${result.details.changedFields.join(', ')}`);
    }

    console.log('\n📊 Opportunity Details:');
    console.log(JSON.stringify(result.details, null, 2));
    console.log('='.repeat(60));
  }
}

/**
 * Login and get JWT token
 */
async function loginAndGetToken(apiBaseUrl: string, username: string, password: string): Promise<string> {
  try {
    console.log(`🔐 Logging in as ${username}...`);
    const response = await axios.post(
      `${apiBaseUrl}/api/auth/login`,
      {
        username,
        password,
      },
      {
        headers: {
          'Content-Type': 'application/json',
        },
        timeout: 10000,
      }
    );

    const accessToken = response.data?.accessToken;
    if (!accessToken) {
      throw new Error('No access token received from login');
    }

    console.log('✅ Login successful');
    return accessToken;
  } catch (error: any) {
    if (error.response) {
      console.error(`❌ Login failed: ${error.response.status} - ${error.response.statusText}`);
      if (error.response.data) {
        console.error('Response:', JSON.stringify(error.response.data, null, 2));
      }
    } else {
      console.error(`❌ Login error: ${error.message}`);
    }
    throw error;
  }
}

/**
 * Get JWT token from login or environment
 */
async function getJwtToken(apiBaseUrl: string): Promise<string> {
  // Check environment variable first
  const envToken = process.env.JWT_TOKEN;
  if (envToken) {
    console.log('✅ Using JWT token from JWT_TOKEN environment variable');
    return envToken;
  }

  // Try to use credentials from environment or defaults
  // Use GHL_USERNAME to avoid conflict with Windows USERNAME env var
  // Default credentials: ramzi.daher / Ramzi@2003
  const username = process.env.GHL_USERNAME || 'ramzi.daher';
  const password = process.env.GHL_PASSWORD || 'Ramzi@2003';

  // If credentials are provided, try to login
  if (username && password) {
    try {
      return await loginAndGetToken(apiBaseUrl, username, password);
    } catch (error) {
      console.error('❌ Auto-login failed. Please provide JWT_TOKEN or valid credentials.');
      process.exit(1);
    }
  }

  // Otherwise, prompt user for token
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve) => {
    rl.question('🔑 Enter your JWT token (or set JWT_TOKEN env var): ', (token) => {
      rl.close();
      if (!token.trim()) {
        console.error('❌ JWT token is required');
        process.exit(1);
      }
      resolve(token.trim());
    });
  });
}

/**
 * Get API base URL from environment or use default
 */
function getApiBaseUrl(): string {
  const envUrl = process.env.API_BASE_URL || process.env.BACKEND_URL;
  if (envUrl) {
    // Ensure URL doesn't have trailing slash, we'll add /api/ in the requests
    return envUrl.replace(/\/$/, '');
  }
  return 'https://creativuk-app.paldev.tech';
}

/**
 * Main function
 */
async function main() {
  console.log('🔍 GHL API Latency Tracker');
  console.log('='.repeat(60));

  const apiBaseUrl = getApiBaseUrl();
  console.log(`🌐 API Base URL: ${apiBaseUrl}`);

  const jwtToken = await getJwtToken(apiBaseUrl);
  const pollingInterval = parseInt(process.env.POLLING_INTERVAL || '3000', 10);
  const maxWaitTime = parseInt(process.env.MAX_WAIT_TIME || '300000', 10);

  const tracker = new GHLLatencyTracker(
    apiBaseUrl,
    jwtToken,
    pollingInterval,
    maxWaitTime
  );

  try {
    // Create baseline
    await tracker.createBaseline();

    // Wait a moment for user to make changes
    console.log('\n⏳ Waiting 5 seconds before starting to track...');
    console.log('   (Use this time to make your change in GHL dashboard)');
    await new Promise((resolve) => setTimeout(resolve, 5000));

    // Start tracking
    const result = await tracker.startTracking();

    if (!result) {
      console.log('\n❌ Tracking completed but no changes were detected.');
      process.exit(1);
    } else {
      console.log('\n✅ Tracking completed successfully!');
      process.exit(0);
    }
  } catch (error: any) {
    console.error('\n❌ Error:', error.message);
    if (error.stack) {
      console.error(error.stack);
    }
    process.exit(1);
  }
}

// Run the script if executed directly
if (require.main === module) {
  main().catch(console.error);
}

export { GHLLatencyTracker, main };



















