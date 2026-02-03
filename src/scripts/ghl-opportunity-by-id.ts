import * as dotenv from 'dotenv';
import axios from 'axios';
import * as fs from 'fs';
import * as path from 'path';

dotenv.config({ override: true });

const BASE_URL = 'https://services.leadconnectorhq.com';
const VERSION = '2021-07-28';

async function fetchOpportunityById() {
  const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
  const locationId = process.env.GHL_LOCATION_ID;
  const opportunityId = process.env.GHL_OPPORTUNITY_ID;

  if (!accessToken || !locationId || !opportunityId) {
    console.error('Missing GOHIGHLEVEL_API_TOKEN, GHL_LOCATION_ID, or GHL_OPPORTUNITY_ID in .env');
    process.exit(1);
  }

  try {
    const pipelinesRes = await axios.get(`${BASE_URL}/pipelines/`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        Version: VERSION,
      },
      timeout: 10000,
    });

    const pipelines = (pipelinesRes.data as any)?.pipelines || [];
    if (!pipelines.length) {
      console.error('No pipelines returned for location.');
      process.exit(1);
    }

    for (const pipeline of pipelines) {
      try {
        const response = await axios.get(
          `${BASE_URL}/pipelines/${pipeline.id}/opportunities/${encodeURIComponent(opportunityId)}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: 'application/json',
              Version: VERSION,
            },
            timeout: 10000,
          }
        );

        const outputDir = path.join(process.cwd(), 'tmp');
        const outputPath = path.join(outputDir, `ghl-opportunity-${opportunityId}.json`);
        fs.mkdirSync(outputDir, { recursive: true });
        fs.writeFileSync(outputPath, JSON.stringify(response.data, null, 2), 'utf8');

        console.log(`Found opportunity in pipeline ${pipeline.id}. Saved to ${outputPath}`);
        return;
      } catch (error: any) {
        const status = error.response?.status;
        if (status === 404) {
          continue;
        }
        if (status === 401) {
          throw error;
        }
        continue;
      }
    }

    console.error('Opportunity not found in any pipeline.');
    process.exit(1);
  } catch (error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    const code = error.code;
    const message = error.message;
    console.error('Failed to fetch opportunity.', { status, code, message, data });
    process.exit(1);
  }
}

fetchOpportunityById();
