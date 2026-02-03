import * as dotenv from 'dotenv';
import axios from 'axios';
import * as fs from 'fs';
import * as path from 'path';

dotenv.config({ override: true });

const BASE_URL = 'https://services.leadconnectorhq.com';
const VERSION = '2021-07-28';

async function fetchCalendars() {
  const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
  const locationId = process.env.GHL_LOCATION_ID;

  if (!accessToken || !locationId) {
    console.error('Missing GOHIGHLEVEL_API_TOKEN or GHL_LOCATION_ID in .env');
    process.exit(1);
  }

  try {
    const response = await axios.get(`${BASE_URL}/calendars/`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        Version: VERSION,
      },
      params: {
        locationId,
      },
      timeout: 10000,
    });

    const outputDir = path.join(process.cwd(), 'tmp');
    const outputPath = path.join(outputDir, 'ghl-calendars.json');
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, JSON.stringify(response.data, null, 2), 'utf8');

    const calendars = (response.data as any)?.calendars || [];
    console.log(`Saved ${calendars.length} calendar(s) to ${outputPath}`);
  } catch (error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    const code = error.code;
    const message = error.message;
    console.error('Failed to fetch calendars.', { status, code, message, data });
    process.exit(1);
  }
}

fetchCalendars();
