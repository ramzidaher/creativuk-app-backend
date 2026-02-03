import * as dotenv from 'dotenv';
import axios from 'axios';
import * as fs from 'fs';
import * as path from 'path';

dotenv.config({ override: true });

const BASE_URL = 'https://services.leadconnectorhq.com';
const VERSION = '2021-07-28';

async function fetchEventById() {
  const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
  const eventId = process.env.GHL_EVENT_ID;

  if (!accessToken || !eventId) {
    console.error('Missing GOHIGHLEVEL_API_TOKEN or GHL_EVENT_ID in .env');
    process.exit(1);
  }

  try {
    const response = await axios.get(
      `${BASE_URL}/calendars/events/appointments/${encodeURIComponent(eventId)}`,
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
    const outputPath = path.join(outputDir, `ghl-event-${eventId}.json`);
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, JSON.stringify(response.data, null, 2), 'utf8');

    console.log(`Saved event to ${outputPath}`);
  } catch (error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    const code = error.code;
    const message = error.message;
    console.error('Failed to fetch event.', { status, code, message, data });
    process.exit(1);
  }
}

fetchEventById();
