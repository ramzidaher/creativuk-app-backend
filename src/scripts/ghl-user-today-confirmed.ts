import * as dotenv from 'dotenv';
import axios from 'axios';
import * as fs from 'fs';
import * as path from 'path';

dotenv.config({ override: true });

const BASE_URL = 'https://services.leadconnectorhq.com';
const VERSION = '2021-07-28';

function getTodayRange(): { startTime: number; endTime: number; label: string } {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  const label = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
    now.getDate()
  ).padStart(2, '0')}`;

  return { startTime: start.getTime(), endTime: end.getTime(), label };
}

async function fetchTodayConfirmedAppointments() {
  const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
  const locationId = process.env.GHL_LOCATION_ID;
  const userId = process.env.GHL_USER_ID || 'VAv7oPFnEiukCXR70vn6';

  if (!accessToken || !locationId) {
    console.error('Missing GOHIGHLEVEL_API_TOKEN or GHL_LOCATION_ID in .env');
    process.exit(1);
  }

  const range = getTodayRange();

  try {
    const response = await axios.get(`${BASE_URL}/calendars/events`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        Version: VERSION,
      },
      params: {
        locationId,
        userId,
        startTime: range.startTime,
        endTime: range.endTime,
      },
      timeout: 15000,
    });

    const events = (response.data as any)?.events || [];
    const confirmed = events.filter((event: any) => event.appointmentStatus === 'confirmed');

    const outputDir = path.join(process.cwd(), 'tmp');
    const outputPath = path.join(outputDir, `ghl-user-confirmed-${userId}-${range.label}.json`);
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, JSON.stringify({ events: confirmed }, null, 2), 'utf8');

    console.log(`Saved ${confirmed.length} confirmed appointment(s) to ${outputPath}`);
  } catch (error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    const code = error.code;
    const message = error.message;
    console.error('Failed to fetch today confirmed appointments.', { status, code, message, data });
    process.exit(1);
  }
}

fetchTodayConfirmedAppointments();
