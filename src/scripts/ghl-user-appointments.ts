import * as dotenv from 'dotenv';
import axios from 'axios';
import * as fs from 'fs';
import * as path from 'path';

dotenv.config({ override: true });

const BASE_URL = 'https://services.leadconnectorhq.com';
const VERSION = '2021-07-28';

function getMonthRange(month?: string): { startTime: number; endTime: number; label: string } {
  const now = new Date();
  const [yearStr, monthStr] = (month || '').split('-');
  const year = yearStr ? Number(yearStr) : now.getFullYear();
  const monthIndex = monthStr ? Number(monthStr) - 1 : now.getMonth();

  if (Number.isNaN(year) || Number.isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    throw new Error('GHL_EVENTS_MONTH must be in YYYY-MM format');
  }

  const start = new Date(year, monthIndex, 1, 0, 0, 0, 0);
  const end = new Date(year, monthIndex + 1, 0, 23, 59, 59, 999);
  const label = `${year}-${String(monthIndex + 1).padStart(2, '0')}`;

  return { startTime: start.getTime(), endTime: end.getTime(), label };
}

async function fetchUserAppointments() {
  const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
  const locationId = process.env.GHL_LOCATION_ID;
  const userId = process.env.GHL_USER_ID;
  const month = process.env.GHL_EVENTS_MONTH;

  if (!accessToken || !locationId || !userId) {
    console.error('Missing GOHIGHLEVEL_API_TOKEN, GHL_LOCATION_ID, or GHL_USER_ID in .env');
    process.exit(1);
  }

  let range;
  try {
    range = getMonthRange(month);
  } catch (error: any) {
    console.error(error.message);
    process.exit(1);
  }

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

    const outputDir = path.join(process.cwd(), 'tmp');
    const outputPath = path.join(outputDir, `ghl-user-appointments-${userId}-${range.label}.json`);
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, JSON.stringify(response.data, null, 2), 'utf8');

    const events = (response.data as any)?.events || [];
    console.log(`Saved ${events.length} event(s) to ${outputPath}`);
  } catch (error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    const code = error.code;
    const message = error.message;
    console.error('Failed to fetch user appointments.', { status, code, message, data });
    process.exit(1);
  }
}

fetchUserAppointments();
