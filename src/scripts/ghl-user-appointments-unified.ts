import * as dotenv from 'dotenv';
import axios from 'axios';
import * as fs from 'fs';
import * as path from 'path';

dotenv.config({ override: true });

const BASE_URL = 'https://services.leadconnectorhq.com';
const VERSION = '2021-07-28';

function getRangeDays(daysBack = 30, daysForward = 30) {
  const now = new Date();
  const startDate = new Date(now.getTime() - daysBack * 24 * 60 * 60 * 1000);
  const endDate = new Date(now.getTime() + daysForward * 24 * 60 * 60 * 1000);
  return {
    startDate: Math.floor(startDate.getTime()),
    endDate: Math.floor(endDate.getTime()),
  };
}

async function fetchAppointmentsByUserId(accessToken: string, locationId: string, userId: string) {
  const { startDate, endDate } = getRangeDays();
  const url = `${BASE_URL}/appointments/`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        Version: VERSION,
      },
      params: {
        startDate,
        endDate,
        userId,
        includeAll: true,
        limit: 1000,
      },
      timeout: 10000,
    });
    return (response.data as any)?.appointments || [];
  } catch (error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    throw new Error(`UserId fetch failed (${status}): ${JSON.stringify(data)}`);
  }
}

async function fetchAppointmentsFallback(accessToken: string, locationId: string) {
  const { startDate, endDate } = getRangeDays();
  const url = `${BASE_URL}/appointments/`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      Version: VERSION,
    },
    params: {
      startDate,
      endDate,
      includeAll: true,
      limit: 1000,
    },
    timeout: 10000,
  });

  return (response.data as any)?.appointments || [];
}

async function fetchEventsByUserId(accessToken: string, locationId: string, userId: string) {
  const { startDate, endDate } = getRangeDays();
  const url = `${BASE_URL}/calendars/events`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      Version: VERSION,
    },
    params: {
      locationId,
      userId,
      startTime: startDate,
      endTime: endDate,
    },
    timeout: 10000,
  });

  return (response.data as any)?.events || [];
}

async function fetchUnifiedAppointments() {
  const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
  const locationId = process.env.GHL_LOCATION_ID;
  const ghlUserId = process.env.GHL_USER_ID;

  if (!accessToken || !locationId || !ghlUserId) {
    console.error('Missing GOHIGHLEVEL_API_TOKEN, GHL_LOCATION_ID, or GHL_USER_ID in .env');
    process.exit(1);
  }

  let appointments: any[] = [];
  let usedFallback = false;

  try {
    appointments = await fetchAppointmentsByUserId(accessToken, locationId, ghlUserId);
  } catch (error: any) {
    const status = error.response?.status;
    if (status === 404) {
      // v2 may not support /appointments; use calendar events instead
      usedFallback = true;
      appointments = await fetchEventsByUserId(accessToken, locationId, ghlUserId);
    } else {
      usedFallback = true;
      appointments = await fetchAppointmentsFallback(accessToken, locationId);
    }
  }

  const filtered = appointments.filter((appointment: any) => {
    const assignedUserId = appointment.userId || appointment.assignedTo || appointment.assignedUserId;
    const contactAssignedTo = appointment.contact?.assignedTo;

    if (assignedUserId) {
      return assignedUserId === ghlUserId;
    }
    return contactAssignedTo === ghlUserId;
  });

  const outputDir = path.join(process.cwd(), 'tmp');
  const outputPath = path.join(outputDir, `ghl-user-unified-${ghlUserId}.json`);
  fs.mkdirSync(outputDir, { recursive: true });
  fs.writeFileSync(
    outputPath,
    JSON.stringify({ usedFallback, total: filtered.length, appointments: filtered }, null, 2),
    'utf8'
  );

  console.log(
    `Saved ${filtered.length} appointment(s) to ${outputPath} (fallback: ${usedFallback})`
  );
}

fetchUnifiedAppointments();
