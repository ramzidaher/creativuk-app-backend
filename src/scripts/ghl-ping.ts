import * as dotenv from 'dotenv';
import axios from 'axios';

dotenv.config({ override: true });

const BASE_URL = 'https://services.leadconnectorhq.com';
const VERSION = '2021-07-28';

async function ghlPing() {
  const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
  const locationId = process.env.GHL_LOCATION_ID;

  if (!accessToken || !locationId) {
    console.error('Missing GOHIGHLEVEL_API_TOKEN or GHL_LOCATION_ID in .env');
    process.exit(1);
  }

  const tokenPrefix = accessToken.slice(0, 4);
  const tokenSuffix = accessToken.slice(-4);
  console.log(`Using locationId: ${locationId}`);
  console.log(`Using token: ${tokenPrefix}...${tokenSuffix}`);

  try {
    const response = await axios.get(`${BASE_URL}/locations/${locationId}`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        Version: VERSION,
      },
      timeout: 10000,
    });

    const location = (response.data as any)?.location;
    console.log(`GHL ping OK. Location: ${location?.name || locationId}`);
  } catch (error: any) {
    const status = error.response?.status;
    const data = error.response?.data;
    console.error('GHL ping failed.', { status, data });
    process.exit(1);
  }
}

ghlPing();
