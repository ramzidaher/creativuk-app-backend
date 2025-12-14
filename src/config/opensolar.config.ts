import { registerAs } from '@nestjs/config';

export default registerAs('opensolar', () => ({
  username: process.env.OPENSOLAR_USERNAME || 'ramzi@paldev.tech',
  password: process.env.OPENSOLAR_PASSWORD || 'pUH6WdNCC,ZUdKd',
  baseUrl: process.env.OPENSOLAR_BASE_URL || 'https://api.opensolar.com',
  apiVersion: process.env.OPENSOLAR_API_VERSION || 'v1',
}));

