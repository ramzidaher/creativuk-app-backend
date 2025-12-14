import { Injectable } from '@nestjs/common';
import { UserService } from '../../user/user.service';
import { JwtService } from '@nestjs/jwt';
import axios from 'axios';

@Injectable()
export class GhlAuthService {
  constructor(
    private userService: UserService,
    private jwtService: JwtService,
  ) {}

  async handleOAuthCallback(code: string): Promise<string> {
    // 1. Exchange code for tokens (use x-www-form-urlencoded)
    const params = new URLSearchParams();
    params.append('grant_type', 'authorization_code');
    params.append('code', code);
    params.append('client_id', process.env.GHL_CLIENT_ID!);
    params.append('client_secret', process.env.GHL_CLIENT_SECRET!);
    params.append('redirect_uri', process.env.GHL_REDIRECT_URI!);

    const tokenRes = await axios.post(
      'https://services.leadconnectorhq.com/oauth/token',
      params.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );
    const { access_token, refresh_token, expires_in } = tokenRes.data as {
      access_token: string;
      refresh_token: string;
      expires_in: number;
    };

    // Log the full token response and scopes
    console.log('Token response:', tokenRes.data);
    console.log('Access token:', access_token);
    console.log('Scopes:', (tokenRes.data as any).scope);

    // 2. Fetch GHL user info using /users/search
    const companyId = (tokenRes.data as any).companyId;
    const locationId = (tokenRes.data as any).locationId;
    const userId = (tokenRes.data as any).userId;
    let usersRes;
    try {
      usersRes = await axios.get('https://services.leadconnectorhq.com/users/search', {
        headers: {
          Authorization: `Bearer ${access_token}`,
          Version: '2021-07-28',
        },
        params: {
          companyId,
          locationId,
          limit: 25, // get more users to ensure we find the matching one
        },
      });
      console.log('GHL /users/search response:', usersRes.data);
    } catch (err) {
      console.error('Error fetching /users/search:', err.response?.data || err.message);
      throw new Error('Failed to fetch users from GHL /users/search');
    }

    const users = usersRes.data.users || [];
    const currentUser = users.find((u: any) => u.id === userId) || users[0];
    if (!currentUser) {
      throw new Error('No users found in GHL /users/search response');
    }

    // 3. Upsert user in DB
    const user = await this.userService.upsertByGhlUserId({
      ghlUserId: currentUser.id,
      name: currentUser.name,
      email: currentUser.email,
      ghlAccessToken: access_token,
      ghlRefreshToken: refresh_token,
      tokenExpiresAt: new Date(Date.now() + expires_in * 1000),
    });

    // 4. Issue JWT
    const payload = { sub: user.id, ghlUserId: user.ghlUserId, email: user.email };
    return this.jwtService.sign(payload);
  }

  async forceTokenRefresh(ghlUserId: string): Promise<string> {
    console.log(`Force refreshing token for user: ${ghlUserId}`);
    
    // Get the current user
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new Error('User not found');
    }

    // Check if we have a refresh token
    if (!user.ghlRefreshToken) {
      throw new Error('No refresh token available. User needs to re-authenticate.');
    }

    // Use refresh token to get new access token
    const params = new URLSearchParams();
    params.append('grant_type', 'refresh_token');
    params.append('refresh_token', user.ghlRefreshToken);
    params.append('client_id', process.env.GHL_CLIENT_ID!);
    params.append('client_secret', process.env.GHL_CLIENT_SECRET!);

    try {
      const tokenRes = await axios.post(
        'https://services.leadconnectorhq.com/oauth/token',
        params.toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );

      const { access_token, refresh_token, expires_in } = tokenRes.data as {
        access_token: string;
        refresh_token: string;
        expires_in: number;
      };

      console.log('Token refresh response:', tokenRes.data);
      console.log('New scopes:', (tokenRes.data as any).scope);

      // Update user with new tokens
      const updatedUser = await this.userService.upsertByGhlUserId({
        ghlUserId: user.ghlUserId,
        name: user.name,
        email: user.email,
        ghlAccessToken: access_token,
        ghlRefreshToken: refresh_token,
        tokenExpiresAt: new Date(Date.now() + expires_in * 1000),
      });

      // Issue new JWT
      const payload = { sub: updatedUser.id, ghlUserId: updatedUser.ghlUserId, email: updatedUser.email };
      return this.jwtService.sign(payload);
    } catch (error) {
      console.error('Token refresh failed:', error.response?.data || error.message);
      throw new Error('Token refresh failed. User needs to re-authenticate.');
    }
  }
} 






