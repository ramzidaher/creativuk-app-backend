import { Controller, Get, Query, Res, UseGuards, Request } from '@nestjs/common';
import { Response } from 'express';
import { GhlAuthService } from './ghl-auth.service';
import { JwtAuthGuard } from '../../auth/jwt-auth.guard';

@Controller('creativ-crm')
export class GhlAuthController {
  constructor(private readonly ghlAuthService: GhlAuthService) {}

  @Get('callback')
  async ghlCallback(@Query('code') code: string, @Query('test') test: string, @Res() res: Response) {
    try {
      console.log('OAuth callback: Received code:', code ? 'present' : 'missing');
      console.log('OAuth callback: Test parameter:', test);
      
      const jwt = await this.ghlAuthService.handleOAuthCallback(code);
      console.log('OAuth callback: JWT generated successfully');
      
      // Check if this is a mobile app request
      const userAgent = res.req.headers['user-agent'] || '';
      const isMobileApp = userAgent.includes('Expo') || userAgent.includes('ReactNative') || test === 'mobile' || test === '1';
      
      console.log('OAuth callback: User agent:', userAgent);
      console.log('OAuth callback: Is mobile app:', isMobileApp);
      
      let redirectUrl: string;
      
      if (isMobileApp) {
        // For mobile app, redirect to the mobile app's OAuth callback URL
        // Try different mobile app URLs
        const mobileAppUrl = process.env.MOBILE_APP_URL || 'exp://localhost:8081';
        redirectUrl = `${mobileAppUrl}/oauth-callback?token=${jwt}`;
        
        // Also try the web URL with mobile parameter for better compatibility
        const webUrl = process.env.FRONTEND_URL || 'https://0686fe28fffc.ngrok-free.app';
        const webRedirectUrl = `${webUrl}/oauth-callback?token=${jwt}&mobile=1`;
        
        console.log('OAuth callback: Mobile redirect URL:', redirectUrl);
        console.log('OAuth callback: Web redirect URL (fallback):', webRedirectUrl);
        
        // For now, use the web URL as it's more reliable
        redirectUrl = webRedirectUrl;
      } else {
        // For web app, redirect to the frontend URL
        const frontendUrl = process.env.FRONTEND_URL || 'https://0686fe28fffc.ngrok-free.app';
        redirectUrl = `${frontendUrl}/oauth-callback?token=${jwt}`;
        console.log('OAuth callback: Web redirect URL:', redirectUrl);
      }
      
      console.log(`OAuth callback: Final redirect URL: ${redirectUrl}`);
      return res.redirect(redirectUrl);
    } catch (error) {
      console.error('OAuth callback error:', error);
      const errorUrl = process.env.FRONTEND_URL || 'https://0686fe28fffc.ngrok-free.app';
      return res.redirect(`${errorUrl}/oauth-callback?error=${encodeURIComponent(error.message)}`);
    }
  }

  @Get('refresh-token')
  @UseGuards(JwtAuthGuard)
  async refreshToken(@Request() req) {
    try {
      const { ghlUserId } = req.user;
      const newJwt = await this.ghlAuthService.forceTokenRefresh(ghlUserId);
      return { success: true, token: newJwt };
    } catch (error) {
      console.error('Token refresh error:', error);
      return { success: false, error: error.message };
    }
  }

  @Get('force-reauth')
  @UseGuards(JwtAuthGuard)
  async forceReauth(@Request() req, @Res() res: Response) {
    try {
      const { ghlUserId } = req.user;
      console.log(`Force re-authentication for user: ${ghlUserId}`);
      
      // Redirect to OAuth flow to get new token with updated scopes
      const clientId = process.env.GHL_CLIENT_ID!;
      const redirectUri = process.env.GHL_REDIRECT_URI!;
      const scopes = 'businesses.readonly,calendars.readonly,businesses.write,users.readonly,locations.readonly,opportunities.readonly,contacts.readonly';
      
      const authUrl = `https://marketplace.gohighlevel.com/oauth/chooselocation?response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&client_id=${clientId}&scope=${encodeURIComponent(scopes)}&state=force_reauth`;
      
      console.log(`Redirecting to OAuth with updated scopes: ${authUrl}`);
      return res.redirect(authUrl);
    } catch (error) {
      console.error('Force reauth error:', error);
      return res.status(500).json({ success: false, error: error.message });
    }
  }
} 