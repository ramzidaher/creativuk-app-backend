import { Controller, Get, Query, Res } from '@nestjs/common';
import { Response } from 'express';

@Controller('oauth')
export class OAuthController {
  
  @Get('callback')
  async handleOAuthCallback(
    @Query('code') code: string,
    @Query('state') state: string,
    @Query('error') error: string,
    @Query('error_description') errorDescription: string,
    @Res() res: Response
  ) {
    console.log('🔍 OAuth Callback received:', { code, state, error, errorDescription });
    
    if (error) {
      // Handle OAuth error
      return res.send(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>OAuth Error</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
            .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            .error { color: #e74c3c; background: #fdf2f2; padding: 15px; border-radius: 5px; margin: 20px 0; }
            .code { background: #f8f9fa; padding: 10px; border-radius: 5px; font-family: monospace; word-break: break-all; }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>❌ OAuth Authorization Failed</h1>
            <div class="error">
              <strong>Error:</strong> ${error}<br>
              <strong>Description:</strong> ${errorDescription || 'No description provided'}
            </div>
            <p>Please try again or contact support if the problem persists.</p>
          </div>
        </body>
        </html>
      `);
    }
    
    if (!code) {
      return res.send(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>OAuth Callback</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
            .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            .success { color: #27ae60; background: #f0f9f0; padding: 15px; border-radius: 5px; margin: 20px 0; }
            .code { background: #f8f9fa; padding: 15px; border-radius: 5px; font-family: monospace; word-break: break-all; margin: 20px 0; }
            .copy-btn { background: #3498db; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; margin: 10px 0; }
            .copy-btn:hover { background: #2980b9; }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>⚠️ No Authorization Code</h1>
            <p>The OAuth callback was received but no authorization code was provided.</p>
            <p>Please try the authorization process again.</p>
          </div>
        </body>
        </html>
      `);
    }
    
    // Success - show the authorization code
    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>OAuth Success</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
          .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          .success { color: #27ae60; background: #f0f9f0; padding: 15px; border-radius: 5px; margin: 20px 0; }
          .code { background: #f8f9fa; padding: 15px; border-radius: 5px; font-family: monospace; word-break: break-all; margin: 20px 0; }
          .copy-btn { background: #3498db; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; margin: 10px 0; }
          .copy-btn:hover { background: #2980b9; }
          .instructions { background: #fff3cd; padding: 15px; border-radius: 5px; margin: 20px 0; }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>✅ OAuth Authorization Successful!</h1>
          <div class="success">
            <strong>Success!</strong> Your Adobe Sign application has been authorized.
          </div>
          
          <div class="instructions">
            <h3>Next Steps:</h3>
            <ol>
              <li>Copy the authorization code below</li>
              <li>Go back to your mobile app</li>
              <li>Paste the code in the "Authorization Code" field</li>
              <li>Click "Submit" to complete the token exchange</li>
            </ol>
          </div>
          
          <h3>Authorization Code:</h3>
          <div class="code" id="authCode">${code}</div>
          
          <button class="copy-btn" onclick="copyToClipboard()">📋 Copy Code</button>
          
          <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 14px;">
            <p><strong>State:</strong> ${state || 'None'}</p>
            <p><strong>Timestamp:</strong> ${new Date().toLocaleString()}</p>
          </div>
        </div>
        
        <script>
          function copyToClipboard() {
            const codeElement = document.getElementById('authCode');
            const text = codeElement.textContent;
            
            if (navigator.clipboard) {
              navigator.clipboard.writeText(text).then(() => {
                alert('Authorization code copied to clipboard!');
              }).catch(err => {
                console.error('Failed to copy: ', err);
                fallbackCopyTextToClipboard(text);
              });
            } else {
              fallbackCopyTextToClipboard(text);
            }
          }
          
          function fallbackCopyTextToClipboard(text) {
            const textArea = document.createElement("textarea");
            textArea.value = text;
            textArea.style.position = "fixed";
            textArea.style.left = "-999999px";
            textArea.style.top = "-999999px";
            document.body.appendChild(textArea);
            textArea.focus();
            textArea.select();
            
            try {
              document.execCommand('copy');
              alert('Authorization code copied to clipboard!');
            } catch (err) {
              console.error('Fallback: Oops, unable to copy', err);
              alert('Please manually copy the authorization code');
            }
            
            document.body.removeChild(textArea);
          }
        </script>
      </body>
      </html>
    `);
  }
}
