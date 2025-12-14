import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import * as docusign from 'docusign-esign';
import * as jwt from 'jsonwebtoken';
import * as fs from 'fs';
import * as path from 'path';

@Injectable()
export class DocuSignService {
  private readonly logger = new Logger(DocuSignService.name);
  private apiClient: docusign.ApiClient;
  private accountId: string;
  private integrationKey: string;
  private userId: string;
  private privateKey: string;

  constructor(private readonly configService: ConfigService) {
    console.log('=== DocuSignService constructor called ===');
    this.logger.log('DocuSignService constructor called - initializing...');
    this.initializeDocuSign();
  }

  private initializeDocuSign() {
    console.log('=== initializeDocuSign called ===');
    
    // DocuSign configuration
    this.integrationKey = this.configService.get<string>('DOCUSIGN_INTEGRATION_KEY') || '';
    this.accountId = this.configService.get<string>('DOCUSIGN_ACCOUNT_ID', '42299021');
    this.userId = this.configService.get<string>('DOCUSIGN_USER_ID') || '';
    
    console.log(`Integration Key: ${this.integrationKey ? 'SET' : 'NOT SET'}`);
    console.log(`Account ID: ${this.accountId}`);
    console.log(`User ID: ${this.userId ? 'SET' : 'NOT SET'}`);
    
    this.logger.log(`DocuSign config - Integration Key: ${this.integrationKey ? 'SET' : 'NOT SET'}`);
    this.logger.log(`DocuSign config - Account ID: ${this.accountId}`);
    this.logger.log(`DocuSign config - User ID: ${this.userId ? 'SET' : 'NOT SET'}`);
    
    // Load private key
    const privateKeyPath = this.configService.get<string>('DOCUSIGN_PRIVATE_KEY_PATH');
    console.log(`Private key path: ${privateKeyPath}`);
    console.log(`Current working directory: ${process.cwd()}`);
    console.log(`Full private key path: ${path.resolve(privateKeyPath || '')}`);
    
    this.logger.log(`DocuSign private key path: ${privateKeyPath}`);
    this.logger.log(`Current working directory: ${process.cwd()}`);
    this.logger.log(`Full private key path: ${path.resolve(privateKeyPath || '')}`);
    
    if (privateKeyPath && fs.existsSync(privateKeyPath)) {
      let rawKey = fs.readFileSync(privateKeyPath, 'utf8');
      
      // Clean up the private key - remove any extra whitespace and ensure proper line endings
      this.privateKey = rawKey
        .replace(/\r\n/g, '\n')  // Normalize line endings
        .replace(/\r/g, '\n')    // Handle old Mac line endings
        .trim();                 // Remove leading/trailing whitespace
      
      console.log(`Private key loaded successfully. Length: ${this.privateKey.length}`);
      console.log(`Private key starts with: ${this.privateKey.substring(0, 50)}...`);
      console.log(`Private key ends with: ...${this.privateKey.substring(this.privateKey.length - 50)}`);
      this.logger.log(`DocuSign private key loaded successfully. Length: ${this.privateKey.length}`);
      this.logger.log(`Private key starts with: ${this.privateKey.substring(0, 50)}...`);
      this.logger.log(`Private key ends with: ...${this.privateKey.substring(this.privateKey.length - 50)}`);
    } else {
      console.log(`Private key not found at path: ${privateKeyPath}`);
      console.log(`File exists check: ${fs.existsSync(privateKeyPath || '')}`);
      this.logger.warn(`DocuSign private key not found at path: ${privateKeyPath}`);
      this.logger.warn(`File exists check: ${fs.existsSync(privateKeyPath || '')}`);
    }

    // Initialize API client
    this.apiClient = new docusign.ApiClient();
    // Use demo URL for testing - change to production when ready
    this.apiClient.setBasePath('https://demo.docusign.net/restapi');
  }

  /**
   * Generate consent URL for user to grant permission
   */
  getConsentUrl(): string {
    // Use localhost for testing - make sure this is registered in DocuSign admin panel
    const redirectUri = 'http://localhost:3000';
    const consentUrl = `https://account-d.docusign.com/oauth/auth?response_type=code&scope=signature%20impersonation&client_id=${this.integrationKey}&redirect_uri=${encodeURIComponent(redirectUri)}`;
    
    this.logger.log(`Generated consent URL: ${consentUrl}`);
    return consentUrl;
  }

  /**
   * Get JWT access token for DocuSign API
   */
  async getAccessToken(): Promise<string> {
    try {
      console.log('=== getAccessToken called ===');
      console.log(`Private key exists: ${!!this.privateKey}`);
      console.log(`Integration key exists: ${!!this.integrationKey}`);
      console.log(`User ID exists: ${!!this.userId}`);
      
      if (!this.privateKey || !this.integrationKey || !this.userId) {
        const missing: string[] = [];
        if (!this.integrationKey) missing.push('DOCUSIGN_INTEGRATION_KEY');
        if (!this.userId) missing.push('DOCUSIGN_USER_ID');
        if (!this.privateKey) missing.push('DOCUSIGN_PRIVATE_KEY_PATH');
        
        console.log(`Missing configuration: ${missing.join(', ')}`);
        throw new Error(`DocuSign configuration incomplete. Missing: ${missing.join(', ')}`);
      }

      // Create JWT assertion
      const now = Math.floor(Date.now() / 1000);
      const payload = {
        iss: this.integrationKey,
        sub: this.userId,
        aud: 'account-d.docusign.com', // Demo environment (without https://)
        iat: now,
        exp: now + 3600, // 1 hour
        scope: 'signature impersonation'
      };
      
      console.log('JWT Payload:', JSON.stringify(payload, null, 2));

      // Use DocuSign SDK's built-in JWT method instead of manual JWT signing
      console.log(`Using DocuSign SDK JWT method. Key length: ${this.privateKey.length}`);
      console.log(`Private key format check - starts with BEGIN: ${this.privateKey.includes('BEGIN')}`);
      
      this.logger.log(`Using DocuSign SDK JWT method. Key length: ${this.privateKey.length}`);
      this.logger.log(`Private key format check - starts with BEGIN: ${this.privateKey.includes('BEGIN')}`);
      
      // Use DocuSign SDK's requestJWTUserToken method directly with the private key
      const response = await this.apiClient.requestJWTUserToken(
        this.integrationKey,
        this.userId,
        'signature impersonation',
        this.privateKey,
        3600
      );

      this.logger.log('DocuSign access token obtained successfully');
      return response.body.access_token;
    } catch (error) {
      this.logger.error('Failed to get DocuSign access token:', error);
      
      // Log more details about the error
      if (error.response) {
        console.log('=== DocuSign API Error Response ===');
        console.log('Status:', error.response.status);
        console.log('Status Text:', error.response.statusText);
        console.log('Response Data:', JSON.stringify(error.response.data, null, 2));
        console.log('Response Headers:', error.response.headers);
      } else if (error.request) {
        console.log('=== DocuSign API Request Error ===');
        console.log('Request:', error.request);
      } else {
        console.log('=== DocuSign API Other Error ===');
        console.log('Error:', error.message);
      }
      
      throw error;
    }
  }

  /**
   * Get user information
   */
  async getUserInfo(): Promise<any> {
    try {
      const accessToken = await this.getAccessToken();
      this.apiClient.addDefaultHeader('Authorization', `Bearer ${accessToken}`);
      
      const userInfoApi = new docusign.UserInfoApi(this.apiClient);
      const userInfo = await userInfoApi.getUserInfo();
      
      this.logger.log('DocuSign user info retrieved successfully');
      return userInfo;
    } catch (error) {
      this.logger.error('Failed to get DocuSign user info:', error);
      throw error;
    }
  }

  /**
   * Get account information
   */
  async getAccountInfo(): Promise<any> {
    try {
      const accessToken = await this.getAccessToken();
      this.apiClient.addDefaultHeader('Authorization', `Bearer ${accessToken}`);
      
      const accountsApi = new docusign.AccountsApi(this.apiClient);
      const accountInfo = await accountsApi.getAccountInformation(this.accountId);
      
      this.logger.log('DocuSign account info retrieved successfully');
      return accountInfo;
    } catch (error) {
      this.logger.error('Failed to get DocuSign account info:', error);
      throw error;
    }
  }

  /**
   * Test DocuSign connection
   */
  async testConnection(): Promise<{ success: boolean; message: string; data?: any; consentUrl?: string }> {
    try {
      this.logger.log('Testing DocuSign connection...');
      
      // Test 1: Get access token
      const accessToken = await this.getAccessToken();
      if (!accessToken) {
        throw new Error('Failed to obtain access token');
      }

      // Test 2: Get user info
      const userInfo = await this.getUserInfo();
      
      // Test 3: Get account info
      const accountInfo = await this.getAccountInfo();

      return {
        success: true,
        message: 'DocuSign connection successful!',
        data: {
          hasAccessToken: !!accessToken,
          userInfo: {
            userId: userInfo.sub,
            name: userInfo.name,
            email: userInfo.email
          },
          accountInfo: {
            accountId: accountInfo.accountId,
            accountName: accountInfo.accountName,
            accountType: accountInfo.accountType
          }
        }
      };
    } catch (error) {
      this.logger.error('DocuSign connection test failed:', error);
      
      // Check if this is a consent-related error
      let message = `DocuSign connection failed: ${error.message}`;
      let consentUrl: string | undefined = undefined;
      
      // Check for various consent-related error patterns
      if (error.message.includes('invalid_request') || 
          error.message.includes('consent_required') ||
          error.message.includes('Request failed with status code 400')) {
        message = 'User consent is required. Please visit the consent URL to grant permission.';
        consentUrl = this.getConsentUrl();
      }
      
      return {
        success: false,
        message,
        data: null,
        consentUrl
      };
    }
  }

  /**
   * Create a simple test envelope (document for signing)
   */
  async createTestEnvelope(recipientEmail: string, recipientName: string): Promise<any> {
    try {
      const accessToken = await this.getAccessToken();
      this.apiClient.addDefaultHeader('Authorization', `Bearer ${accessToken}`);
      
      // Create envelope definition
      const envelopeDefinition = new docusign.EnvelopeDefinition();
      envelopeDefinition.emailSubject = 'Test Document for Signing';
      envelopeDefinition.status = 'sent';

      // Create document
      const document = new docusign.Document();
      document.documentBase64 = Buffer.from('Test document content').toString('base64');
      document.name = 'Test Document';
      document.fileExtension = 'txt';
      document.documentId = '1';

      envelopeDefinition.documents = [document];

      // Create recipient
      const signer = new docusign.Signer();
      signer.email = recipientEmail;
      signer.name = recipientName;
      signer.recipientId = '1';
      signer.routingOrder = '1';

      // Create sign here tab
      const signHere = new docusign.SignHere();
      signHere.documentId = '1';
      signHere.pageNumber = '1';
      signHere.recipientId = '1';
      signHere.tabLabel = 'SignHereTab';
      signHere.xPosition = '100';
      signHere.yPosition = '100';

      signer.tabs = new docusign.Tabs();
      signer.tabs.signHereTabs = [signHere];

      envelopeDefinition.recipients = new docusign.Recipients();
      envelopeDefinition.recipients.signers = [signer];

      // Send envelope
      const envelopesApi = new docusign.EnvelopesApi(this.apiClient);
      const result = await envelopesApi.createEnvelope(this.accountId, envelopeDefinition);

      this.logger.log(`Test envelope created successfully: ${result.envelopeId}`);
      return {
        success: true,
        envelopeId: result.envelopeId,
        status: result.status,
        message: 'Test envelope created and sent successfully!'
      };
    } catch (error) {
      this.logger.error('Failed to create test envelope:', error);
      throw error;
    }
  }
}
