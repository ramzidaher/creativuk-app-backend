import { Controller, Post, Get, Body, UseGuards } from '@nestjs/common';
import { DocuSignService } from './docusign.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';

@Controller('docusign')
@UseGuards(JwtAuthGuard)
export class DocuSignController {
  constructor(private readonly docuSignService: DocuSignService) {}

  @Get('test-connection')
  async testConnection() {
    return await this.docuSignService.testConnection();
  }

  @Get('user-info')
  async getUserInfo() {
    return await this.docuSignService.getUserInfo();
  }

  @Get('account-info')
  async getAccountInfo() {
    return await this.docuSignService.getAccountInfo();
  }

  @Post('create-test-envelope')
  async createTestEnvelope(@Body() body: { recipientEmail: string; recipientName: string }) {
    return await this.docuSignService.createTestEnvelope(body.recipientEmail, body.recipientName);
  }
}
