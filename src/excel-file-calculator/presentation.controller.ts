import { Controller, Post, Body, Get, Param, Res, UseGuards, Query } from '@nestjs/common';
import { Response } from 'express';
import { PresentationService } from './presentation.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';

@Controller('presentation')
@UseGuards(JwtAuthGuard)
export class PresentationController {
  constructor(private readonly presentationService: PresentationService) {}

  /**
   * Generate presentation with session isolation
   */
  @Post('session/generate')
  async generatePresentationWithSession(@Body() data: {
    userId: string;
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      const result = await this.presentationService.generatePresentationWithSession(
        data.userId,
        data
      );
      return { success: true, data: result };
    } catch (error) {
      console.error('Session-based presentation generation error:', error);
      return { success: false, error: error.message };
    }
  }

  @Post('generate')
  async generatePresentation(@Body() data: {
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      const result = await this.presentationService.generatePresentation(data);
      return { success: true, data: result };
    } catch (error) {
      console.error('Presentation generation error:', error);
      return { success: false, error: error.message };
    }
  }

  @Post('generate-video')
  async generateVideoPresentation(@Body() data: {
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      // Changed to generate images instead of video
      const result = await this.presentationService.generateImagePresentation(data);
      return result;
    } catch (error) {
      console.error('Image presentation generation error:', error);
      return { success: false, error: error.message };
    }
  }

  @Post('generate-images')
  async generateImagePresentation(@Body() data: {
    opportunityId: string;
    calculatorType?: 'flux' | 'off-peak' | 'epvs';
    customerName?: string;
    date?: string;
    postcode?: string;
    solarData?: any;
  }) {
    try {
      const result = await this.presentationService.generateImagePresentation(data);
      return result;
    } catch (error) {
      console.error('Image presentation generation error:', error);
      return { success: false, error: error.message };
    }
  }

  @Get('download/:filename')
  async downloadPresentation(@Param('filename') filename: string, @Res() res: Response) {
    try {
      const { filePath, mimeType, size } = await this.presentationService.downloadPresentation(filename);
      
      // Set proper headers for file download
      res.setHeader('Content-Type', mimeType);
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.setHeader('Content-Length', size.toString());
      res.setHeader('Cache-Control', 'no-cache');
      res.setHeader('Pragma', 'no-cache');
      
      // Stream the file to the response
      const fs = require('fs');
      const fileStream = fs.createReadStream(filePath);
      
      fileStream.on('error', (error) => {
        console.error('File stream error:', error);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error reading file' });
        }
      });
      
      fileStream.pipe(res);
    } catch (error) {
      console.error('Download error:', error);
      if (!res.headersSent) {
        res.status(404).json({ success: false, error: 'File not found' });
      }
    }
  }

  @Get('view/:filename')
  async viewPresentation(@Param('filename') filename: string, @Res() res: Response) {
    try {
      const { filePath, mimeType, size } = await this.presentationService.downloadPresentation(filename);
      
      res.setHeader('Content-Type', mimeType);
      res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
      res.setHeader('Content-Length', size.toString());
      
      res.sendFile(filePath);
    } catch (error) {
      console.error('View error:', error);
      res.status(404).json({ success: false, error: 'File not found' });
    }
  }

  @Get('info/:filename')
  async getPresentationInfo(@Param('filename') filename: string) {
    try {
      const info = await this.presentationService.getPresentationInfo(filename);
      return { success: true, data: info };
    } catch (error) {
      console.error('Info error:', error);
      return { success: false, error: error.message };
    }
  }

  @Get('list')
  async listPresentations() {
    try {
      const presentations = await this.presentationService.listPresentations();
      return { success: true, data: presentations };
    } catch (error) {
      console.error('List error:', error);
      return { success: false, error: error.message };
    }
  }

  @Get('variables/:opportunityId')
  async getPresentationVariables(
    @Param('opportunityId') opportunityId: string,
    @Query('calculatorType') calculatorType?: 'flux' | 'off-peak' | 'epvs'
  ) {
    try {
      console.log(`🔍 getPresentationVariables called with opportunityId: ${opportunityId}, calculatorType: ${calculatorType}`);
      
      const variables = await this.presentationService.extractVariablesFromExcel(
        opportunityId, 
        calculatorType
      );
      return { success: true, data: variables };
    } catch (error) {
      console.error('Variable extraction error:', error);
      return { success: false, error: error.message };
    }
  }

  @Get('solar-projection/:opportunityId')
  async getSolarProjection(
    @Param('opportunityId') opportunityId: string,
    @Query('calculatorType') calculatorType?: 'flux' | 'off-peak' | 'epvs',
    @Query('fileName') fileName?: string
  ) {
    try {
      console.log(`🔍 getSolarProjection called with opportunityId: ${opportunityId}, calculatorType: ${calculatorType}, fileName: ${fileName}`);
      
      const solarProjection = await this.presentationService.extractSolarProjectionData(
        opportunityId, 
        calculatorType,
        fileName
      );
      return { success: true, data: solarProjection };
    } catch (error) {
      console.error('Solar projection extraction error:', error);
      return { success: false, error: error.message };
    }
  }

  @Post('solar-projection/:opportunityId/payment-method')
  async updatePaymentMethod(
    @Param('opportunityId') opportunityId: string,
    @Body() data: { paymentMethod: string; calculatorType?: 'flux' | 'off-peak' | 'epvs'; fileName?: string }
  ) {
    try {
      console.log(`🔧 updatePaymentMethod called with opportunityId: ${opportunityId}, paymentMethod: ${data.paymentMethod}, fileName: ${data.fileName}`);
      
      const result = await this.presentationService.updatePaymentMethodAndExtractData(
        opportunityId,
        data.paymentMethod,
        data.calculatorType,
        data.fileName
      );
      return { success: true, data: result };
    } catch (error) {
      console.error('Payment method update error:', error);
      return { success: false, error: error.message };
    }
  }

  @Post('solar-projection/:opportunityId/terms')
  async updateTerms(
    @Param('opportunityId') opportunityId: string,
    @Body() data: { terms: number; calculatorType?: 'flux' | 'off-peak' | 'epvs'; fileName?: string }
  ) {
    try {
      console.log(`🔧 updateTerms called with opportunityId: ${opportunityId}, terms: ${data.terms}, fileName: ${data.fileName}`);
      
      const result = await this.presentationService.updateTermsAndExtractData(
        opportunityId,
        data.terms,
        data.calculatorType,
        data.fileName
      );
      return { success: true, data: result };
    } catch (error) {
      console.error('Terms update error:', error);
      return { success: false, error: error.message };
    }
  }
}

// Public controller for downloads without authentication
@Controller('public/presentation')
export class PublicPresentationController {
  constructor(private readonly presentationService: PresentationService) {}

  @Get('download/:filename')
  async downloadPresentation(@Param('filename') filename: string, @Res() res: Response) {
    try {
      console.log(`📥 Public download requested for: ${filename}`);
      const { filePath, mimeType, size } = await this.presentationService.downloadPresentation(filename);
      
      // Set proper headers for file download
      res.setHeader('Content-Type', mimeType);
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.setHeader('Content-Length', size.toString());
      res.setHeader('Cache-Control', 'no-cache');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.setHeader('Access-Control-Allow-Methods', 'GET');
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
      
      // Stream the file to the response
      const fs = require('fs');
      const fileStream = fs.createReadStream(filePath);
      
      fileStream.on('error', (error) => {
        console.error('Public file stream error:', error);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error reading file' });
        }
      });
      
      fileStream.pipe(res);
    } catch (error) {
      console.error('Public download error:', error);
      if (!res.headersSent) {
        res.status(404).json({ success: false, error: 'File not found' });
      }
    }
  }

  @Get('view/:filename')
  async viewPresentation(@Param('filename') filename: string, @Res() res: Response) {
    try {
      console.log(`👁️ Public view requested for: ${filename}`);
      const { filePath, mimeType, size } = await this.presentationService.downloadPresentation(filename);
      
      res.setHeader('Content-Type', mimeType);
      res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
      res.setHeader('Content-Length', size.toString());
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.setHeader('Access-Control-Allow-Methods', 'GET');
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
      
      res.sendFile(filePath);
    } catch (error) {
      console.error('Public view error:', error);
      res.status(404).json({ success: false, error: 'File not found' });
    }
  }
}




