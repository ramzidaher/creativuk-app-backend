import { 
  Controller, 
  Get, 
  Post, 
  Put, 
  Delete, 
  Body, 
  Param, 
  UseGuards, 
  Request,
  HttpCode,
  HttpStatus,
  UseInterceptors,
  UploadedFiles,
  BadRequestException,
  Query,
  Res,
  Logger
} from '@nestjs/common';
import { FilesInterceptor, AnyFilesInterceptor } from '@nestjs/platform-express';
import { diskStorage } from 'multer';
import { extname } from 'path';
import { Response as ExpressResponse } from 'express';
import * as fs from 'fs';
import { SurveyService } from './survey.service';
import { 
  CompleteSurveyDto, 
  UpdateSurveyDto, 
  SurveyResponseDto 
} from './dto/survey.dto';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { RolesGuard } from '../auth/roles.guard';
import { Roles } from '../auth/roles.decorator';
import { UserRole } from '../auth/dto/auth.dto';

@Controller('surveys')
@UseGuards(JwtAuthGuard)
export class SurveyController {
  private readonly logger = new Logger(SurveyController.name);

  constructor(private readonly surveyService: SurveyService) {
    this.logger.log('🚀 SurveyController initialized');
  }

  /**
   * Helper method to truncate base64 data for logging
   */
  private truncateBase64ForLogging(data: any): any {
    if (typeof data === 'string' && data.length > 100 && data.match(/^[A-Za-z0-9+/=]+$/)) {
      // Likely base64 data, truncate it
      return `${data.substring(0, 50)}... (truncated, ${data.length} chars)`;
    }
    
    if (Array.isArray(data)) {
      return data.map(item => this.truncateBase64ForLogging(item));
    }
    
    if (data && typeof data === 'object') {
      const truncated: any = {};
      for (const [key, value] of Object.entries(data)) {
        if (key === 'base64' || key === 'base64Data') {
          truncated[key] = typeof value === 'string' && value.length > 100 
            ? `${value.substring(0, 50)}... (truncated, ${value.length} chars)`
            : value;
        } else {
          truncated[key] = this.truncateBase64ForLogging(value);
        }
      }
      return truncated;
    }
    
    return data;
  }

  @Post(':ghlOpportunityId')
  async createSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<SurveyResponseDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    // Use the user's internal ID if ghlUserId is not set
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    // console.log(`Survey Controller - JWT Token Data:`, {
    //   sub: req.user.sub,
    //   ghlUserId: req.user.ghlUserId,
    //   role: req.user.role,
    //   extractedGhlUserId: ghlUserId,
    //   ghlOpportunityId
    // });
    
    return this.surveyService.createSurvey(ghlUserId, ghlOpportunityId);
  }

  @Get(':ghlOpportunityId')
  async getSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<SurveyResponseDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return this.surveyService.getSurvey(ghlUserId, ghlOpportunityId);
  }

  @Get()
  async getUserSurveys(@Request() req): Promise<SurveyResponseDto[]> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return this.surveyService.getUserSurveys(ghlUserId);
  }

  @Put(':ghlOpportunityId')
  async updateSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() updateDto: UpdateSurveyDto
  ): Promise<SurveyResponseDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    // console.log(`Survey Controller - Update Survey JWT Token Data:`, {
    //   sub: req.user.sub,
    //   ghlUserId: req.user.ghlUserId,
    //   role: req.user.role,
    //   extractedGhlUserId: ghlUserId,
    //   ghlOpportunityId,
    //   updateDto: Object.keys(updateDto)
    // });
    
    return this.surveyService.updateSurvey(ghlUserId, ghlOpportunityId, updateDto);
  }

  @Post(':ghlOpportunityId/submit')
  @HttpCode(HttpStatus.OK)
  @UseInterceptors(AnyFilesInterceptor({
    storage: diskStorage({
      destination: './uploads/surveys',
      filename: (req, file, callback) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        const ext = extname(file.originalname || 'image.jpg');
        const filename = `${file.fieldname}-${uniqueSuffix}${ext}`;
        callback(null, filename);
      },
    }),
    fileFilter: (req, file, callback) => {
      // Skip file filter for non-file fields (like surveyData)
      if (!file.mimetype) {
        callback(null, true);
        return;
      }
      if (file.mimetype.match(/\/(jpg|jpeg|png|gif|pdf)$/)) {
        callback(null, true);
      } else {
        callback(new Error('Only image and PDF files are allowed!'), false);
      }
    },
    limits: {
      fileSize: 10 * 1024 * 1024, // 10MB per file
    },
    preservePath: false,
  }))
  async submitSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() body: any,
    @UploadedFiles() files: Express.Multer.File[]
  ): Promise<SurveyResponseDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    // console.log(`Submitting survey for opportunity ${ghlOpportunityId}`);
    // console.log(`Received ${files?.length || 0} files`);
    // console.log(`Body keys:`, body ? Object.keys(body) : 'Body is null/undefined');
    // console.log(`Content-Type:`, req.headers['content-type']);
    // console.log(`User from request:`, req.user);
    // console.log(`GHL User ID:`, ghlUserId);
    
    // Check if this is a FormData request (from web frontend) or JSON request (from mobile)
    const isFormData = req.headers['content-type']?.includes('multipart/form-data');
    
    if (isFormData) {
      // Handle FormData request (from web frontend)
      // console.log(`Processing FormData request`);
      
      // Parse survey data from form body
      let surveyData: any = {};
      if (body.surveyData) {
        try {
          surveyData = typeof body.surveyData === 'string' 
            ? JSON.parse(body.surveyData) 
            : body.surveyData;
          // console.log(`Parsed survey data structure:`, Object.keys(this.truncateBase64ForLogging(surveyData)));
        } catch (error) {
          console.error('Failed to parse survey data:', error);
          throw new BadRequestException('Invalid survey data format');
        }
      }
      
      // Submit survey with files (FormData approach)
      return this.surveyService.submitSurveyWithFiles(ghlUserId, ghlOpportunityId, surveyData, files);
    } else {
      // Handle JSON request (from mobile app)
      // console.log(`Processing JSON request`);
      
      // Check if body exists
      if (!body) {
        throw new BadRequestException('Request body is missing');
      }
      
      // Extract survey data and images from JSON payload
      const surveyData = body.surveyData || {};
      const images = body.images || {};
      
      // console.log(`Survey data structure:`, Object.keys(this.truncateBase64ForLogging(surveyData)));
      // console.log(`Images data:`, Object.keys(images).length > 0 ? Object.keys(images) : 'No images');
      
      // Submit survey with images (JSON approach)
      return this.surveyService.submitSurveyWithImages(ghlUserId, ghlOpportunityId, surveyData, images);
    }
  }


  @Get(':ghlOpportunityId/images')
  async getSurveyImages(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<any[]> {
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return await this.surveyService.getSurveyImages(ghlUserId, ghlOpportunityId);
  }

  @Get(':ghlOpportunityId/images/:fieldName')
  async getSurveyImagesByField(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Param('fieldName') fieldName: string
  ): Promise<any[]> {
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return await this.surveyService.getSurveyImagesByField(ghlUserId, ghlOpportunityId, fieldName);
  }

  @Get(':ghlOpportunityId/images/:fieldName/:imageId/base64')
  async getSurveyImageAsBase64(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Param('imageId') imageId: string
  ): Promise<{ base64: string | null }> {
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    const base64 = await this.surveyService.getSurveyImageAsBase64(ghlUserId, ghlOpportunityId, imageId);
    
    return { base64 };
  }

  @Post(':ghlOpportunityId/approve')
  @HttpCode(HttpStatus.OK)
  async approveSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<SurveyResponseDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return this.surveyService.approveSurvey(ghlUserId, ghlOpportunityId);
  }

  @Post(':ghlOpportunityId/reject')
  @HttpCode(HttpStatus.OK)
  async rejectSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() body: { rejectionReason: string }
  ): Promise<SurveyResponseDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return this.surveyService.rejectSurvey(ghlUserId, ghlOpportunityId, body.rejectionReason);
  }

  @Delete(':ghlOpportunityId')
  async deleteSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<void> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return this.surveyService.deleteSurvey(ghlUserId, ghlOpportunityId);
  }

  @Post(':ghlOpportunityId/reset')
  @HttpCode(HttpStatus.OK)
  async resetSurvey(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<SurveyResponseDto> {
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    return this.surveyService.resetSurvey(ghlUserId, ghlOpportunityId);
  }

  @Post(':ghlOpportunityId/send-email')
  @HttpCode(HttpStatus.OK)
  @Roles(UserRole.ADMIN)
  @UseGuards(RolesGuard)
  async sendSurveyEmail(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() body: { recipientEmail?: string }
  ): Promise<{ success: boolean; message: string }> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    // Default to paldevtechnologies@gmail.com if no recipient email provided
    const recipientEmail = body.recipientEmail || 'paldevtechnologies@gmail.com';
    
    try {
      const success = await this.surveyService.sendSurveyEmail(
        ghlUserId, 
        ghlOpportunityId, 
        recipientEmail
      );
      
      if (success) {
        return { 
          success: true, 
          message: 'Survey email sent successfully' 
        };
      } else {
        return { 
          success: false, 
          message: 'Failed to send survey email' 
        };
      }
    } catch (error) {
      return { 
        success: false, 
        message: `Error sending email: ${error.message}` 
      };
    }
  }

  @Get(':ghlOpportunityId/report')
  async getSurveyReport(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Query('format') format: 'html' | 'pdf' = 'html'
  ) {
    try {
      const userId = req.user.sub;
      const ghlUserId = req.user.ghlUserId || req.user.sub;
      
      console.log(`📊 Getting survey report for opportunity ${ghlOpportunityId}, format: ${format}`);
      
      // Check if survey exists and user has access
      const survey = await this.surveyService.getSurvey(ghlUserId, ghlOpportunityId);
      if (!survey) {
        return {
          success: false,
          message: 'Survey not found or access denied'
        };
      }

      // Get report path
      const reportPath = await this.surveyService.getSurveyReportPath(ghlOpportunityId, format);
      
      if (!reportPath) {
        return {
          success: false,
          message: 'Survey report not found. Please submit the survey first.'
        };
      }

      return {
        success: true,
        data: {
          reportPath,
          format,
          opportunityId: ghlOpportunityId
        }
      };
    } catch (error) {
      console.error('Error getting survey report:', error);
      return {
        success: false,
        message: `Error getting survey report: ${error.message}`
      };
    }
  }

  @Get(':ghlOpportunityId/report/view')
  async viewSurveyReport(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Query('format') format: 'html' | 'pdf' = 'html',
    @Res() res: ExpressResponse
  ) {
    try {
      const userId = req.user.sub;
      const ghlUserId = req.user.ghlUserId || req.user.sub;
      
      console.log(`📊 Viewing survey report for opportunity ${ghlOpportunityId}, format: ${format}`);
      
      // Check if survey exists and user has access
      const survey = await this.surveyService.getSurvey(ghlUserId, ghlOpportunityId);
      if (!survey) {
        return res.status(404).json({
          success: false,
          message: 'Survey not found or access denied'
        });
      }

      // Get report path
      const reportPath = await this.surveyService.getSurveyReportPath(ghlOpportunityId, format);
      
      if (!reportPath || !fs.existsSync(reportPath)) {
        return res.status(404).json({
          success: false,
          message: 'Survey report not found. Please submit the survey first.'
        });
      }

      // Set appropriate content type
      if (format === 'html') {
        res.setHeader('Content-Type', 'text/html');
      } else if (format === 'pdf') {
        res.setHeader('Content-Type', 'application/pdf');
      }

      // Stream the file
      const fileStream = fs.createReadStream(reportPath);
      fileStream.pipe(res);
      
    } catch (error) {
      console.error('Error viewing survey report:', error);
      return res.status(500).json({
        success: false,
        message: `Error viewing survey report: ${error.message}`
      });
    }
  }

  @Put(':ghlOpportunityId/save-page')
  async saveSurveyPage(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() body: { pageData: any; images?: any }
  ): Promise<SurveyResponseDto> {
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    this.logger.log(`🔧 [CONTROLLER] PUT /surveys/${ghlOpportunityId}/save-page called`);
    this.logger.log(`🔧 [CONTROLLER] User: ${req.user.sub}, ghlUserId: ${ghlUserId}`);
    this.logger.log(`🔧 [CONTROLLER] Request body keys: ${Object.keys(body || {}).join(', ')}`);
    this.logger.log(`🔧 [CONTROLLER] Page data keys: ${Object.keys(body.pageData || {}).join(', ')}`);
    this.logger.log(`🔧 [CONTROLLER] Images provided: ${body.images ? 'Yes' : 'No'}`);
    
    try {
      const result = await this.surveyService.saveSurveyPage(ghlUserId, ghlOpportunityId, body.pageData, body.images);
      this.logger.log(`✅ [CONTROLLER] Successfully saved survey page for opportunity ${ghlOpportunityId}`);
      return result;
    } catch (error) {
      this.logger.error(`❌ [CONTROLLER] Error saving survey page:`, error);
      throw error;
    }
  }

  @Post(':ghlOpportunityId/upload-images')
  async uploadImagesAndGetUrls(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() body: { fieldName: string; images: any[] }
  ): Promise<{ success: boolean; data?: { urls: string[] }; error?: string }> {
    const ghlUserId = req.user.ghlUserId || req.user.sub;
    
    this.logger.log(`📷 [CONTROLLER] POST /surveys/${ghlOpportunityId}/upload-images called`);
    this.logger.log(`📷 [CONTROLLER] User: ${req.user.sub}, ghlUserId: ${ghlUserId}`);
    this.logger.log(`📷 [CONTROLLER] Field name: ${body.fieldName}`);
    this.logger.log(`📷 [CONTROLLER] Number of images: ${body.images?.length || 0}`);
    this.logger.log(`📷 [CONTROLLER] Request body keys: ${Object.keys(body || {}).join(', ')}`);
    
    if (body.images && body.images.length > 0) {
      this.logger.log(`📷 [CONTROLLER] Image details:`, body.images.map((img, index) => ({
        index,
        name: img.name,
        mimeType: img.mimeType,
        size: img.size,
        hasBase64Data: !!img.base64Data,
        base64Length: img.base64Data?.length || 0
      })));
    }
    
    try {
      const result = await this.surveyService.uploadImagesAndGetUrls(
        ghlUserId, 
        ghlOpportunityId, 
        body.fieldName, 
        body.images
      );
      
      this.logger.log(`✅ [CONTROLLER] Successfully uploaded images for field ${body.fieldName}, got ${result.urls.length} URLs`);
      
      return {
        success: true,
        data: result
      };
    } catch (error) {
      this.logger.error(`❌ [CONTROLLER] Error uploading images:`, error);
      this.logger.error(`❌ [CONTROLLER] Error details:`, {
        fieldName: body.fieldName,
        imageCount: body.images?.length || 0,
        error: error.message,
        stack: error.stack
      });
      
      return {
        success: false,
        error: error.message
      };
    }
  }
} 