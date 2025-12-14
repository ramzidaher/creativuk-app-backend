import { 
  Controller, 
  Post, 
  Get, 
  Delete,
  Body, 
  Param, 
  UseGuards, 
  Request,
  HttpCode,
  HttpStatus,
  Query
} from '@nestjs/common';
import { AutoSaveService } from './auto-save.service';
import { AutoSaveDto, AutoSaveImageDto, GetAutoSaveDataDto } from './dto/auto-save.dto';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';

@Controller('auto-save')
@UseGuards(JwtAuthGuard)
export class AutoSaveController {
  constructor(private readonly autoSaveService: AutoSaveService) {}

  @Post('field')
  @HttpCode(HttpStatus.OK)
  async autoSaveField(
    @Request() req,
    @Body() autoSaveDto: AutoSaveDto
  ): Promise<{ success: boolean; message: string }> {
    // Use the internal user ID (req.user.sub) for auto-save operations
    // The auto-save service will handle the ghlUserId lookup internally
    const userId = req.user.sub;
    
    console.log(`Auto-save field request:`, {
      userId,
      ghlUserId: req.user.ghlUserId,
      opportunityId: autoSaveDto.opportunityId,
      fieldName: autoSaveDto.fieldName,
      hasFieldValue: autoSaveDto.fieldValue !== undefined,
      hasPageData: !!autoSaveDto.pageData
    });

    return this.autoSaveService.autoSaveField(userId, autoSaveDto);
  }

  @Post('image')
  @HttpCode(HttpStatus.OK)
  async autoSaveImage(
    @Request() req,
    @Body() autoSaveImageDto: AutoSaveImageDto
  ): Promise<{ success: boolean; message: string; imageUrl?: string }> {
    const userId = req.user.sub;
    
    console.log(`Auto-save image request:`, {
      userId,
      ghlUserId: req.user.ghlUserId,
      opportunityId: autoSaveImageDto.opportunityId,
      fieldName: autoSaveImageDto.fieldName,
      hasBase64Data: !!autoSaveImageDto.base64Data,
      base64Length: autoSaveImageDto.base64Data?.length || 0
    });

    return this.autoSaveService.autoSaveImage(userId, autoSaveImageDto);
  }

  @Get(':opportunityId')
  async getAutoSaveData(
    @Request() req,
    @Param('opportunityId') opportunityId: string,
    @Query('pageName') pageName?: string
  ): Promise<{ success: boolean; data?: any; message?: string }> {
    const userId = req.user.sub;
    
    const getAutoSaveDataDto: GetAutoSaveDataDto = {
      opportunityId,
      pageName: pageName
    };

    console.log(`Get auto-save data request:`, {
      userId,
      ghlUserId: req.user.ghlUserId,
      opportunityId,
      pageName: pageName
    });

    return this.autoSaveService.getAutoSaveData(userId, getAutoSaveDataDto);
  }

  @Delete(':opportunityId')
  @HttpCode(HttpStatus.OK)
  async clearAutoSaveData(
    @Request() req,
    @Param('opportunityId') opportunityId: string
  ): Promise<{ success: boolean; message: string }> {
    const userId = req.user.sub;
    
    console.log(`Clear auto-save data request:`, {
      userId,
      ghlUserId: req.user.ghlUserId,
      opportunityId
    });

    return this.autoSaveService.clearAutoSaveData(userId, opportunityId);
  }

  @Post(':opportunityId/transfer-to-survey')
  @HttpCode(HttpStatus.OK)
  async transferToSurvey(
    @Request() req,
    @Param('opportunityId') opportunityId: string
  ): Promise<{ success: boolean; message: string; surveyData?: any }> {
    const userId = req.user.sub;
    
    console.log(`Transfer auto-save data to survey request:`, {
      userId,
      ghlUserId: req.user.ghlUserId,
      opportunityId
    });

    return this.autoSaveService.transferToSurvey(userId, opportunityId);
  }
}
