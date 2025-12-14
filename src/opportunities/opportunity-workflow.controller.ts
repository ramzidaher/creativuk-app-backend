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
  Res
} from '@nestjs/common';
import { Response } from 'express';
import { OpportunityWorkflowService } from './opportunity-workflow.service';
import { 
  StartOpportunityDto, 
  CompleteStepDto, 
  UpdateStepDto,
  OpportunityProgressDto
} from './dto/opportunity-workflow.dto';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { RolesGuard } from '../auth/roles.guard';
import { Roles } from '../auth/roles.decorator';
import { UserRole } from '../auth/dto/auth.dto';

@Controller('opportunity-workflow')
@UseGuards(JwtAuthGuard)
export class OpportunityWorkflowController {
  constructor(private readonly workflowService: OpportunityWorkflowService) {}

  @Get('steps')
  async getWorkflowSteps(): Promise<any[]> {
    return this.workflowService.getWorkflowSteps();
  }

  @Post('start')
  async startOpportunity(
    @Request() req,
    @Body() startDto: StartOpportunityDto
  ): Promise<OpportunityProgressDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || 'default';
    
    return this.workflowService.startOpportunity(ghlUserId, startDto);
  }

  @Get('progress/:ghlOpportunityId')
  async getWorkflowProgress(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<OpportunityProgressDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || 'default';
    
    return this.workflowService.getOpportunityProgress(ghlUserId, ghlOpportunityId);
  }

  @Post('progress/:ghlOpportunityId/complete-step')
  async completeStep(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() completeStepDto: CompleteStepDto
  ): Promise<OpportunityProgressDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || 'default';
    
    console.log('🔧 Workflow completeStep API called:', {
      ghlOpportunityId,
      stepNumber: completeStepDto.stepNumber,
      ghlUserId
    });
    
    const result = await this.workflowService.completeStep(ghlUserId, ghlOpportunityId, completeStepDto);
    
    console.log('🔧 Workflow completeStep API result:', {
      success: !!result,
      currentStep: result?.currentStep,
      status: result?.status
    });
    
    return result;
  }

  @Post('get-opportunity-sheets')
  async getOpportunitySheets(
    @Request() req,
    @Body() body: { opportunityId: string }
  ): Promise<any> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || 'default';
    
    return this.workflowService.getOpportunitySheets(ghlUserId, body.opportunityId);
  }

  @Delete('sheet/:opportunityId/:fileName')
  async deleteSheet(
    @Request() req,
    @Param('opportunityId') opportunityId: string,
    @Param('fileName') fileName: string
  ): Promise<{ success: boolean; message: string; error?: string }> {
    const userId = req.user.sub;
    const ghlUserId = req.user.ghlUserId || 'default';
    
    return this.workflowService.deleteSheet(ghlUserId, opportunityId, fileName);
  }

  @Post('sheet/download')
  async downloadSheet(
    @Request() req,
    @Body() body: { opportunityId: string; fileName: string },
    @Res() res: Response
  ): Promise<void> {
    const userId = req.user.sub;
    const ghlUserId = req.user.ghlUserId || 'default';
    
    await this.workflowService.downloadSheet(ghlUserId, body.opportunityId, body.fileName, res);
  }

  @Put('progress/:ghlOpportunityId/step')
  async updateStep(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string,
    @Body() updateStepDto: UpdateStepDto
  ): Promise<OpportunityProgressDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || 'default';
    
    return this.workflowService.updateStep(ghlUserId, ghlOpportunityId, updateStepDto);
  }

  @Put('progress/:ghlOpportunityId/reset')
  @HttpCode(HttpStatus.OK)
  async resetWorkflow(
    @Request() req,
    @Param('ghlOpportunityId') ghlOpportunityId: string
  ): Promise<OpportunityProgressDto> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    const ghlUserId = req.user.ghlUserId || 'default';
    
    return this.workflowService.resetWorkflow(ghlUserId, ghlOpportunityId);
  }

  @Delete('clear-all')
  @UseGuards(RolesGuard)
  @Roles(UserRole.ADMIN)
  async clearAllWorkflows(@Request() req): Promise<{ success: boolean; message: string }> {
    const userId = req.user.sub;
    const ghlUserId = req.user.ghlUserId || 'default';
    
    await this.workflowService.clearAllWorkflows(ghlUserId);
    return { success: true, message: 'All workflows cleared successfully' };
  }

  @Get('user/progress')
  async getUserWorkflows(@Request() req): Promise<OpportunityProgressDto[]> {
    const userId = req.user.sub;
    const userRole = req.user.role;
    
    console.log('🔍 getUserWorkflows controller called with:', {
      userId,
      userRole,
      ghlUserId: req.user.ghlUserId,
      username: req.user.username,
      email: req.user.email
    });
    
    // For admins, show all workflows. For surveyors, show only their workflows
    if (userRole === 'ADMIN') {
      console.log('🔍 Admin user detected, returning all workflows');
      return this.workflowService.getAllWorkflowsForAdmin();
    } else {
      console.log('🔍 Surveyor user detected, returning user-specific workflows');
      return this.workflowService.getUserWorkflowsByUserId(userId);
    }
  }
} 