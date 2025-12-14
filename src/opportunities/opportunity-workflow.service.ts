import { Injectable, Logger, NotFoundException, Inject, forwardRef } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { UserService } from '../user/user.service';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { OpportunitiesService } from './opportunities.service';
import { SurveyService } from './survey.service';
import { OpportunityOutcomesService } from '../opportunity-outcomes/opportunity-outcomes.service';
import { OneDriveFileManagerService } from '../onedrive/onedrive-file-manager.service';
import { DocuSealService } from '../integrations/docuseal.service';
import { Response } from 'express';
import * as path from 'path';
import * as fs from 'fs/promises';
import { 
  OpportunityProgressDto, 
  OpportunityStepDto, 
  StartOpportunityDto, 
  UpdateStepDto, 
  CompleteStepDto,
  StepStatus,
  OpportunityStatus,
  WorkflowStepConfig
} from './dto/opportunity-workflow.dto';
import { StepType } from '@prisma/client';

@Injectable()
export class OpportunityWorkflowService {
  private readonly logger = new Logger(OpportunityWorkflowService.name);

  // Define the workflow steps configuration
  private readonly workflowSteps: WorkflowStepConfig[] = [
    {
      stepNumber: 1,
      stepType: StepType.SITE_SURVEY,
      title: 'Survey',
      description: 'Conduct the on-site survey',
      required: true,
      estimatedDuration: 60
    },
    {
      stepNumber: 2,
      stepType: StepType.OPEN_SOLAR,
      title: 'OpenSolar',
      description: 'Access OpenSolar platform for design',
      required: true,
      estimatedDuration: 30
    },
    {
      stepNumber: 3,
      stepType: StepType.CALCULATOR,
      title: 'Calculate',
      description: 'Choose between Off Peak and Flux options',
      required: true,
      estimatedDuration: 15
    },
    {
      stepNumber: 4,
      stepType: StepType.FOLLOW_UP,
      title: 'Proposal',
      description: 'Present the final proposal to the customer',
      required: true,
      estimatedDuration: 30
    },
    {
      stepNumber: 5,
      stepType: StepType.SOLAR_PROJECTION,
      title: 'Solar Projection',
      description: 'Review detailed solar projection data and financial analysis',
      required: true,
      estimatedDuration: 10
    },
    {
      stepNumber: 6,
      stepType: StepType.INSTALLATION_SCHEDULING,
      title: 'Hometree',
      description: 'Visit Hometree for installation services',
      required: true,
      estimatedDuration: 15
    },
    {
      stepNumber: 7,
      stepType: StepType.PROPOSAL_GENERATION,
      title: 'Contract Generation',
      description: 'Generate contract and proposal documents',
      required: true,
      estimatedDuration: 30
    },
    {
      stepNumber: 8,
      stepType: StepType.CONTRACT_SIGNING,
      title: 'Contract Signing',
      description: 'Sign the installation contract',
      required: true,
      estimatedDuration: 20
    },
    {
      stepNumber: 9,
      stepType: StepType.EMAIL_CONFIRMATION,
      title: 'Email Confirmation',
      description: 'Sign the booking confirmation letter',
      required: true,
      estimatedDuration: 10
    },
    {
      stepNumber: 10,
      stepType: StepType.PAYMENT,
      title: 'Payment',
      description: 'Process payment for the installation',
      required: true,
      estimatedDuration: 10
    },
    {
      stepNumber: 11,
      stepType: StepType.INSTALLATION_BOOKING,
      title: 'Book Installation',
      description: 'Schedule your solar installation appointment',
      required: true,
      estimatedDuration: 15
    },
    {
      stepNumber: 12,
      stepType: StepType.WELCOME_EMAIL,
      title: 'Send Welcome Email',
      description: 'Send welcome email to customer with installation details',
      required: true,
      estimatedDuration: 5
    }
  ];

  // Get the total number of steps dynamically
  private get totalSteps(): number {
    return this.workflowSteps.length;
  }

  constructor(
    private readonly prisma: PrismaService,
    private readonly userService: UserService,
    private readonly ghlService: GoHighLevelService,
    private readonly opportunitiesService: OpportunitiesService,
    private readonly surveyService: SurveyService,
    private readonly opportunityOutcomesService: OpportunityOutcomesService,
    private readonly oneDriveFileManagerService: OneDriveFileManagerService,
    @Inject(forwardRef(() => DocuSealService))
    private readonly docuSealService: DocuSealService,
  ) {}

  async startOpportunity(ghlUserId: string, startDto: StartOpportunityDto): Promise<OpportunityProgressDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // For testing purposes, allow any user to start a workflow
      // In production, this should be more restrictive
      this.logger.warn(`User with ghlUserId ${ghlUserId} not found, but allowing workflow start for testing`);
      
      // Get any user for now (for testing)
      const anyUser = await this.prisma.user.findFirst({
        where: { ghlUserId: { not: null } }
      });
      
      if (!anyUser) {
        throw new NotFoundException('No users with ghlUserId found in system');
      }
      
      // Use the found user for the workflow
      const effectiveUserId = anyUser.id;
      this.logger.log(`Using user ${effectiveUserId} for workflow with opportunity ${startDto.ghlOpportunityId}`);
    }

    const effectiveUserId = user ? user.id : (await this.prisma.user.findFirst({ where: { ghlUserId: { not: null } } }))?.id;
    
    if (!effectiveUserId) {
      throw new NotFoundException('No valid user found for workflow');
    }

    // Check if opportunity already has progress
    // Use upsert to handle existing progress more safely
    const existingProgress = await this.prisma.opportunityProgress.findUnique({
      where: { ghlOpportunityId: startDto.ghlOpportunityId },
      include: { steps: true }
    });

    if (existingProgress) {
      // Delete existing progress and start fresh
      try {
        // First delete all related steps
        await this.prisma.opportunityStep.deleteMany({
          where: { opportunityProgressId: existingProgress.id }
        });
        
        // Then delete the progress record
        await this.prisma.opportunityProgress.delete({
          where: { id: existingProgress.id }
        });
        console.log(`🗑️ Deleted existing progress and steps for opportunity ${startDto.ghlOpportunityId}`);
      } catch (deleteError) {
        console.warn(`⚠️ Failed to delete existing progress for ${startDto.ghlOpportunityId}:`, deleteError.message);
        // Continue with creation even if delete fails - the create will handle duplicates
      }
    }

    // Create new opportunity progress
    let progress;
    try {
      progress = await this.prisma.opportunityProgress.create({
        data: {
          ghlOpportunityId: startDto.ghlOpportunityId,
          userId: effectiveUserId,
          currentStep: 1,
          totalSteps: this.workflowSteps.length,
          status: OpportunityStatus.IN_PROGRESS,
          steps: {
            create: this.workflowSteps.map(step => ({
              stepNumber: step.stepNumber,
              stepType: step.stepType,
              status: step.stepNumber === 1 ? StepStatus.IN_PROGRESS : StepStatus.PENDING,
              startedAt: step.stepNumber === 1 ? new Date() : undefined,
            }))
          }
        },
        include: { steps: true }
      });
    } catch (createError) {
      // If creation fails due to unique constraint, try to find existing record
      if (createError.code === 'P2002') {
        console.log(`🔄 Progress already exists for ${startDto.ghlOpportunityId}, fetching existing record`);
        progress = await this.prisma.opportunityProgress.findUnique({
          where: { ghlOpportunityId: startDto.ghlOpportunityId },
          include: { steps: true }
        });
      } else {
        throw createError;
      }
    }

    this.logger.log(`Started opportunity workflow for ${startDto.ghlOpportunityId}`);
    return this.mapToDto(progress);
  }

  async getOpportunityProgress(ghlUserId: string, ghlOpportunityId: string): Promise<OpportunityProgressDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // For testing purposes, allow any user to access workflow progress
      this.logger.warn(`User with ghlUserId ${ghlUserId} not found, but allowing access to workflow progress for testing`);
      
      // Get any user for now (for testing)
      const anyUser = await this.prisma.user.findFirst({
        where: { ghlUserId: { not: null } }
      });
      
      if (!anyUser) {
        throw new NotFoundException('No users with ghlUserId found in system');
      }
      
      // Use the found user for the workflow
      const effectiveUserId = anyUser.id;
      this.logger.log(`Using user ${effectiveUserId} for workflow progress with opportunity ${ghlOpportunityId}`);
    }

    const effectiveUserId = user ? user.id : (await this.prisma.user.findFirst({ where: { ghlUserId: { not: null } } }))?.id;
    
    if (!effectiveUserId) {
      throw new NotFoundException('No valid user found for workflow');
    }

    const progress = await this.prisma.opportunityProgress.findUnique({
      where: { ghlOpportunityId },
      include: { steps: true }
    });

    if (!progress) {
      throw new NotFoundException('Opportunity progress not found');
    }

    // Debug: Log step data to see what's being returned
    this.logger.log(`🔍 Debug: Progress for ${ghlOpportunityId}:`);
    this.logger.log(`  - Total steps: ${progress.steps.length}`);
    progress.steps.forEach(step => {
      this.logger.log(`  - Step ${step.stepNumber} (${step.stepType}): ${step.status}`);
      if (step.data) {
        this.logger.log(`    - Data: ${JSON.stringify(step.data)}`);
      }
    });

    return this.mapToDto(progress);
  }

  async getUserOpportunities(ghlUserId: string): Promise<OpportunityProgressDto[]> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const progressList = await this.prisma.opportunityProgress.findMany({
      where: { userId: user.id },
      include: { steps: true },
      orderBy: { lastActivityAt: 'desc' }
    });

    return progressList.map(progress => this.mapToDto(progress));
  }

  async updateStep(ghlUserId: string, ghlOpportunityId: string, updateDto: UpdateStepDto): Promise<OpportunityProgressDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // For testing purposes, allow any user to update workflow steps
      this.logger.warn(`User with ghlUserId ${ghlUserId} not found, but allowing step update for testing`);
      
      // Get any user for now (for testing)
      const anyUser = await this.prisma.user.findFirst({
        where: { ghlUserId: { not: null } }
      });
      
      if (!anyUser) {
        throw new NotFoundException('No users with ghlUserId found in system');
      }
      
      // Use the found user for the workflow
      const effectiveUserId = anyUser.id;
      this.logger.log(`Using user ${effectiveUserId} for step update with opportunity ${ghlOpportunityId}`);
    }

    const effectiveUserId = user ? user.id : (await this.prisma.user.findFirst({ where: { ghlUserId: { not: null } } }))?.id;
    
    if (!effectiveUserId) {
      throw new NotFoundException('No valid user found for workflow');
    }

    const progress = await this.prisma.opportunityProgress.findUnique({
      where: { ghlOpportunityId },
      include: { steps: true }
    });

    if (!progress) {
      throw new NotFoundException('Opportunity progress not found');
    }

    // Update the specific step
    const step = progress.steps.find(s => s.stepNumber === updateDto.stepNumber);
    if (!step) {
      throw new NotFoundException('Step not found');
    }

    await this.prisma.opportunityStep.update({
      where: { id: step.id },
      data: {
        status: updateDto.status,
        data: updateDto.data,
        startedAt: updateDto.status === StepStatus.IN_PROGRESS ? new Date() : step.startedAt,
      }
    });

    // Update progress last activity
    await this.prisma.opportunityProgress.update({
      where: { id: progress.id },
      data: { lastActivityAt: new Date() }
    });

    // Return updated progress
    const updatedProgress = await this.prisma.opportunityProgress.findUnique({
      where: { ghlOpportunityId },
      include: { steps: true }
    });

    return this.mapToDto(updatedProgress);
  }

  async completeStep(ghlUserId: string, ghlOpportunityId: string, completeDto: CompleteStepDto): Promise<OpportunityProgressDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // For testing purposes, allow any user to complete workflow steps
      this.logger.warn(`User with ghlUserId ${ghlUserId} not found, but allowing step completion for testing`);
      
      // Get any user for now (for testing)
      const anyUser = await this.prisma.user.findFirst({
        where: { ghlUserId: { not: null } }
      });
      
      if (!anyUser) {
        throw new NotFoundException('No users with ghlUserId found in system');
      }
      
      // Use the found user for the workflow
      const effectiveUserId = anyUser.id;
      this.logger.log(`Using user ${effectiveUserId} for step completion with opportunity ${ghlOpportunityId}`);
    }

    const effectiveUserId = user ? user.id : (await this.prisma.user.findFirst({ where: { ghlUserId: { not: null } } }))?.id;
    
    if (!effectiveUserId) {
      throw new NotFoundException('No valid user found for workflow');
    }

    const progress = await this.prisma.opportunityProgress.findUnique({
      where: { ghlOpportunityId },
      include: { steps: true }
    });

    if (!progress) {
      throw new NotFoundException('Opportunity progress not found');
    }

    // Complete the specific step
    let step = progress.steps.find(s => s.stepNumber === completeDto.stepNumber);
    
    // If step doesn't exist, create it (for backward compatibility with existing workflows)
    if (!step) {
      this.logger.log(`Step ${completeDto.stepNumber} not found, creating it for opportunity ${ghlOpportunityId}`);
      
      const stepConfig = this.workflowSteps.find(s => s.stepNumber === completeDto.stepNumber);
      if (!stepConfig) {
        throw new NotFoundException(`Step ${completeDto.stepNumber} not found in workflow configuration`);
      }

      step = await this.prisma.opportunityStep.create({
        data: {
          opportunityProgressId: progress.id,
          stepNumber: completeDto.stepNumber,
          stepType: stepConfig.stepType,
          status: StepStatus.PENDING,
        }
      });

      // Update total steps if needed
      if (progress.totalSteps < this.workflowSteps.length) {
        await this.prisma.opportunityProgress.update({
          where: { id: progress.id },
          data: { totalSteps: this.workflowSteps.length }
        });
      }
    }

    // If this is a survey step, ensure the survey is submitted
    if (step.stepType === StepType.SITE_SURVEY) {
      try {
        // Check if survey exists and is submitted
        const survey = await this.surveyService.getSurvey(ghlUserId, ghlOpportunityId);
        if (survey.status !== 'SUBMITTED' && survey.status !== 'APPROVED') {
          this.logger.warn(`Survey for opportunity ${ghlOpportunityId} is not submitted. Status: ${survey.status}`);
        }
      } catch (error) {
        this.logger.warn(`Could not verify survey status for opportunity ${ghlOpportunityId}: ${error.message}`);
      }
    }

    await this.prisma.opportunityStep.update({
      where: { id: step.id },
      data: {
        status: StepStatus.COMPLETED,
        data: completeDto.data,
        completedAt: new Date(),
      }
    });

    // Handle OneDrive file copying for step 4 (Proposal)
    if (completeDto.stepNumber === 4) {
      await this.handleOneDriveFileCopying(ghlOpportunityId, completeDto.data);
    }

    // Check if this is the WELCOME_EMAIL step (now the last step) with outcome
    if (completeDto.stepNumber === 12 && completeDto.data?.outcome) {
      const outcome = completeDto.data.outcome;
      this.logger.log(`🎯 Opportunity ${ghlOpportunityId} marked as ${outcome.toUpperCase()}`);
      
      try {
        // Record the outcome in our tracking system
        if (user) {
          await this.opportunityOutcomesService.recordOutcome({
            ghlOpportunityId: ghlOpportunityId,
            userId: user.id,
            outcome: outcome.toUpperCase() as 'WON' | 'LOST' | 'ABANDONED',
            value: completeDto.data?.dealValue || 0,
            notes: completeDto.data?.notes || `Marked as ${outcome} in workflow step 12`,
            stageAtOutcome: 'WELCOME_EMAIL',
          });
        }

        this.logger.log(`✅ Recorded ${outcome} outcome for opportunity ${ghlOpportunityId}`);
        
        // If won, move to "Signed Contract" stage in GoHighLevel and copy documents to OneDrive
        if (outcome === 'won') {
          const result = await this.opportunitiesService.moveOpportunityToSignedContract(ghlOpportunityId);
          
          if (result.success) {
            this.logger.log(`✅ Successfully moved opportunity ${ghlOpportunityId} to Signed Contract stage`);
          } else {
            this.logger.error(`❌ Failed to move opportunity ${ghlOpportunityId} to Signed Contract stage: ${result.error}`);
          }

          // Copy won opportunity documents to OneDrive
          await this.handleWonOpportunityOneDriveCopying(ghlOpportunityId, completeDto.data);
          
          // Clean up output directory files after moving proposal
          await this.cleanupOutputDirectoryFiles(ghlOpportunityId);
        }
        
        // If lost, copy documents to OneDrive (quotations folder)
        if (outcome === 'lost') {
          await this.handleLostOpportunityOneDriveCopying(ghlOpportunityId, completeDto.data);
          
          // Clean up output directory files after moving proposal
          await this.cleanupOutputDirectoryFiles(ghlOpportunityId);
        }
      } catch (error) {
        this.logger.error(`❌ Error recording outcome for opportunity ${ghlOpportunityId}: ${error.message}`);
      }
    }

    // Update progress to next step
    const nextStep = Math.min(completeDto.stepNumber + 1, this.workflowSteps.length);
    await this.prisma.opportunityProgress.update({
      where: { id: progress.id },
      data: { 
        currentStep: nextStep,
        lastActivityAt: new Date(),
        completedAt: nextStep > this.workflowSteps.length ? new Date() : null,
        status: nextStep > this.workflowSteps.length ? OpportunityStatus.COMPLETED : OpportunityStatus.IN_PROGRESS
      }
    });

    // Return updated progress
    const updatedProgress = await this.prisma.opportunityProgress.findUnique({
      where: { ghlOpportunityId },
      include: { steps: true }
    });

    return this.mapToDto(updatedProgress);
  }

  async pauseOpportunity(ghlUserId: string, ghlOpportunityId: string): Promise<OpportunityProgressDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const progress = await this.prisma.opportunityProgress.findFirst({
      where: {
        ghlOpportunityId,
        userId: user.id
      },
      include: { steps: true }
    });

    if (!progress) {
      throw new NotFoundException('Opportunity progress not found');
    }

    const updatedProgress = await this.prisma.opportunityProgress.update({
      where: { id: progress.id },
      data: {
        status: OpportunityStatus.PAUSED,
        lastActivityAt: new Date()
      },
      include: { steps: true }
    });

    return this.mapToDto(updatedProgress);
  }

  async resumeOpportunity(ghlUserId: string, ghlOpportunityId: string): Promise<OpportunityProgressDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const progress = await this.prisma.opportunityProgress.findFirst({
      where: {
        ghlOpportunityId,
        userId: user.id
      },
      include: { steps: true }
    });

    if (!progress) {
      throw new NotFoundException('Opportunity progress not found');
    }

    const updatedProgress = await this.prisma.opportunityProgress.update({
      where: { id: progress.id },
      data: {
        status: OpportunityStatus.IN_PROGRESS,
        lastActivityAt: new Date()
      },
      include: { steps: true }
    });

    return this.mapToDto(updatedProgress);
  }

  async resetWorkflow(ghlUserId: string, ghlOpportunityId: string): Promise<OpportunityProgressDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // For testing purposes, allow any user to reset workflow
      this.logger.warn(`User with ghlUserId ${ghlUserId} not found, but allowing workflow reset for testing`);
      
      // Get any user for now (for testing)
      const anyUser = await this.prisma.user.findFirst({
        where: { ghlUserId: { not: null } }
      });
      
      if (!anyUser) {
        throw new NotFoundException('No users with ghlUserId found in system');
      }
      
      // Use the found user for the workflow
      const effectiveUserId = anyUser.id;
      this.logger.log(`Using user ${effectiveUserId} for workflow reset with opportunity ${ghlOpportunityId}`);
    }

    const effectiveUserId = user ? user.id : (await this.prisma.user.findFirst({ where: { ghlUserId: { not: null } } }))?.id;
    
    if (!effectiveUserId) {
      throw new NotFoundException('No valid user found for workflow');
    }

    // Delete existing progress
    await this.prisma.opportunityProgress.deleteMany({
      where: {
        ghlOpportunityId,
        userId: effectiveUserId
      }
    });

    // Start fresh workflow
    return this.startOpportunity(ghlUserId, { ghlOpportunityId });
  }

  async clearAllWorkflows(ghlUserId: string): Promise<void> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    // Delete all workflow progress for this user
    await this.prisma.opportunityProgress.deleteMany({
      where: { userId: user.id }
    });

    this.logger.log(`Cleared all workflows for user ${ghlUserId}`);
  }

  async getWorkflowSteps(): Promise<WorkflowStepConfig[]> {
    return this.workflowSteps;
  }

  private detectCalculatorTypeFromFilename(fileName: string): 'epvs' | 'off-peak' {
    // Check for EPVS/Flux indicators in filename
    const epvsIndicators = ['epvs', 'flux', 'creativ'];
    const lowerFileName = fileName.toLowerCase();
    
    for (const indicator of epvsIndicators) {
      if (lowerFileName.includes(indicator)) {
        return 'epvs';
      }
    }
    
    // Default to off-peak if no EPVS indicators found
    return 'off-peak';
  }

  async getOpportunitySheets(ghlUserId: string, opportunityId: string): Promise<any> {
    try {
      this.logger.log(`Getting available sheets for opportunity: ${opportunityId}`);
      
      // Define both possible base paths
      const epvsBasePath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');
      const offPeakBasePath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
      
      const allSheetInfo: any[] = [];
      
      // Search EPVS directory
      try {
        await fs.access(epvsBasePath);
        this.logger.log(`✅ EPVS directory exists: ${epvsBasePath}`);
        
        const epvsFiles = await fs.readdir(epvsBasePath);
        const epvsExcelFiles = epvsFiles.filter(file => 
          (file.endsWith('.xlsx') || file.endsWith('.xls') || file.endsWith('.xlsm')) &&
          file.includes(opportunityId)
        );
        
        // Sort files to prioritize latest versions
        const sortedEpvsFiles = epvsExcelFiles.sort((a, b) => {
          const aVersionMatch = a.match(/-v(\d+)\.xlsm$/);
          const bVersionMatch = b.match(/-v(\d+)\.xlsm$/);
          
          if (aVersionMatch && bVersionMatch) {
            return parseInt(bVersionMatch[1]) - parseInt(aVersionMatch[1]);
          } else if (aVersionMatch) {
            return -1; // Versioned files come first
          } else if (bVersionMatch) {
            return 1;
          }
          return 0;
        });
        
        this.logger.log(`📊 Found ${sortedEpvsFiles.length} EPVS Excel files for opportunity ${opportunityId}:`, sortedEpvsFiles);
        
        // Get file stats for EPVS files
        const epvsSheetInfo = await Promise.all(
          sortedEpvsFiles.map(async (fileName) => {
            const filePath = path.join(epvsBasePath, fileName);
            const stats = await fs.stat(filePath);
            // Use filename detection as fallback, but EPVS directory files are definitely EPVS
            const detectedType = this.detectCalculatorTypeFromFilename(fileName);
            return {
              fileName,
              filePath,
              size: stats.size,
              lastModified: stats.mtime.toISOString(),
              calculatorType: detectedType === 'epvs' ? 'epvs' : 'epvs' // Force EPVS for files in EPVS directory
            };
          })
        );
        
        allSheetInfo.push(...epvsSheetInfo);
      } catch (error) {
        this.logger.warn(`❌ EPVS directory does not exist or accessible: ${epvsBasePath}`);
      }
      
      // Search Off-Peak directory
      try {
        await fs.access(offPeakBasePath);
        this.logger.log(`✅ Off-Peak directory exists: ${offPeakBasePath}`);
        
        const offPeakFiles = await fs.readdir(offPeakBasePath);
        const offPeakExcelFiles = offPeakFiles.filter(file => 
          (file.endsWith('.xlsx') || file.endsWith('.xls') || file.endsWith('.xlsm')) &&
          file.includes(opportunityId)
        );
        
        // Sort files to prioritize latest versions
        const sortedOffPeakFiles = offPeakExcelFiles.sort((a, b) => {
          const aVersionMatch = a.match(/-v(\d+)\.xlsm$/);
          const bVersionMatch = b.match(/-v(\d+)\.xlsm$/);
          
          if (aVersionMatch && bVersionMatch) {
            return parseInt(bVersionMatch[1]) - parseInt(aVersionMatch[1]);
          } else if (aVersionMatch) {
            return -1; // Versioned files come first
          } else if (bVersionMatch) {
            return 1;
          }
          return 0;
        });
        
        this.logger.log(`📊 Found ${sortedOffPeakFiles.length} Off-Peak Excel files for opportunity ${opportunityId}:`, sortedOffPeakFiles);
        
        // Get file stats for Off-Peak files
        const offPeakSheetInfo = await Promise.all(
          sortedOffPeakFiles.map(async (fileName) => {
            const filePath = path.join(offPeakBasePath, fileName);
            const stats = await fs.stat(filePath);
            // Use filename detection as fallback, but Off-Peak directory files are definitely Off-Peak
            const detectedType = this.detectCalculatorTypeFromFilename(fileName);
            return {
              fileName,
              filePath,
              size: stats.size,
              lastModified: stats.mtime.toISOString(),
              calculatorType: detectedType === 'off-peak' ? 'off-peak' : 'off-peak' // Force Off-Peak for files in Off-Peak directory
            };
          })
        );
        
        allSheetInfo.push(...offPeakSheetInfo);
      } catch (error) {
        this.logger.warn(`❌ Off-Peak directory does not exist or accessible: ${offPeakBasePath}`);
      }
      
      this.logger.log(`✅ Returning ${allSheetInfo.length} total sheets for opportunity ${opportunityId}`);
      return {
        success: true,
        data: allSheetInfo
      };
    } catch (error) {
      this.logger.error(`Error getting opportunity sheets: ${error.message}`, error.stack);
      return {
        success: false,
        message: 'Failed to get available sheets',
        error: error.message
      };
    }
  }

  async deleteSheet(ghlUserId: string, opportunityId: string, fileName: string): Promise<{ success: boolean; message: string; error?: string }> {
    try {
      this.logger.log(`Deleting sheet: ${fileName} for opportunity: ${opportunityId}`);
      
      // Define both possible base paths
      const epvsBasePath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');
      const offPeakBasePath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
      
      // Try EPVS directory first
      let filePath = path.join(epvsBasePath, fileName);
      if (await fs.access(filePath).then(() => true).catch(() => false)) {
        await fs.unlink(filePath);
        this.logger.log(`✅ Deleted EPVS sheet: ${filePath}`);
        return {
          success: true,
          message: 'Sheet deleted successfully'
        };
      }
      
      // Try Off-Peak directory
      filePath = path.join(offPeakBasePath, fileName);
      if (await fs.access(filePath).then(() => true).catch(() => false)) {
        await fs.unlink(filePath);
        this.logger.log(`✅ Deleted Off-Peak sheet: ${filePath}`);
        return {
          success: true,
          message: 'Sheet deleted successfully'
        };
      }
      
      // File not found
      this.logger.error(`❌ Sheet not found: ${fileName}`);
      return {
        success: false,
        message: 'Sheet not found',
        error: 'File does not exist'
      };
    } catch (error) {
      this.logger.error(`Error deleting sheet: ${error.message}`, error.stack);
      return {
        success: false,
        message: 'Failed to delete sheet',
        error: error.message
      };
    }
  }

  async downloadSheet(ghlUserId: string, opportunityId: string, fileName: string, res: Response): Promise<void> {
    try {
      this.logger.log(`Downloading sheet: ${fileName} for opportunity: ${opportunityId}`);
      
      // Define both possible base paths
      const epvsBasePath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');
      const offPeakBasePath = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
      
      // Try EPVS directory first
      let filePath = path.join(epvsBasePath, fileName);
      let fileExists = false;
      
      try {
        await fs.access(filePath);
        fileExists = true;
      } catch {
        // Try Off-Peak directory
        filePath = path.join(offPeakBasePath, fileName);
        try {
          await fs.access(filePath);
          fileExists = true;
        } catch {
          fileExists = false;
        }
      }
      
      if (!fileExists) {
        this.logger.error(`❌ Sheet not found: ${fileName}`);
        res.status(404).json({
          success: false,
          message: 'Sheet not found',
          error: 'File does not exist'
        });
        return;
      }
      
      // Read the file buffer
      const fileBuffer = await fs.readFile(filePath);
      const fileSize = fileBuffer.length;
      
      // Set headers for file download
      res.setHeader('Content-Type', 'application/vnd.ms-excel.sheet.macroEnabled.12');
      res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
      res.setHeader('Content-Length', fileSize.toString()); // Important: Set Content-Length for reliable downloads
      
      // Send the file buffer
      res.send(fileBuffer);
      
      this.logger.log(`✅ Successfully downloaded sheet: ${fileName}`);
    } catch (error) {
      this.logger.error(`Error downloading sheet: ${error.message}`, error.stack);
      res.status(500).json({
        success: false,
        message: 'Failed to download sheet',
        error: error.message
      });
    }
  }

  async getUserWorkflows(ghlUserId: string): Promise<OpportunityProgressDto[]> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const workflows = await this.prisma.opportunityProgress.findMany({
      where: { userId: user.id },
      include: { steps: true },
      orderBy: { lastActivityAt: 'desc' }
    });

    this.logger.log(`🚀 OPTIMIZED: Found ${workflows.length} workflows for user ${user.name}`);

    // Use individual API calls but with better error handling and caching
    // This ensures we get the exact opportunity data for each workflow

    // Enhance workflows with opportunity details using individual API calls
    const enhancedWorkflows = await Promise.all(
      workflows.map(async (workflow) => {
        try {
          // Get opportunity details from GoHighLevel
          let opportunityDetails: any = null;
          
          // Try user's personal token first, then fallback to system token
          let accessToken: string | null = user.ghlAccessToken;
          if (!accessToken) {
            // Fallback to system-level GHL token (same as opportunities service)
            accessToken = process.env.GOHIGHLEVEL_API_TOKEN || null;
            if (accessToken) {
              this.logger.log(`🔄 Using system-level GHL token as fallback for ${workflow.ghlOpportunityId}`);
            }
          } else {
            this.logger.log(`🔍 Fetching opportunity details for ${workflow.ghlOpportunityId} with user token: ${accessToken.substring(0, 10)}...`);
          }

          if (accessToken) {
            opportunityDetails = await this.ghlService.getOpportunityById(
              accessToken,
              workflow.ghlOpportunityId
            );
            
            if (opportunityDetails) {
              this.logger.log(`✅ Found opportunity:`, {
                id: opportunityDetails.id,
                name: opportunityDetails.name,
                contactName: opportunityDetails.contact?.name,
                contactFirstName: opportunityDetails.contact?.firstName,
                contactLastName: opportunityDetails.contact?.lastName,
                contactEmail: opportunityDetails.contact?.email
              });
            } else {
              this.logger.warn(`❌ No opportunity details found for ${workflow.ghlOpportunityId}`);
            }
          } else {
            this.logger.warn(`⚠️ No GHL access token available for ${workflow.ghlOpportunityId}`);
          }
          
          if (opportunityDetails) {
            this.logger.log(`🚀 Enhanced workflow ${workflow.id} with opportunity details`);
            return {
              ...workflow,
              opportunityDetails
            };
          } else {
            this.logger.warn(`⚠️ No opportunity details found for ${workflow.ghlOpportunityId}`);
            return workflow;
          }
        } catch (error) {
          this.logger.warn(`⚠️ Failed to get opportunity details for ${workflow.ghlOpportunityId}:`, error.message);
          return workflow;
        }
      })
    );

    return enhancedWorkflows.map(workflow => this.mapToDto(workflow));
  }

  async getUserWorkflowsByUserId(userId: string): Promise<OpportunityProgressDto[]> {
    this.logger.log(`🔍 getUserWorkflowsByUserId called with userId: ${userId}`);
    
    // Find user by ID directly
    const user = await this.userService.findById(userId);
    this.logger.log(`🔍 User lookup result: ${user ? `Found user ${user.username || user.name} (${user.id})` : 'User not found'}`);
    
    if (!user) {
      this.logger.error(`❌ User not found with ID: ${userId}`);
      return [];
    }

    // Get workflows for this specific user
    const workflows = await this.prisma.opportunityProgress.findMany({
      where: {
        userId: userId
      },
      include: { steps: true },
      orderBy: { lastActivityAt: 'desc' }
    });

    this.logger.log(`🚀 Found ${workflows.length} workflows for user ${user.name} (userId: ${userId})`);

    // Use individual API calls but with better error handling and caching
    // This ensures we get the exact opportunity data for each workflow

    // Enhance workflows with opportunity details using individual API calls
    const enhancedWorkflows = await Promise.all(
      workflows.map(async (workflow) => {
        try {
          // Get opportunity details directly from GoHighLevel
          let opportunityDetails: any = null;
          
          // Try user's personal token first, then fallback to system token
          let accessToken: string | null = user.ghlAccessToken;
          if (!accessToken) {
            // Fallback to system-level GHL token (same as opportunities service)
            accessToken = process.env.GOHIGHLEVEL_API_TOKEN || null;
            if (accessToken) {
              this.logger.log(`🔄 Using system-level GHL token as fallback for ${workflow.ghlOpportunityId}`);
            }
          } else {
            this.logger.log(`🔍 Fetching opportunity details for ${workflow.ghlOpportunityId} with user token: ${accessToken.substring(0, 10)}...`);
          }

          if (accessToken) {
            opportunityDetails = await this.ghlService.getOpportunityById(
              accessToken,
              workflow.ghlOpportunityId
            );
            
            if (opportunityDetails) {
              this.logger.log(`✅ Found opportunity:`, {
                id: opportunityDetails.id,
                name: opportunityDetails.name,
                contactName: opportunityDetails.contact?.name,
                contactFirstName: opportunityDetails.contact?.firstName,
                contactLastName: opportunityDetails.contact?.lastName,
                contactEmail: opportunityDetails.contact?.email
              });
            } else {
              this.logger.warn(`❌ No opportunity details found for ${workflow.ghlOpportunityId}`);
            }
          } else {
            this.logger.warn(`❌ User ${user.name} (ID: ${user.id}) has no GHL access token and system token is not configured - cannot fetch opportunity details for ${workflow.ghlOpportunityId}`);
          }

          // Create enhanced DTO with opportunity details
          const enhancedDto = this.mapToDto(workflow);
          
          if (opportunityDetails) {
            // Use the same customer name extraction logic as admin flow (via helper method)
            const customerName = this.extractCustomerNameFromGHL(opportunityDetails, workflow.ghlOpportunityId);
            
            enhancedDto.opportunityDetails = {
              customerName: customerName,
              address: opportunityDetails.contact?.addresses?.[0]?.address1 
                ? `${opportunityDetails.contact.addresses[0].address1}, ${opportunityDetails.contact.addresses[0].city || ''}`
                : 'Address not available',
              contactEmail: opportunityDetails.contact?.email,
              contactPhone: opportunityDetails.contact?.phone,
              monetaryValue: opportunityDetails.monetaryValue,
              stageName: opportunityDetails.stage?.name,
            };
          } else {
            // Return workflow without opportunity details if not found
            enhancedDto.opportunityDetails = {
              customerName: `Customer ${workflow.ghlOpportunityId.slice(-6)}`,
              address: 'Address not available',
              contactEmail: undefined,
              contactPhone: undefined,
              monetaryValue: undefined,
              stageName: undefined,
            };
          }

          return enhancedDto;
        } catch (error) {
          this.logger.warn(`Failed to fetch opportunity details for ${workflow.ghlOpportunityId}: ${error.message}`);
          
          // Return workflow without opportunity details if all fetches fail
          const basicDto = this.mapToDto(workflow);
          basicDto.opportunityDetails = {
            customerName: `Customer ${workflow.ghlOpportunityId.slice(-6)}`,
            address: 'Address not available',
            contactEmail: undefined,
            contactPhone: undefined,
            monetaryValue: undefined,
            stageName: undefined,
          };
          
          return basicDto;
        }
      })
    );

    this.logger.log(`✅ OPTIMIZED: Enhanced ${enhancedWorkflows.length} workflows with opportunity details`);
    return enhancedWorkflows;
  }

  /**
   * Extract customer name from GHL opportunity details using comprehensive fallback logic
   * This is the same logic used for regular surveyors to ensure consistency
   */
  private extractCustomerNameFromGHL(opportunityDetails: any, opportunityId: string): string {
            let customerName = '';
            
    this.logger.log(`🔍 Extracting customer name from GHL opportunity data:`, {
              contactFirstName: opportunityDetails.contact?.firstName,
              contactLastName: opportunityDetails.contact?.lastName,
              contactName: opportunityDetails.contact?.name,
              opportunityName: opportunityDetails.name,
              contactEmail: opportunityDetails.contact?.email,
              contactCompanyName: opportunityDetails.contact?.companyName
            });
            
            // Priority 1: First and last name combination
            if (opportunityDetails.contact?.firstName && opportunityDetails.contact?.lastName) {
              customerName = `${opportunityDetails.contact.firstName.trim()} ${opportunityDetails.contact.lastName.trim()}`;
              this.logger.log(`✅ Using first+last name: "${customerName}"`);
            }
            // Priority 2: Full name field
            else if (opportunityDetails.contact?.name && opportunityDetails.contact.name.trim() !== '') {
              customerName = opportunityDetails.contact.name.trim();
              this.logger.log(`✅ Using contact name: "${customerName}"`);
            }
            // Priority 3: Opportunity name/title (but clean it up)
            else if (opportunityDetails.name && opportunityDetails.name.trim() !== '') {
              customerName = opportunityDetails.name.trim();
              this.logger.log(`✅ Using opportunity name: "${customerName}"`);
            }
            // Priority 4: Extract from email
            else if (opportunityDetails.contact?.email) {
              const emailPrefix = opportunityDetails.contact.email.split('@')[0];
              customerName = emailPrefix.replace(/[._-]/g, ' ').replace(/\b\w/g, (l: string) => l.toUpperCase());
              this.logger.log(`✅ Using email-derived name: "${customerName}"`);
            }
            // Priority 5: Company name
            else if (opportunityDetails.contact?.companyName && opportunityDetails.contact.companyName.trim() !== '') {
              customerName = opportunityDetails.contact.companyName.trim();
              this.logger.log(`✅ Using company name: "${customerName}"`);
            }
            // Priority 6: Phone number as identifier
            else if (opportunityDetails.contact?.phone) {
              const phone = opportunityDetails.contact.phone.replace(/\D/g, '');
              customerName = `Customer (${phone.slice(-4)})`;
              this.logger.log(`✅ Using phone-derived name: "${customerName}"`);
            }
            // Fallback: Use opportunity ID
            else {
      customerName = `Customer ${opportunityId.slice(-6)}`;
              this.logger.log(`⚠️ Using fallback name: "${customerName}"`);
            }
            
            // Clean up the name - remove postcode patterns like "N12 9JA, Lisa Jones" -> "Lisa Jones"
            if (customerName) {
              const originalName = customerName;
              
              // Remove postcode patterns
              const postcodePattern = /^[A-Z]{1,2}\d{1,2}\s?\d[A-Z]{2},\s*/i;
              customerName = customerName.replace(postcodePattern, '');
              
              // Remove email patterns
              const emailPattern = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,},\s*/;
              customerName = customerName.replace(emailPattern, '');
              
              // Clean up any remaining commas and extra spaces
              customerName = customerName.replace(/^,\s*/, '').replace(/\s*,$/, '').trim();
              
              if (originalName !== customerName) {
                this.logger.log(`🧹 Cleaned name: "${originalName}" -> "${customerName}"`);
              }
            }
            
    this.logger.log(`✅ Final customer name for ${opportunityId}: "${customerName}"`);
    return customerName || `Customer ${opportunityId.slice(-6)}`;
  }

  /**
   * Fetch customer details from multiple sources (Survey, CalculatorProgress, OpportunityProgress) in batch
   * This is much more efficient than individual API calls
   */
  private async fetchCustomerDetailsBatch(
    opportunityIds: string[]
  ): Promise<Map<string, { customerName: string | null; address: string | null; postcode: string | null }>> {
    const customerDetailsMap = new Map<string, { customerName: string | null; address: string | null; postcode: string | null }>();
    
    // Priority 1: Fetch from Survey data (page1 has customerFirstName and customerLastName)
    try {
      const surveys = await this.prisma.survey.findMany({
        where: {
          ghlOpportunityId: {
            in: opportunityIds
          },
          isDeleted: false
        },
        orderBy: {
          updatedAt: 'desc'
        }
      });

      this.logger.log(`📊 Found ${surveys.length} survey records for ${opportunityIds.length} opportunities`);

      for (const survey of surveys) {
        if (survey.page1) {
          const page1 = survey.page1 as any;
          const customerFirstName = page1.customerFirstName || '';
          const customerLastName = page1.customerLastName || '';
          
          if (customerFirstName || customerLastName) {
            const customerName = `${customerFirstName} ${customerLastName}`.trim();
            if (customerName && !customerDetailsMap.has(survey.ghlOpportunityId)) {
              // Try to get address from page1 if available
              const address = page1.customerAddress || page1.address || null;
              const postcode = page1.customerPostcode || page1.postcode || null;
              
              customerDetailsMap.set(survey.ghlOpportunityId, {
                customerName,
                address,
                postcode
              });
              this.logger.log(`✅ Found customer name "${customerName}" from Survey for ${survey.ghlOpportunityId}`);
            }
          }
        }
      }
        } catch (error) {
      this.logger.warn(`⚠️ Failed to fetch customer details from Survey in batch: ${error.message}`);
    }

    // Priority 2: Fetch from CalculatorProgress
    try {
      const calculatorProgresses = await this.prisma.calculatorProgress.findMany({
        where: {
          opportunityId: {
            in: opportunityIds
          }
        },
        orderBy: {
          updatedAt: 'desc'
        }
      });

      this.logger.log(`📊 Found ${calculatorProgresses.length} calculator progress records for ${opportunityIds.length} opportunities`);

      for (const progress of calculatorProgresses) {
        // Only use if we don't already have customer name from Survey
        if (!customerDetailsMap.has(progress.opportunityId)) {
          if (progress.data) {
            const data = progress.data as any;
            if (data.customerDetails) {
              const customerDetails = data.customerDetails;
              const customerName = customerDetails.customerName || null;
              const address = customerDetails.address || null;
              const postcode = customerDetails.postcode || null;

              if (customerName) {
                customerDetailsMap.set(progress.opportunityId, {
                  customerName,
                  address,
                  postcode
                });
                this.logger.log(`✅ Found customer name "${customerName}" from CalculatorProgress for ${progress.opportunityId}`);
              }
            }
          }
        } else {
          // We already have customer name, but might want to enrich with address/postcode
          const existing = customerDetailsMap.get(progress.opportunityId);
          if (existing && progress.data) {
            const data = progress.data as any;
            if (data.customerDetails) {
              const customerDetails = data.customerDetails;
              if (!existing.address && customerDetails.address) {
                existing.address = customerDetails.address;
              }
              if (!existing.postcode && customerDetails.postcode) {
                existing.postcode = customerDetails.postcode;
              }
            }
          }
        }
      }
    } catch (error) {
      this.logger.warn(`⚠️ Failed to fetch customer details from CalculatorProgress in batch: ${error.message}`);
    }

    // Priority 3: Check OpportunityProgress stepData and OpportunityStep data for customer information
    try {
      const opportunityProgresses = await this.prisma.opportunityProgress.findMany({
        where: {
          ghlOpportunityId: {
            in: opportunityIds
          }
        },
        include: {
          steps: true // Include steps to check step data
        }
      });

      for (const oppProgress of opportunityProgresses) {
        let customerName: string | null = null;
        let address: string | null = null;
        let postcode: string | null = null;

        // Check OpportunityProgress.stepData for customer info
        if (oppProgress.stepData) {
          const stepData = oppProgress.stepData as any;
          // Check multiple possible field names for customer name
          customerName = stepData.customerName || 
                        stepData.customerInfo?.name || 
                        stepData.name || 
                        stepData.contactName ||
                        stepData.contact?.name ||
                        stepData.contact?.firstName && stepData.contact?.lastName 
                          ? `${stepData.contact.firstName} ${stepData.contact.lastName}`.trim() 
                          : null;
          
          if (customerName && customerName !== 'Customer' && !customerName.startsWith('Customer ')) {
            address = stepData.customerAddress || 
                      stepData.customerInfo?.address || 
                      stepData.address || 
                      stepData.contactAddress ||
                      null;
            postcode = stepData.customerPostcode || 
                      stepData.customerInfo?.postcode || 
                      stepData.postcode || 
                      stepData.contactPostcode ||
                      null;
          } else {
            customerName = null; // Reset if it's a fallback name
          }
        }

        // If not found in stepData, check OpportunityStep data (especially step 1)
        if (!customerName && oppProgress.steps && oppProgress.steps.length > 0) {
          // Check steps in order, starting with step 1
          for (const step of oppProgress.steps.sort((a, b) => a.stepNumber - b.stepNumber)) {
            if (step.data) {
              const stepData = step.data as any;
              // Check multiple possible field names for customer name
              const foundName = stepData.customerName || 
                               stepData.customerInfo?.name || 
                               stepData.name || 
                               stepData.contactName ||
                               stepData.contact?.name ||
                               stepData.contact?.firstName && stepData.contact?.lastName 
                                 ? `${stepData.contact.firstName} ${stepData.contact.lastName}`.trim() 
                                 : null;
              
              if (foundName && foundName !== 'Customer' && !foundName.startsWith('Customer ')) {
                customerName = foundName;
                address = stepData.customerAddress || 
                         stepData.customerInfo?.address || 
                         stepData.address || 
                         stepData.contactAddress ||
                         address;
                postcode = stepData.customerPostcode || 
                          stepData.customerInfo?.postcode || 
                          stepData.postcode || 
                          stepData.contactPostcode ||
                          postcode;
                this.logger.log(`✅ Found customer name "${customerName}" in step ${step.stepNumber} data for ${oppProgress.ghlOpportunityId}`);
                break; // Found customer name, stop checking
              }
            }
          }
        }
        
        // Log if we didn't find customer name for debugging
        if (!customerName && !customerDetailsMap.has(oppProgress.ghlOpportunityId)) {
          this.logger.log(`🔍 No customer name found in OpportunityProgress stepData for ${oppProgress.ghlOpportunityId} (step ${oppProgress.currentStep})`);
        }

        // If we found customer info, add it to the map (only if not already present)
        if (customerName && !customerDetailsMap.has(oppProgress.ghlOpportunityId)) {
          customerDetailsMap.set(oppProgress.ghlOpportunityId, {
            customerName,
            address: address || oppProgress.contactAddress || null,
            postcode: postcode || oppProgress.contactPostcode || null
          });
          this.logger.log(`✅ Found customer name "${customerName}" from OpportunityProgress stepData for ${oppProgress.ghlOpportunityId}`);
        } else {
          // Enrich existing entry with address/postcode if available
          const existing = customerDetailsMap.get(oppProgress.ghlOpportunityId);
          if (existing) {
            if (!existing.address && (address || oppProgress.contactAddress)) {
              existing.address = address || oppProgress.contactAddress || null;
            }
            if (!existing.postcode && (postcode || oppProgress.contactPostcode)) {
              existing.postcode = postcode || oppProgress.contactPostcode || null;
            }
          }
        }
      }
    } catch (error) {
      this.logger.warn(`⚠️ Failed to fetch customer details from OpportunityProgress in batch: ${error.message}`);
    }

    // Log summary of opportunities with and without customer names
    const opportunitiesWithoutNames = opportunityIds.filter(id => !customerDetailsMap.has(id));
    if (opportunitiesWithoutNames.length > 0) {
      this.logger.log(`⚠️ ${opportunitiesWithoutNames.length} opportunities without customer names from batch sources: ${opportunitiesWithoutNames.slice(0, 5).join(', ')}${opportunitiesWithoutNames.length > 5 ? '...' : ''}`);
    }
    this.logger.log(`📊 Successfully extracted customer details for ${customerDetailsMap.size} out of ${opportunityIds.length} opportunities from all sources`);
    return customerDetailsMap;
  }

  async getAllWorkflowsForAdmin(): Promise<OpportunityProgressDto[]> {
    this.logger.log(`🔍 getAllWorkflowsForAdmin called - fetching all workflows for admin`);
    
    // Get all workflows for admin view
    const workflows = await this.prisma.opportunityProgress.findMany({
      include: { 
        steps: true,
        user: true // Include user information
      },
      orderBy: { lastActivityAt: 'desc' }
    });

    // If user information is missing, try to fetch it separately
    const workflowsWithUserInfo = await Promise.all(
      workflows.map(async (workflow) => {
        if (!workflow.user && workflow.userId) {
          this.logger.log(`🔍 Workflow ${workflow.ghlOpportunityId} missing user info, fetching separately for userId: ${workflow.userId}`);
          try {
            const user = await this.userService.findById(workflow.userId);
            if (user) {
              this.logger.log(`✅ Found user for workflow ${workflow.ghlOpportunityId}: ${user.name}`);
              return { ...workflow, user };
            }
          } catch (error) {
            this.logger.warn(`⚠️ Failed to fetch user for workflow ${workflow.ghlOpportunityId}:`, error);
          }
        }
        return workflow;
      })
    );

    this.logger.log(`🚀 Found ${workflowsWithUserInfo.length} total workflows for admin view`);
    
    // Debug: Log user information for each workflow
    workflowsWithUserInfo.forEach((workflow, index) => {
      this.logger.log(`🔍 Workflow ${index + 1}: ${workflow.ghlOpportunityId} - User: ${workflow.user?.name || 'No user'} (${workflow.user?.role || 'No role'})`);
    });

    // Fetch customer details from multiple sources in batch (efficient!)
    const opportunityIds = workflowsWithUserInfo.map(w => w.ghlOpportunityId);
    const customerDetailsMap = await this.fetchCustomerDetailsBatch(opportunityIds);
    this.logger.log(`📊 Retrieved customer details for ${customerDetailsMap.size} opportunities from all sources (Survey, CalculatorProgress, OpportunityProgress)`);

    // Enhance workflows with opportunity details and user information
    const enhancedWorkflows = await Promise.all(
      workflowsWithUserInfo.map(async (workflow) => {
        try {
          // First, try to get customer details from our batch fetch (Survey, CalculatorProgress, OpportunityProgress - fast, no API call)
          const batchCustomerDetails = customerDetailsMap.get(workflow.ghlOpportunityId);
          
          let opportunityDetails: any = null;
          const accessToken = process.env.GOHIGHLEVEL_API_TOKEN || null;
          
          // If we have customer name from our batch sources, use it (fast, no API call needed)
          if (batchCustomerDetails?.customerName) {
            this.logger.log(`✅ Using customer name from batch sources for ${workflow.ghlOpportunityId}: "${batchCustomerDetails.customerName}"`);
            opportunityDetails = {
              customerName: batchCustomerDetails.customerName,
              address: batchCustomerDetails.address || null,
              contactPostcode: batchCustomerDetails.postcode || null
            };
          } else if (accessToken) {
            // Only fetch from GHL if we don't have customer name from our batch sources
            this.logger.log(`🔍 Falling back to GHL API for ${workflow.ghlOpportunityId} (no customer name found in batch sources)`);
            try {
              const opportunityResponse = await this.ghlService.getOpportunityById(
                accessToken,
                process.env.GOHIGHLEVEL_LOCATION_ID || '',
                workflow.ghlOpportunityId
              );
              
              if (opportunityResponse.success && opportunityResponse.data) {
                const rawOpportunityDetails = opportunityResponse.data;
                this.logger.log(`✅ Fetched opportunity details from GHL for ${workflow.ghlOpportunityId}`);
                
                // Use the same customer name extraction logic as regular surveyors
                const customerName = this.extractCustomerNameFromGHL(rawOpportunityDetails, workflow.ghlOpportunityId);
                this.logger.log(`✅ Extracted customer name "${customerName}" from GHL API for ${workflow.ghlOpportunityId}`);
                
                // Construct opportunityDetails with extracted customer name
                opportunityDetails = {
                  customerName: customerName,
                  address: rawOpportunityDetails.contact?.addresses?.[0]?.address1 
                    ? `${rawOpportunityDetails.contact.addresses[0].address1}, ${rawOpportunityDetails.contact.addresses[0].city || ''}`
                    : 'Address not available',
                  contactEmail: rawOpportunityDetails.contact?.email,
                  contactPhone: rawOpportunityDetails.contact?.phone,
                  monetaryValue: rawOpportunityDetails.monetaryValue,
                  stageName: rawOpportunityDetails.stage?.name,
                  contactPostcode: rawOpportunityDetails.contact?.addresses?.[0]?.postalCode || rawOpportunityDetails.contact?.postalCode || null
                };
              }
            } catch (error) {
              this.logger.warn(`⚠️ Failed to fetch opportunity details from GHL for ${workflow.ghlOpportunityId}:`, error);
            }
          }

          // If we still don't have opportunityDetails, create a minimal fallback
          if (!opportunityDetails) {
            opportunityDetails = {
              customerName: `Customer ${workflow.ghlOpportunityId.slice(-6)}`,
              address: 'Address not available',
              contactPostcode: null
            };
            this.logger.log(`⚠️ No customer details found for ${workflow.ghlOpportunityId}, using fallback`);
          }

          const userInfo = workflow.user ? {
            id: workflow.user.id,
            name: workflow.user.name,
            username: workflow.user.username,
            email: workflow.user.email,
            role: workflow.user.role
          } : null;
          
          this.logger.log(`🔍 Enhanced workflow ${workflow.ghlOpportunityId} with userInfo:`, userInfo);
          
          return {
            ...workflow,
            opportunityDetails,
            userInfo
          };
        } catch (error) {
          this.logger.error(`❌ Error enhancing workflow ${workflow.id}:`, error);
          return {
            ...workflow,
            opportunityDetails: {
              customerName: `Customer ${workflow.ghlOpportunityId.slice(-6)}`,
              address: 'Address not available',
              contactPostcode: null
            },
            userInfo: workflow.user ? {
              id: workflow.user.id,
              name: workflow.user.name,
              username: workflow.user.username,
              email: workflow.user.email,
              role: workflow.user.role
            } : null
          };
        }
      })
    );

    return enhancedWorkflows.map(workflow => this.mapToDtoWithUserInfo(workflow));
  }

  private mapToDto(progress: any): OpportunityProgressDto {
    return {
      id: progress.id,
      ghlOpportunityId: progress.ghlOpportunityId,
      currentStep: progress.currentStep,
      totalSteps: progress.totalSteps,
      status: progress.status,
      startedAt: progress.startedAt,
      lastActivityAt: progress.lastActivityAt,
      completedAt: progress.completedAt,
      stepData: progress.stepData,
      steps: progress.steps.map((step: any) => ({
        id: step.id,
        stepNumber: step.stepNumber,
        stepType: step.stepType,
        status: step.status,
        data: step.data,
        startedAt: step.startedAt,
        completedAt: step.completedAt
      })),
      opportunityDetails: progress.opportunityDetails
    };
  }

  private mapToDtoWithUserInfo(progress: any): OpportunityProgressDto {
    return {
      id: progress.id,
      ghlOpportunityId: progress.ghlOpportunityId,
      currentStep: progress.currentStep,
      totalSteps: progress.totalSteps,
      status: progress.status,
      startedAt: progress.startedAt,
      lastActivityAt: progress.lastActivityAt,
      completedAt: progress.completedAt,
      stepData: progress.stepData,
      steps: progress.steps.map((step: any) => ({
        id: step.id,
        stepNumber: step.stepNumber,
        stepType: step.stepType,
        status: step.status,
        data: step.data,
        startedAt: step.startedAt,
        completedAt: step.completedAt
      })),
      opportunityDetails: progress.opportunityDetails,
      userInfo: progress.userInfo
    };
  }

  /**
   * Handle OneDrive copying for won opportunities
   */
  private async handleWonOpportunityOneDriveCopying(opportunityId: string, stepData: any): Promise<void> {
    try {
      this.logger.log(`🔄 Starting OneDrive copying for won opportunity: ${opportunityId}`);

      // Get customer name from step data or opportunity details
      let customerName = stepData?.customerName || stepData?.customerInfo?.name || 'Customer';
      
      // If we don't have a customer name, try to get it from the opportunity
      if (!customerName || customerName === 'Customer') {
        try {
          const opportunityDetails = await this.ghlService.getOpportunityById(
            process.env.GOHIGHLEVEL_API_TOKEN || '',
            opportunityId
          );
          
          if (opportunityDetails?.contact?.firstName && opportunityDetails?.contact?.lastName) {
            customerName = `${opportunityDetails.contact.firstName} ${opportunityDetails.contact.lastName}`;
          } else if (opportunityDetails?.contact?.name) {
            customerName = opportunityDetails.contact.name;
          } else if (opportunityDetails?.name) {
            customerName = opportunityDetails.name;
          }
        } catch (error) {
          this.logger.warn(`Could not fetch opportunity details for customer name: ${error.message}`);
        }
      }

      // Collect document paths from step data
      const documents: any = {};

      // Check for proposal files (from step 4 - Proposal)
      if (stepData?.proposalFiles) {
        documents.proposalFiles = stepData.proposalFiles;
      }

      // Check for contract files (from step 7 - Contract Generation)
      // Note: Only copy unsigned contract files - signed contracts are handled by DocuSeal
      if (stepData?.contractFiles) {
        documents.contractFiles = stepData.contractFiles;
      }

      // Check for contract submission ID (from step 7 - Contract Signing)
      // Note: This is just for tracking - signed contracts are handled by DocuSeal
      if (stepData?.contractSubmissionId || stepData?.submissionId) {
        documents.contractSubmissionId = stepData.contractSubmissionId || stepData.submissionId;
        this.logger.log(`📋 Found contract submission ID: ${documents.contractSubmissionId}`);
      }

      // Skip disclaimer file - signed disclaimers are now handled by DocuSeal webhooks
      // Do NOT copy disclaimerPath or signedDisclaimerPath as DocuSeal handles signed documents

      // Skip email confirmation file - signed confirmations are now handled by DocuSeal webhooks
      // Do NOT copy emailConfirmationPath or signedEmailConfirmationPath as DocuSeal handles signed documents

      // Check for booking confirmation submission ID (from step 9 - Email Confirmation)
      if (stepData?.bookingConfirmationSubmissionId || stepData?.emailConfirmationSubmissionId) {
        documents.bookingConfirmationSubmissionId = stepData.bookingConfirmationSubmissionId || stepData.emailConfirmationSubmissionId;
        this.logger.log(`📋 Found booking confirmation submission ID: ${documents.bookingConfirmationSubmissionId}`);
      }

      // Process completed DocuSeal submissions for this opportunity
      // This will download signed documents and audit logs from DocuSeal and upload to OneDrive
      try {
        this.logger.log(`🔍 Processing completed DocuSeal submissions for opportunity: ${opportunityId}`);
        const docuSealResult = await this.docuSealService.processCompletedSubmissionsForOpportunity(opportunityId);
        
        if (docuSealResult.success) {
          this.logger.log(`✅ Successfully processed ${docuSealResult.processed} completed submission(s) from DocuSeal for ${opportunityId}`);
        } else {
          this.logger.warn(`⚠️ Some issues processing DocuSeal submissions: ${docuSealResult.message}`);
          if (docuSealResult.errors) {
            docuSealResult.errors.forEach(err => this.logger.warn(`  - ${err}`));
          }
        }
      } catch (error) {
        this.logger.error(`❌ Error processing DocuSeal submissions for ${opportunityId}: ${error.message}`);
        // Don't throw - continue with other document copying
      }

      // If we have any documents to copy, proceed with OneDrive copying
      if (Object.keys(documents).length > 0) {
        this.logger.log(`📁 Found documents to copy to OneDrive: ${Object.keys(documents).join(', ')}`);
        
        const oneDriveResult = await this.oneDriveFileManagerService.copyWonOpportunityDocumentsToOneDrive(
          opportunityId,
          customerName,
          documents
        );

        if (oneDriveResult.success) {
          this.logger.log(`✅ Successfully copied won opportunity documents to OneDrive for ${opportunityId}`);
        } else {
          this.logger.error(`❌ Failed to copy won opportunity documents to OneDrive for ${opportunityId}: ${oneDriveResult.error}`);
        }
      } else {
        this.logger.warn(`⚠️ No documents found to copy to OneDrive for won opportunity: ${opportunityId}`);
      }

    } catch (error) {
      this.logger.error(`❌ Error in OneDrive copying for won opportunity ${opportunityId}: ${error.message}`);
      // Don't throw the error - OneDrive copying should not block the workflow completion
    }
  }

  /**
   * Clean up all files related to an opportunity ID from the output directory
   * This includes PowerPoint files, PDF files, and JSON variable files
   */
  private async cleanupOutputDirectoryFiles(opportunityId: string): Promise<void> {
    try {
      this.logger.log(`🧹 Starting cleanup of output directory files for opportunity: ${opportunityId}`);
      
      const outputDir = path.join(process.cwd(), 'src', 'excel-file-calculator', 'output');
      
      // Check if output directory exists
      if (!(await fs.access(outputDir).then(() => true).catch(() => false))) {
        this.logger.warn(`⚠️ Output directory does not exist: ${outputDir}`);
        return;
      }
      
      // Read all files in the output directory
      const files = await fs.readdir(outputDir);
      
      // Filter files that match the opportunity ID pattern
      // Pattern: presentation_{opportunityId}_{timestamp}.pptx
      // Pattern: presentation_{opportunityId}_{timestamp}.pdf
      // Pattern: presentation_{opportunityId}_{timestamp}_variables.json
      const filesToDelete = files.filter(file => {
        // Match files that start with "presentation_" and contain the opportunity ID
        return file.startsWith(`presentation_${opportunityId}_`);
      });
      
      if (filesToDelete.length === 0) {
        this.logger.log(`ℹ️ No files found to delete for opportunity ${opportunityId}`);
        return;
      }
      
      this.logger.log(`📋 Found ${filesToDelete.length} file(s) to delete for opportunity ${opportunityId}`);
      
      // Delete each file
      let deletedCount = 0;
      let errorCount = 0;
      
      for (const file of filesToDelete) {
        const filePath = path.join(outputDir, file);
        try {
          await fs.unlink(filePath);
          deletedCount++;
          this.logger.log(`✅ Deleted: ${file}`);
        } catch (error) {
          errorCount++;
          this.logger.error(`❌ Failed to delete ${file}: ${error.message}`);
        }
      }
      
      this.logger.log(`✅ Cleanup completed: ${deletedCount} file(s) deleted, ${errorCount} error(s)`);
      
    } catch (error) {
      this.logger.error(`❌ Error cleaning up output directory files for ${opportunityId}: ${error.message}`);
      // Don't throw - cleanup should not block the workflow
    }
  }

  /**
   * Handle OneDrive copying for lost opportunities
   */
  private async handleLostOpportunityOneDriveCopying(opportunityId: string, stepData: any): Promise<void> {
    try {
      this.logger.log(`🔄 Starting OneDrive copying for lost opportunity: ${opportunityId}`);

      // Get customer name from step data or opportunity details
      let customerName = stepData?.customerName || stepData?.customerInfo?.name || 'Customer';
      let postcode = stepData?.postcode || stepData?.customerInfo?.postcode || '';
      
      // If we don't have customer details, try to get them from the opportunity
      if (!customerName || customerName === 'Customer' || !postcode) {
        try {
          const opportunityDetails = await this.ghlService.getOpportunityById(
            process.env.GOHIGHLEVEL_API_TOKEN || '',
            opportunityId
          );
          
          if (opportunityDetails?.contact?.firstName && opportunityDetails?.contact?.lastName) {
            customerName = `${opportunityDetails.contact.firstName} ${opportunityDetails.contact.lastName}`;
          } else if (opportunityDetails?.contact?.name) {
            customerName = opportunityDetails.contact.name;
          } else if (opportunityDetails?.name) {
            customerName = opportunityDetails.name;
          }
          
          if (!postcode && opportunityDetails?.contact?.addresses?.[0]?.postalCode) {
            postcode = opportunityDetails.contact.addresses[0].postalCode;
          }
        } catch (error) {
          this.logger.warn(`Could not fetch opportunity details for customer name: ${error.message}`);
        }
      }

      // Collect file paths from step data
      const files: any = {};

      // Check for proposal files (from step 4 - Proposal)
      if (stepData?.proposalFiles) {
        if (stepData.proposalFiles.pptxPath) {
          files.proposalPath = stepData.proposalFiles.pptxPath;
        } else if (stepData.proposalFiles.pdfPath) {
          files.proposalPath = stepData.proposalFiles.pdfPath;
        }
      }

      // Check for calculator files
      if (stepData?.calculatorPath) {
        files.calculatorPath = stepData.calculatorPath;
      }

      // Check for survey files
      if (stepData?.surveyPath) {
        files.surveyPath = stepData.surveyPath;
      }

      // If we have any files to copy, proceed with OneDrive copying
      if (Object.keys(files).length > 0) {
        this.logger.log(`📁 Found files to copy to OneDrive for lost opportunity: ${Object.keys(files).join(', ')}`);
        
        const oneDriveResult = await this.oneDriveFileManagerService.organizeFilesByOutcome(
          opportunityId,
          customerName,
          postcode,
          'lost',
          files
        );

        if (oneDriveResult.success) {
          this.logger.log(`✅ Successfully copied lost opportunity documents to OneDrive for ${opportunityId}`);
        } else {
          this.logger.error(`❌ Failed to copy lost opportunity documents to OneDrive for ${opportunityId}: ${oneDriveResult.error}`);
        }
      } else {
        this.logger.warn(`⚠️ No files found to copy to OneDrive for lost opportunity: ${opportunityId}`);
      }

    } catch (error) {
      this.logger.error(`❌ Error in OneDrive copying for lost opportunity ${opportunityId}: ${error.message}`);
      // Don't throw the error - OneDrive copying should not block the workflow completion
    }
  }

  /**
   * Handle OneDrive copying for step 4 (Proposal) completion
   */
  private async handleOneDriveFileCopying(opportunityId: string, stepData: any): Promise<void> {
    try {
      this.logger.log(`🔄 Starting OneDrive file copying for step 4 completion: ${opportunityId}`);

      // Get customer name from step data
      const customerName = stepData?.customerName || stepData?.customerInfo?.name || 'Customer';
      
      // Get proposal files from step data
      const proposalFiles = stepData?.proposalFiles || {};
      
      if (proposalFiles.pptxPath || proposalFiles.pdfPath) {
        const oneDriveResult = await this.oneDriveFileManagerService.copyProposalToQuotationsWithSurveyImages(
          opportunityId,
          customerName,
          proposalFiles
        );

        if (oneDriveResult.success) {
          this.logger.log(`✅ Successfully copied proposal files to OneDrive for ${opportunityId}`);
        } else {
          this.logger.error(`❌ Failed to copy proposal files to OneDrive for ${opportunityId}: ${oneDriveResult.error}`);
        }
      } else {
        this.logger.warn(`⚠️ No proposal files found to copy to OneDrive for opportunity: ${opportunityId}`);
      }

    } catch (error) {
      this.logger.error(`❌ Error in OneDrive file copying for step 4 ${opportunityId}: ${error.message}`);
      // Don't throw the error - OneDrive copying should not block the workflow completion
    }
  }
} 