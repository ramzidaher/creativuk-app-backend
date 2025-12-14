import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { OneDriveFileManagerService } from '../onedrive/onedrive-file-manager.service';
import * as path from 'path';
import * as fs from 'fs';

@Injectable()
export class AdminOpportunityDetailsService {
  private readonly logger = new Logger(AdminOpportunityDetailsService.name);
  
  private readonly OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'opportunities');
  private readonly EPVS_OPPORTUNITIES_FOLDER = path.join(process.cwd(), 'src', 'excel-file-calculator', 'epvs-opportunities');

  constructor(
    private readonly prisma: PrismaService,
    private readonly oneDriveFileManagerService: OneDriveFileManagerService
  ) {}

  /**
   * Get all users with their opportunities and all related details
   */
  async getAllUsersWithOpportunities() {
    try {
      this.logger.log('Fetching all users with their opportunities...');

      // Get all users
      const users = await this.prisma.user.findMany({
        select: {
          id: true,
          name: true,
          email: true,
          username: true,
          role: true,
          status: true,
          ghlUserId: true,
          ghlUserName: true,
          createdAt: true,
          lastLoginAt: true
        },
        orderBy: { createdAt: 'desc' }
      });

      // Get all opportunity progress records
      const allOpportunities = await this.prisma.opportunityProgress.findMany({
        include: {
          user: {
            select: {
              id: true,
              name: true,
              email: true,
              username: true,
              ghlUserId: true
            }
          },
          steps: {
            orderBy: { stepNumber: 'asc' }
          }
        },
        orderBy: { createdAt: 'desc' }
      });

      // Get all calculator progress records to extract customer names
      const allCalculatorProgress = await this.prisma.calculatorProgress.findMany({
        select: {
          opportunityId: true,
          calculatorType: true,
          data: true
        }
      });

      // Create a map of opportunityId -> customerName
      const customerNameMap = new Map<string, string>();
      for (const calc of allCalculatorProgress) {
        const data = calc.data as any;
        if (data?.customerDetails?.customerName && !customerNameMap.has(calc.opportunityId)) {
          customerNameMap.set(calc.opportunityId, data.customerDetails.customerName);
        }
      }

      // Group opportunities by user
      const usersWithOpportunities = users.map(user => {
        const userOpportunities = allOpportunities.filter(
          opp => opp.userId === user.id
        );

        return {
          user: {
            id: user.id,
            name: user.name,
            email: user.email,
            username: user.username,
            role: user.role,
            status: user.status,
            ghlUserId: user.ghlUserId,
            ghlUserName: user.ghlUserName,
            createdAt: user.createdAt,
            lastLoginAt: user.lastLoginAt
          },
          opportunities: userOpportunities.map(opp => ({
            id: opp.id,
            ghlOpportunityId: opp.ghlOpportunityId,
            customerName: customerNameMap.get(opp.ghlOpportunityId) || null,
            currentStep: opp.currentStep,
            totalSteps: opp.totalSteps,
            status: opp.status,
            contactAddress: opp.contactAddress,
            contactPostcode: opp.contactPostcode,
            startedAt: opp.startedAt,
            lastActivityAt: opp.lastActivityAt,
            completedAt: opp.completedAt,
            createdAt: opp.createdAt,
            updatedAt: opp.updatedAt,
            steps: opp.steps
          })),
          totalOpportunities: userOpportunities.length
        };
      });

      this.logger.log(`Found ${users.length} users with ${allOpportunities.length} total opportunities`);
      return usersWithOpportunities;

    } catch (error) {
      this.logger.error('Error fetching all users with opportunities:', error);
      throw new Error(`Failed to fetch users with opportunities: ${error.message}`);
    }
  }

  /**
   * Get complete opportunity details including all related data
   */
  async getOpportunityDetails(opportunityId: string) {
    try {
      this.logger.log(`Fetching complete details for opportunity: ${opportunityId}`);

      // Get opportunity progress
      const opportunityProgress = await this.prisma.opportunityProgress.findUnique({
        where: { ghlOpportunityId: opportunityId },
        include: {
          user: {
            select: {
              id: true,
              name: true,
              email: true,
              username: true,
              ghlUserId: true,
              ghlUserName: true
            }
          },
          steps: {
            orderBy: { stepNumber: 'asc' }
          }
        }
      });

      if (!opportunityProgress) {
        throw new Error(`Opportunity not found: ${opportunityId}`);
      }

      // Get all related data in parallel (read-only operations)
      const [
        surveyData,
        surveyImages,
        openSolarProject,
        calculatorProgress,
        excelFiles,
        pdfFiles
      ] = await Promise.all([
        this.getSurveyData(opportunityId),
        this.getSurveyImages(opportunityId),
        this.getOpenSolarProject(opportunityId),
        this.getCalculatorProgress(opportunityId),
        this.getExcelFiles(opportunityId),
        this.getPdfFiles(opportunityId)
      ]);

      // Extract customer name from calculator data
      let customerName: string | null = null;
      if (calculatorProgress?.calculators) {
        // Try to get customer name from any calculator type
        const calculatorTypes = ['off-peak', 'flux', 'epvs'];
        for (const calcType of calculatorTypes) {
          const calc = calculatorProgress.calculators[calcType];
          if (calc?.data?.customerDetails?.customerName) {
            customerName = calc.data.customerDetails.customerName;
            break;
          }
        }
      }

      return {
        opportunity: {
          id: opportunityProgress.id,
          ghlOpportunityId: opportunityProgress.ghlOpportunityId,
          user: opportunityProgress.user,
          customerName: customerName,
          currentStep: opportunityProgress.currentStep,
          totalSteps: opportunityProgress.totalSteps,
          status: opportunityProgress.status,
          contactAddress: opportunityProgress.contactAddress,
          contactPostcode: opportunityProgress.contactPostcode,
          startedAt: opportunityProgress.startedAt,
          lastActivityAt: opportunityProgress.lastActivityAt,
          completedAt: opportunityProgress.completedAt,
          createdAt: opportunityProgress.createdAt,
          updatedAt: opportunityProgress.updatedAt,
          stepData: opportunityProgress.stepData,
          steps: opportunityProgress.steps
        },
        survey: {
          data: surveyData,
          images: surveyImages
        },
        openSolar: openSolarProject,
        calculator: calculatorProgress,
        files: {
          excel: excelFiles,
          pdf: pdfFiles
        }
      };

    } catch (error) {
      this.logger.error(`Error fetching opportunity details for ${opportunityId}:`, error);
      throw new Error(`Failed to fetch opportunity details: ${error.message}`);
    }
  }

  /**
   * Get survey data for an opportunity
   */
  private async getSurveyData(opportunityId: string) {
    try {
      const survey = await this.prisma.survey.findUnique({
        where: { ghlOpportunityId: opportunityId },
        include: {
          createdByUser: {
            select: {
              id: true,
              name: true,
              email: true
            }
          },
          updatedByUser: {
            select: {
              id: true,
              name: true,
              email: true
            }
          }
        }
      });

      if (!survey) {
        return null;
      }

      return {
        id: survey.id,
        ghlOpportunityId: survey.ghlOpportunityId,
        ghlUserId: survey.ghlUserId,
        page1: survey.page1,
        page2: survey.page2,
        page3: survey.page3,
        page4: survey.page4,
        page5: survey.page5,
        page6: survey.page6,
        page7: survey.page7,
        page8: survey.page8,
        status: survey.status,
        eligibilityScore: survey.eligibilityScore,
        rejectionReason: survey.rejectionReason,
        createdAt: survey.createdAt,
        updatedAt: survey.updatedAt,
        submittedAt: survey.submittedAt,
        approvedAt: survey.approvedAt,
        rejectedAt: survey.rejectedAt,
        createdBy: survey.createdByUser,
        updatedBy: survey.updatedByUser
      };
    } catch (error) {
      this.logger.error(`Error fetching survey data for ${opportunityId}:`, error);
      return null;
    }
  }

  /**
   * Get survey images for an opportunity
   */
  private async getSurveyImages(opportunityId: string) {
    try {
      const images = await this.oneDriveFileManagerService.getSurveyImagesForOpportunity(opportunityId);
      return images.map(img => ({
        fieldName: img.fieldName,
        fileName: img.fileName,
        filePath: img.filePath,
        originalName: img.originalName,
        url: img.filePath // Cloudinary URL or file path
      }));
    } catch (error) {
      this.logger.error(`Error fetching survey images for ${opportunityId}:`, error);
      return [];
    }
  }

  /**
   * Get OpenSolar project data
   */
  private async getOpenSolarProject(opportunityId: string) {
    try {
      const project = await this.prisma.openSolarProject.findUnique({
        where: { opportunityId: opportunityId }
      });

      if (!project) {
        return null;
      }

      return {
        id: project.id,
        opportunityId: project.opportunityId,
        opensolarProjectId: project.opensolarProjectId,
        projectName: project.projectName,
        address: project.address,
        systems: project.systems,
        createdAt: project.createdAt,
        updatedAt: project.updatedAt
      };
    } catch (error) {
      this.logger.error(`Error fetching OpenSolar project for ${opportunityId}:`, error);
      return null;
    }
  }

  /**
   * Get calculator progress data (off-peak, flux, epvs)
   */
  private async getCalculatorProgress(opportunityId: string) {
    try {
      // Get all calculator progress records for this opportunity
      const calculatorProgressRecords = await this.prisma.calculatorProgress.findMany({
        where: { opportunityId: opportunityId },
        include: {
          user: {
            select: {
              id: true,
              name: true,
              email: true
            }
          }
        },
        orderBy: { updatedAt: 'desc' }
      });

      if (calculatorProgressRecords.length === 0) {
        return null;
      }

      // Group by calculator type
      const calculators: any = {};

      for (const record of calculatorProgressRecords) {
        const calculatorType = record.calculatorType;
        const data = record.data as any;

        calculators[calculatorType] = {
          id: record.id,
          opportunityId: record.opportunityId,
          calculatorType: record.calculatorType,
          userId: record.userId,
          user: record.user,
          data: {
            // Template Selection Data
            templateSelection: data?.templateSelection || null,
            // Radio Button Selections
            radioButtonSelections: data?.radioButtonSelections || null,
            // Dynamic Inputs Data
            dynamicInputs: data?.dynamicInputs || null,
            // Arrays Data
            arraysData: data?.arraysData || null,
            // Pricing Data
            pricingData: data?.pricingData || null,
            // Customer Details
            customerDetails: data?.customerDetails || null,
            // Progress tracking
            currentStep: data?.currentStep || null,
            completedSteps: data?.completedSteps || null,
            // Raw data for complete reference
            rawData: data
          },
          createdAt: record.createdAt,
          updatedAt: record.updatedAt
        };
      }

      return {
        hasOffPeak: !!calculators['off-peak'],
        hasFlux: !!calculators['flux'],
        hasEpvs: !!calculators['epvs'],
        calculators: calculators
      };
    } catch (error) {
      this.logger.error(`Error fetching calculator progress for ${opportunityId}:`, error);
      return null;
    }
  }

  /**
   * Get Excel files for an opportunity
   */
  private async getExcelFiles(opportunityId: string) {
    try {
      const excelFiles: any[] = [];

      // Search in regular opportunities folder
      if (fs.existsSync(this.OPPORTUNITIES_FOLDER)) {
        const files = fs.readdirSync(this.OPPORTUNITIES_FOLDER);
        const matchingFiles = files.filter(file => 
          file.includes(opportunityId) && 
          (file.endsWith('.xlsm') || file.endsWith('.xlsx'))
        );

        for (const file of matchingFiles) {
          const filePath = path.join(this.OPPORTUNITIES_FOLDER, file);
          const stats = fs.statSync(filePath);
          excelFiles.push({
            fileName: file,
            filePath: filePath,
            calculatorType: 'off-peak',
            size: stats.size,
            modifiedAt: stats.mtime,
            createdAt: stats.birthtime
          });
        }
      }

      // Search in EPVS opportunities folder
      if (fs.existsSync(this.EPVS_OPPORTUNITIES_FOLDER)) {
        const files = fs.readdirSync(this.EPVS_OPPORTUNITIES_FOLDER);
        const matchingFiles = files.filter(file => 
          file.includes(opportunityId) && 
          (file.endsWith('.xlsm') || file.endsWith('.xlsx'))
        );

        for (const file of matchingFiles) {
          const filePath = path.join(this.EPVS_OPPORTUNITIES_FOLDER, file);
          const stats = fs.statSync(filePath);
          excelFiles.push({
            fileName: file,
            filePath: filePath,
            calculatorType: file.includes('EPVS') ? 'epvs' : 'flux',
            size: stats.size,
            modifiedAt: stats.mtime,
            createdAt: stats.birthtime
          });
        }
      }

      return excelFiles.sort((a, b) => b.modifiedAt.getTime() - a.modifiedAt.getTime());
    } catch (error) {
      this.logger.error(`Error fetching Excel files for ${opportunityId}:`, error);
      return [];
    }
  }

  /**
   * Get PDF files for an opportunity
   */
  private async getPdfFiles(opportunityId: string) {
    try {
      const pdfFiles: any[] = [];

      // Search in regular opportunities folder/pdfs
      const pdfFolder1 = path.join(this.OPPORTUNITIES_FOLDER, 'pdfs');
      if (fs.existsSync(pdfFolder1)) {
        const files = fs.readdirSync(pdfFolder1);
        const matchingFiles = files.filter(file => 
          file.includes(opportunityId) && file.endsWith('.pdf')
        );

        for (const file of matchingFiles) {
          const filePath = path.join(pdfFolder1, file);
          const stats = fs.statSync(filePath);
          pdfFiles.push({
            fileName: file,
            filePath: filePath,
            calculatorType: 'off-peak',
            size: stats.size,
            modifiedAt: stats.mtime,
            createdAt: stats.birthtime
          });
        }
      }

      // Search in EPVS opportunities folder/pdfs
      const pdfFolder2 = path.join(this.EPVS_OPPORTUNITIES_FOLDER, 'pdfs');
      if (fs.existsSync(pdfFolder2)) {
        const files = fs.readdirSync(pdfFolder2);
        const matchingFiles = files.filter(file => 
          file.includes(opportunityId) && file.endsWith('.pdf')
        );

        for (const file of matchingFiles) {
          const filePath = path.join(pdfFolder2, file);
          const stats = fs.statSync(filePath);
          pdfFiles.push({
            fileName: file,
            filePath: filePath,
            calculatorType: file.includes('EPVS') ? 'epvs' : 'flux',
            size: stats.size,
            modifiedAt: stats.mtime,
            createdAt: stats.birthtime
          });
        }
      }

      return pdfFiles.sort((a, b) => b.modifiedAt.getTime() - a.modifiedAt.getTime());
    } catch (error) {
      this.logger.error(`Error fetching PDF files for ${opportunityId}:`, error);
      return [];
    }
  }


  /**
   * Get all users with their opportunities (summary only)
   */
  async getAllUsersWithOpportunitiesSummary() {
    try {
      this.logger.log('Fetching all users with opportunities summary...');

      const users = await this.prisma.user.findMany({
        select: {
          id: true,
          name: true,
          email: true,
          username: true,
          role: true,
          status: true,
          ghlUserId: true,
          ghlUserName: true,
          createdAt: true,
          lastLoginAt: true
        },
        orderBy: { createdAt: 'desc' }
      });

      const allOpportunities = await this.prisma.opportunityProgress.findMany({
        select: {
          id: true,
          ghlOpportunityId: true,
          userId: true,
          status: true,
          currentStep: true,
          totalSteps: true,
          contactAddress: true,
          contactPostcode: true,
          createdAt: true,
          updatedAt: true
        },
        orderBy: { createdAt: 'desc' }
      });

      // Get all calculator progress records to extract customer names
      const allCalculatorProgress = await this.prisma.calculatorProgress.findMany({
        select: {
          opportunityId: true,
          calculatorType: true,
          data: true
        }
      });

      // Create a map of opportunityId -> customerName
      const customerNameMap = new Map<string, string>();
      for (const calc of allCalculatorProgress) {
        const data = calc.data as any;
        if (data?.customerDetails?.customerName && !customerNameMap.has(calc.opportunityId)) {
          customerNameMap.set(calc.opportunityId, data.customerDetails.customerName);
        }
      }

      const usersWithOpportunities = users.map(user => {
        const userOpportunities = allOpportunities.filter(
          opp => opp.userId === user.id
        );

        return {
          user: {
            id: user.id,
            name: user.name,
            email: user.email,
            username: user.username,
            role: user.role,
            status: user.status,
            ghlUserId: user.ghlUserId,
            ghlUserName: user.ghlUserName,
            createdAt: user.createdAt,
            lastLoginAt: user.lastLoginAt
          },
          opportunities: userOpportunities.map(opp => ({
            id: opp.id,
            ghlOpportunityId: opp.ghlOpportunityId,
            customerName: customerNameMap.get(opp.ghlOpportunityId) || null,
            status: opp.status,
            currentStep: opp.currentStep,
            totalSteps: opp.totalSteps,
            contactAddress: opp.contactAddress,
            contactPostcode: opp.contactPostcode,
            createdAt: opp.createdAt,
            updatedAt: opp.updatedAt
          })),
          totalOpportunities: userOpportunities.length
        };
      });

      this.logger.log(`Found ${users.length} users with ${allOpportunities.length} total opportunities`);
      return usersWithOpportunities;

    } catch (error) {
      this.logger.error('Error fetching users with opportunities summary:', error);
      throw new Error(`Failed to fetch users with opportunities: ${error.message}`);
    }
  }
}

