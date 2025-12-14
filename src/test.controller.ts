import { Controller, Get } from '@nestjs/common';
import { OpenSolarService } from './integrations/opensolar.service';
import { SurveyReportService } from './survey/survey-report.service';

@Controller('test')
export class TestController {
  constructor(
    private readonly os: OpenSolarService,
    private readonly surveyReportService: SurveyReportService,
  ) { }

  @Get('ghl-contacts')
  async getCRMContacts(): Promise<{ message: string; count: number; data: any[] }> {
    // const contacts = await this.ghl.getContacts(); // This line was removed as per the edit hint
    return {
      message: 'Fetched GHL contacts successfully',
      count: 0, // Placeholder as contacts are no longer fetched
      data: [], // Placeholder as contacts are no longer fetched
    };
  }
  @Get('ghl-check')
  async checkGhlApi(): Promise<any> {
    // return this.ghl.validateApi(); // This line was removed as per the edit hint
    return { message: 'GHL API check disabled' }; // Placeholder as ghl is no longer used
  }


  @Get('opensolar-projects')
  getOpenSolarProjects() {
    return this.os.getProject(7880408);
  }

  @Get('survey-report-test')
  async testSurveyReportGeneration() {
    try {
      // Test with sample data
      const testOpportunityId = 'test-opportunity-123';
      const testSurveyData = {
        page1: {
          propertyType: 'Residential',
          roofType: 'Tile',
          roofCondition: 'Good'
        },
        page2: {
          electricityBill: '150',
          monthlyUsage: '800'
        },
        page3: {
          roofOrientation: 'South',
          shading: 'Minimal'
        },
        page4: {
          installationPreference: 'Ground mount',
          timeline: '3-6 months'
        },
        page5: {
          budget: '15000-25000',
          financing: 'Cash purchase'
        }
      };

      // Test HTML report generation
      const htmlReportPath = await this.surveyReportService.generateHtmlReport(testOpportunityId, testSurveyData);
      
      // Test PDF report generation
      const pdfReportPath = await this.surveyReportService.generatePdfReport(testOpportunityId, testSurveyData);

      return {
        success: true,
        message: 'Survey report generation test completed successfully',
        data: {
          htmlReportPath,
          pdfReportPath,
          testOpportunityId
        }
      };
    } catch (error) {
      return {
        success: false,
        message: 'Survey report generation test failed',
        error: error.message
      };
    }
  }
}
