import { Controller, Post, Get, Body, Param, UseGuards, Request, Logger } from '@nestjs/common';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { OpenSolarService, OpenSolarProjectData } from './opensolar.service';
import { PrismaService } from '../prisma/prisma.service';
import { ExcelAutomationService } from '../excel-automation/excel-automation.service';
import { EPVSAutomationService } from '../epvs-automation/epvs-automation.service';

interface SearchProjectsDto {
  address: string;
}

interface SaveProjectDto {
  opportunityId: string;
  opensolarProjectId: number;
}

@Controller('opensolar')
@UseGuards(JwtAuthGuard)
export class OpenSolarController {
  private readonly logger = new Logger(OpenSolarController.name);

  constructor(
    private readonly openSolarService: OpenSolarService,
    private readonly prisma: PrismaService,
    private readonly excelAutomationService: ExcelAutomationService,
    private readonly epvsAutomationService: EPVSAutomationService,
  ) {}

  /**
   * Search for OpenSolar projects by address
   */
  @Post('search')
  async searchProjects(@Body() searchDto: SearchProjectsDto) {
    try {
      this.logger.log(`🔍 Searching OpenSolar projects for address: ${searchDto.address}`);
      
      const projects = await this.openSolarService.searchProjectsByAddress(searchDto.address);
      
      return {
        success: true,
        data: projects,
        count: projects.length
      };
    } catch (error) {
      this.logger.error('❌ Error searching OpenSolar projects:', error.message);
      return {
        success: false,
        message: 'Failed to search OpenSolar projects',
        error: error.message
      };
    }
  }

  /**
   * Get project data by OpenSolar project ID
   */
  @Get('project/:projectId')
  async getProjectData(@Param('projectId') projectId: string) {
    try {
      const numericProjectId = parseInt(projectId);
      if (isNaN(numericProjectId)) {
        throw new Error('Invalid project ID');
      }

      this.logger.log(`📋 Fetching OpenSolar project data for ID: ${projectId}`);
      
      const projectData = await this.openSolarService.getProjectData(numericProjectId);
      
      return {
        success: true,
        data: projectData
      };
    } catch (error) {
      this.logger.error(`❌ Error fetching project data for ID ${projectId}:`, error.message);
      return {
        success: false,
        message: 'Failed to fetch project data',
        error: error.message
      };
    }
  }

  /**
   * Save OpenSolar project data to database for an opportunity
   */
  @Post('save-project')
  async saveProject(@Body() saveDto: SaveProjectDto) {
    try {
      this.logger.log(`💾 Saving OpenSolar project ${saveDto.opensolarProjectId} for opportunity ${saveDto.opportunityId}`);
      
      // Fetch project data from OpenSolar
      const projectData = await this.openSolarService.getProjectData(saveDto.opensolarProjectId);
      
      // Save to database
      const savedProject = await this.prisma.openSolarProject.upsert({
        where: { opportunityId: saveDto.opportunityId },
                 update: {
           opensolarProjectId: saveDto.opensolarProjectId.toString(),
           projectName: projectData.projectName,
           address: projectData.address,
           systems: projectData.systems as any,
           updatedAt: new Date()
         },
         create: {
           opportunityId: saveDto.opportunityId,
           opensolarProjectId: saveDto.opensolarProjectId.toString(),
           projectName: projectData.projectName,
           address: projectData.address,
           systems: projectData.systems as any
         }
      });

      this.logger.log(`✅ OpenSolar project saved successfully for opportunity ${saveDto.opportunityId}`);
      
      return {
        success: true,
        data: savedProject,
        message: 'OpenSolar project data saved successfully'
      };
    } catch (error) {
      this.logger.error(`❌ Error saving OpenSolar project:`, error.message);
      return {
        success: false,
        message: 'Failed to save OpenSolar project data',
        error: error.message
      };
    }
  }

  /**
   * Get saved OpenSolar project data for an opportunity
   */
  @Get('opportunity/:opportunityId')
  async getOpportunityProject(@Param('opportunityId') opportunityId: string) {
    try {
      this.logger.log(`📋 Fetching saved OpenSolar project for opportunity: ${opportunityId}`);
      
      const savedProject = await this.prisma.openSolarProject.findUnique({
        where: { opportunityId }
      });

      if (!savedProject) {
        return {
          success: false,
          message: 'No OpenSolar project found for this opportunity',
          data: null
        };
      }

      return {
        success: true,
        data: savedProject
      };
    } catch (error) {
      this.logger.error(`❌ Error fetching saved project for opportunity ${opportunityId}:`, error.message);
      return {
        success: false,
        message: 'Failed to fetch saved project data',
        error: error.message
      };
    }
  }

  /**
   * Get calculator-ready data from OpenSolar project
   */
  @Get('calculator-data/:opportunityId')
  async getCalculatorData(@Param('opportunityId') opportunityId: string) {
    try {
      this.logger.log(`🧮 Getting calculator data from OpenSolar for opportunity: ${opportunityId}`);
      
      const savedProject = await this.prisma.openSolarProject.findUnique({
        where: { opportunityId }
      });

      if (!savedProject) {
        return {
          success: false,
          message: 'No OpenSolar project found for this opportunity',
          data: null
        };
      }

      // Check if the saved project has empty systems - if so, try to refresh from OpenSolar API
      let currentProject = savedProject;
      const initialSystems = currentProject.systems as any[];
      if (initialSystems.length === 0) {
        this.logger.log(`🔄 OpenSolar project has empty systems, refreshing from API...`);
        try {
          const freshProjectData = await this.openSolarService.getProjectData(parseInt(currentProject.opensolarProjectId));
          
          // Update the database with fresh data
          currentProject = await this.prisma.openSolarProject.update({
            where: { opportunityId },
            data: {
              systems: freshProjectData.systems as any,
              updatedAt: new Date()
            }
          });
          
          this.logger.log(`✅ Refreshed OpenSolar project data from API`);
          
          // Use the updated project data
          const updatedSystems = currentProject.systems as any[];
          this.logger.log(`🔍 Refreshed systems data:`, {
            systemsCount: updatedSystems.length,
            systems: updatedSystems
          });
        } catch (refreshError) {
          this.logger.error(`❌ Failed to refresh OpenSolar project data:`, refreshError.message);
          // Continue with the original saved project data
        }
      }

             // Transform OpenSolar data to calculator format
       const systems = currentProject.systems as any[];
       const calculatorData: any = {};

       this.logger.log(`🔍 OpenSolar project data structure:`, {
         projectId: currentProject.opensolarProjectId,
         projectName: currentProject.projectName,
         systemsCount: systems.length,
         systems: systems
       });

       if (systems.length > 0) {
         const firstSystem = systems[0];
         this.logger.log(`🔍 First system structure:`, {
           panels: firstSystem.panels || [],
           arrays: firstSystem.arrays || [],
           batteries: firstSystem.batteries || [],
           inverters: firstSystem.inverters || [],
           orientation: firstSystem.orientation || {},
           shading: firstSystem.shading || {}
         });

         const panels = firstSystem.panels || [];
         const arrays = firstSystem.arrays || [];
         const batteries = firstSystem.batteries || [];
         const inverters = firstSystem.inverters || [];
         const orientation = firstSystem.orientation || {};
         const shading = firstSystem.shading || {};

         // Extract panel information
         if (panels.length > 0) {
           const firstPanel = panels[0];
           calculatorData.panel_manufacturer = this.extractManufacturer(firstPanel.model);
           calculatorData.panel_model = firstPanel.model;
           calculatorData.panel_quantity = firstPanel.count;
           calculatorData.panel_wattage = firstPanel.watt_per_module;
           
           // Calculate total system size
           const totalDcKw = panels.reduce((sum: number, panel: any) => 
             sum + (panel.dc_size_kw || 0), 0);
           calculatorData.system_size_kw = Math.round(totalDcKw * 1000) / 1000;
         }

         // Extract array information
         if (arrays.length > 0) {
           this.logger.log(`🔍 Found ${arrays.length} arrays from OpenSolar`);
           calculatorData.arrays = arrays.map((array: any, index: number) => {
             this.logger.log(`🔍 Array ${index + 1}:`, {
               name: array.name,
               panelCount: array.panelCount,
               orientation: array.orientation,
               shading: array.shading
             });
             return {
               name: array.name,
               panelCount: array.panelCount,
               panelModel: array.panelModel,
               // Use system-level orientation and shading for each array if array-specific data is not available
               orientation: array.orientation || orientation,
               shading: array.shading || shading
             };
           });
         } else {
           this.logger.log(`🔍 No arrays found in OpenSolar data - checking if we can create arrays from panels`);
           
           // If no arrays but we have panels, try to create array data from panel information
           if (panels.length > 0) {
             this.logger.log(`🔍 Creating array data from ${panels.length} panels`);
             calculatorData.arrays = panels.map((panel: any, index: number) => {
               this.logger.log(`🔍 Creating array ${index + 1} from panel:`, {
                 model: panel.model,
                 count: panel.count,
                 wattage: panel.watt_per_module
               });
               return {
                 name: `Array ${index + 1}`,
                 panelCount: panel.count,
                 panelModel: panel.model,
                 orientation: orientation,
                 shading: shading
               };
             });
           }
         }

         // Extract battery information
         if (batteries.length > 0) {
           const firstBattery = batteries[0];
           calculatorData.battery_manufacturer = firstBattery.manufacturer;
           calculatorData.battery_model = firstBattery.model;
           calculatorData.battery_capacity = firstBattery.capacity;
           calculatorData.battery_voltage = firstBattery.voltage;
         }

         // Extract inverter information
         if (inverters.length > 0) {
           const solarInverters = inverters.filter((inv: any) => inv.type === 'solar');
           const batteryInverters = inverters.filter((inv: any) => inv.type === 'battery');
           
           if (solarInverters.length > 0) {
             const firstSolarInverter = solarInverters[0];
             calculatorData.solar_inverter_manufacturer = firstSolarInverter.manufacturer;
             calculatorData.solar_inverter_model = firstSolarInverter.model;
             calculatorData.solar_inverter_capacity = firstSolarInverter.capacity;
           }
           
           if (batteryInverters.length > 0) {
             const firstBatteryInverter = batteryInverters[0];
             calculatorData.battery_inverter_manufacturer = firstBatteryInverter.manufacturer;
             calculatorData.battery_inverter_model = firstBatteryInverter.model;
             calculatorData.battery_inverter_capacity = firstBatteryInverter.capacity;
           }
         }

         // Extract orientation information
         if (orientation.tilt || orientation.azimuth) {
           calculatorData.orientation = {
             tilt: orientation.tilt,
             azimuth: orientation.azimuth,
             faces: orientation.faces || []
           };
         }

         // Extract shading information
         if (shading.annualLoss || shading.monthlyLoss) {
           calculatorData.shading = {
             annualLoss: shading.annualLoss,
             monthlyLoss: shading.monthlyLoss
           };
         }
       } else {
         this.logger.log(`⚠️ No systems found in OpenSolar project data`);
       }

       // If no systems/panels found, return empty data with a note
       if (Object.keys(calculatorData).length === 0) {
         calculatorData.note = 'Project found but no design data available - please complete the design in OpenSolar first';
         this.logger.log(`⚠️ No calculator data extracted - OpenSolar project exists but has no design data`);
       } else {
         this.logger.log(`✅ Calculator data extracted successfully:`, {
           hasArrays: !!calculatorData.arrays,
           arrayCount: calculatorData.arrays?.length || 0,
           hasPanels: !!calculatorData.panel_model,
           hasBatteries: !!calculatorData.battery_model,
           hasInverters: !!calculatorData.solar_inverter_model
         });
       }

      return {
        success: true,
        data: calculatorData,
        message: 'Calculator data extracted from OpenSolar project'
      };
    } catch (error) {
      this.logger.error(`❌ Error getting calculator data for opportunity ${opportunityId}:`, error.message);
      return {
        success: false,
        message: 'Failed to get calculator data',
        error: error.message
      };
    }
  }

  /**
   * Automatically populate Excel sheet with OpenSolar data
   */
  @Post('auto-populate-excel/:opportunityId')
  async autoPopulateExcel(@Param('opportunityId') opportunityId: string, @Body() body: { templateFileName?: string }) {
    try {
      this.logger.log(`🤖 Auto-populating Excel with OpenSolar data for opportunity: ${opportunityId}`);
      
      const savedProject = await this.prisma.openSolarProject.findUnique({
        where: { opportunityId }
      });

      if (!savedProject) {
        return {
          success: false,
          message: 'No OpenSolar project found for this opportunity',
          data: null
        };
      }

      // Get calculator data from OpenSolar
      const calculatorData = await this.getCalculatorData(opportunityId);
      
      if (!calculatorData.success) {
        return {
          success: false,
          message: 'Failed to extract calculator data from OpenSolar project',
          data: null
        };
      }

      // Auto-populate Excel with OpenSolar data
      const autoPopulateResult = await this.excelAutomationService.autoPopulateWithOpenSolarData(
        opportunityId,
        calculatorData.data,
        body.templateFileName
      );

      if (autoPopulateResult.success) {
        this.logger.log(`✅ Successfully auto-populated Excel with OpenSolar data`);
        return {
          success: true,
          message: 'Excel sheet auto-populated with OpenSolar data',
          data: autoPopulateResult.data
        };
      } else {
        return {
          success: false,
          message: autoPopulateResult.message,
          error: autoPopulateResult.error
        };
      }

    } catch (error) {
      this.logger.error(`❌ Error auto-populating Excel with OpenSolar data:`, error.message);
      return {
        success: false,
        message: 'Failed to auto-populate Excel with OpenSolar data',
        error: error.message
      };
    }
  }

  /**
   * Automatically populate EPVS Excel sheet with OpenSolar data
   */
  @Post('auto-populate-epvs/:opportunityId')
  async autoPopulateEPVS(@Param('opportunityId') opportunityId: string, @Body() body: { templateFileName?: string }) {
    try {
      this.logger.log(`🤖 Auto-populating EPVS Excel with OpenSolar data for opportunity: ${opportunityId}`);
      
      const savedProject = await this.prisma.openSolarProject.findUnique({
        where: { opportunityId }
      });

      if (!savedProject) {
        return {
          success: false,
          message: 'No OpenSolar project found for this opportunity',
          data: null
        };
      }

      // Get calculator data from OpenSolar
      const calculatorData = await this.getCalculatorData(opportunityId);
      
      if (!calculatorData.success) {
        return {
          success: false,
          message: 'Failed to extract calculator data from OpenSolar project',
          data: null
        };
      }

      // Auto-populate EPVS Excel with OpenSolar data
      const autoPopulateResult = await this.epvsAutomationService.autoPopulateWithOpenSolarData(
        opportunityId,
        calculatorData.data,
        body.templateFileName
      );

      if (autoPopulateResult.success) {
        this.logger.log(`✅ Successfully auto-populated EPVS Excel with OpenSolar data`);
        return {
          success: true,
          message: 'EPVS Excel sheet auto-populated with OpenSolar data',
          data: autoPopulateResult.data
        };
      } else {
        return {
          success: false,
          message: autoPopulateResult.message,
          error: autoPopulateResult.error
        };
      }

    } catch (error) {
      this.logger.error(`❌ Error auto-populating EPVS Excel with OpenSolar data:`, error.message);
      return {
        success: false,
        message: 'Failed to auto-populate EPVS Excel with OpenSolar data',
        error: error.message
      };
    }
  }

  /**
   * Extract manufacturer name from panel model
   */
  private extractManufacturer(modelName: string): string {
    if (!modelName) return '';
    
    // Common manufacturer patterns
    const manufacturers = [
      'Jinko', 'Trina', 'Canadian Solar', 'Longi', 'JA Solar', 'Risen', 'Q Cells',
      'SunPower', 'LG', 'Panasonic', 'REC', 'SolarWorld', 'First Solar', 'Yingli'
    ];

    for (const manufacturer of manufacturers) {
      if (modelName.toLowerCase().includes(manufacturer.toLowerCase())) {
        return manufacturer;
      }
    }

    // Try to extract from common patterns
    const words = modelName.split(/[\s\-_]+/);
    if (words.length > 0) {
      return words[0];
    }

    return 'Unknown';
  }
}
