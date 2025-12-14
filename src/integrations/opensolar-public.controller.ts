import { Controller, Post, Get, Body, Param, Logger, HttpException, HttpStatus } from '@nestjs/common';
import { OpenSolarPublicService, CreateProjectDto, CreateDesignDto } from './opensolar-public.service';

@Controller('opensolar-public')
export class OpenSolarPublicController {
  private readonly logger = new Logger(OpenSolarPublicController.name);

  constructor(
    private readonly openSolarPublicService: OpenSolarPublicService,
  ) {}

  /**
   * Create a new OpenSolar project (no authentication required)
   */
  @Post('create-project')
  async createProject(@Body() createProjectDto: CreateProjectDto) {
    try {
      this.logger.log(`🏗️ Creating new OpenSolar project: ${createProjectDto.name}`);
      
      const project = await this.openSolarPublicService.createProject(createProjectDto);
      
      return {
        success: true,
        data: {
          project,
          projectUrl: this.openSolarPublicService.getProjectUrl(project.id),
          message: 'OpenSolar project created successfully'
        }
      };
    } catch (error: any) {
      this.logger.error('❌ Error creating OpenSolar project:', error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to create OpenSolar project',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a new design for an OpenSolar project (no authentication required)
   */
  @Post('create-design')
  async createDesign(@Body() createDesignDto: CreateDesignDto) {
    try {
      this.logger.log(`🎨 Creating new design for project ${createDesignDto.projectId}: ${createDesignDto.name}`);
      
      const design = await this.openSolarPublicService.createDesign(createDesignDto);
      
      return {
        success: true,
        data: {
          design,
          designUrl: this.openSolarPublicService.getDesignUrl(createDesignDto.projectId, design.id),
          projectUrl: this.openSolarPublicService.getProjectUrl(createDesignDto.projectId),
          message: 'OpenSolar design created successfully'
        }
      };
    } catch (error: any) {
      this.logger.error('❌ Error creating OpenSolar design:', error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to create OpenSolar design',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Get project details by ID (no authentication required)
   */
  @Get('project/:projectId')
  async getProject(@Param('projectId') projectId: string) {
    try {
      const numericProjectId = parseInt(projectId);
      if (isNaN(numericProjectId)) {
        throw new HttpException('Invalid project ID', HttpStatus.BAD_REQUEST);
      }

      this.logger.log(`📋 Fetching OpenSolar project #${projectId}`);
      
      const project = await this.openSolarPublicService.getProject(numericProjectId);
      
      return {
        success: true,
        data: {
          project,
          projectUrl: this.openSolarPublicService.getProjectUrl(project.id)
        }
      };
    } catch (error: any) {
      this.logger.error(`❌ Error fetching project #${projectId}:`, error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to fetch project data',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Get design details by project ID and design ID (no authentication required)
   */
  @Get('project/:projectId/design/:designId')
  async getDesign(
    @Param('projectId') projectId: string,
    @Param('designId') designId: string
  ) {
    try {
      const numericProjectId = parseInt(projectId);
      if (isNaN(numericProjectId)) {
        throw new HttpException('Invalid project ID', HttpStatus.BAD_REQUEST);
      }

      this.logger.log(`🎨 Fetching design #${designId} for project #${projectId}`);
      
      const design = await this.openSolarPublicService.getDesign(numericProjectId, designId);
      
      return {
        success: true,
        data: {
          design,
          designUrl: this.openSolarPublicService.getDesignUrl(numericProjectId, design.id),
          projectUrl: this.openSolarPublicService.getProjectUrl(numericProjectId)
        }
      };
    } catch (error: any) {
      this.logger.error(`❌ Error fetching design #${designId}:`, error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to fetch design data',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * List all designs for a project (no authentication required)
   */
  @Get('project/:projectId/designs')
  async listDesigns(@Param('projectId') projectId: string) {
    try {
      const numericProjectId = parseInt(projectId);
      if (isNaN(numericProjectId)) {
        throw new HttpException('Invalid project ID', HttpStatus.BAD_REQUEST);
      }

      this.logger.log(`📋 Fetching designs for project #${projectId}`);
      
      const designs = await this.openSolarPublicService.listDesigns(numericProjectId);
      
      return {
        success: true,
        data: {
          designs,
          projectUrl: this.openSolarPublicService.getProjectUrl(numericProjectId),
          count: designs.length
        }
      };
    } catch (error: any) {
      this.logger.error(`❌ Error fetching designs for project #${projectId}:`, error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to fetch designs',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create a complete project with design in one call (no authentication required)
   */
  @Post('create-project-with-design')
  async createProjectWithDesign(@Body() body: {
    project: CreateProjectDto;
    design: Omit<CreateDesignDto, 'projectId'>;
  }) {
    try {
      this.logger.log(`🚀 Creating complete OpenSolar project with design: ${body.project.name}`);
      
      // Step 1: Create the project
      const project = await this.openSolarPublicService.createProject(body.project);
      
      // Step 2: Create the design for the project
      const design = await this.openSolarPublicService.createDesign({
        ...body.design,
        projectId: project.id
      });
      
      return {
        success: true,
        data: {
          project,
          design,
          projectUrl: this.openSolarPublicService.getProjectUrl(project.id),
          designUrl: this.openSolarPublicService.getDesignUrl(project.id, design.id),
          message: 'OpenSolar project and design created successfully'
        }
      };
    } catch (error: any) {
      this.logger.error('❌ Error creating project with design:', error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to create project with design',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Get authentication token for frontend use
   */
  @Get('auth-token')
  async getAuthToken() {
    try {
      this.logger.log('🔐 Getting OpenSolar authentication token for frontend');
      
      const authData = await this.openSolarPublicService.getAuthToken();
      
      return {
        success: true,
        data: {
          token: authData.token,
          orgId: authData.orgId,
          message: 'Authentication token retrieved successfully'
        }
      };
    } catch (error: any) {
      this.logger.error('❌ Error getting auth token:', error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to get authentication token',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Create project and get authenticated URL
   */
  @Post('create-project-authenticated')
  async createProjectAuthenticated(@Body() createProjectDto: CreateProjectDto) {
    try {
      this.logger.log(`🏗️ Creating authenticated OpenSolar project: ${createProjectDto.name}`);
      
      // Step 1: Create the project
      const project = await this.openSolarPublicService.createProject(createProjectDto);
      
      // Step 2: Get authentication token
      const authData = await this.openSolarPublicService.getAuthToken();
      
      // Step 3: Generate authenticated URL
      const authenticatedUrl = this.openSolarPublicService.getAuthenticatedProjectUrl(project.id, authData.token);
      
      return {
        success: true,
        data: {
          project,
          token: authData.token,
          orgId: authData.orgId,
          authenticatedUrl,
          projectUrl: this.openSolarPublicService.getProjectUrl(project.id),
          message: 'OpenSolar project created with authentication successfully'
        }
      };
    } catch (error: any) {
      this.logger.error('❌ Error creating authenticated project:', error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to create authenticated project',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }

  /**
   * Get quick start template for common solar configurations
   */
  @Get('templates')
  async getTemplates() {
    try {
      const templates = {
        residential: {
          name: 'Residential Solar System',
          description: 'Standard residential solar installation',
          systemType: 'solar' as const,
          panels: [
            {
              model: 'Jinko JKM400M-72HL4-V',
              count: 20,
              watt_per_module: 400,
              manufacturer: 'Jinko'
            }
          ],
          arrays: [
            {
              name: 'Main Array',
              panel_count: 20,
              panel_model: 'Jinko JKM400M-72HL4-V',
              orientation: {
                tilt: 30,
                azimuth: 180,
                face: 'south'
              }
            }
          ],
          inverters: [
            {
              manufacturer: 'SMA',
              model: 'Sunny Boy 8.0-US',
              type: 'solar' as const,
              capacity: 8.0
            }
          ]
        },
        commercial: {
          name: 'Commercial Solar System',
          description: 'Medium-scale commercial installation',
          systemType: 'solar' as const,
          panels: [
            {
              model: 'Trina Solar TSM-400DE15',
              count: 50,
              watt_per_module: 400,
              manufacturer: 'Trina Solar'
            }
          ],
          arrays: [
            {
              name: 'Commercial Array',
              panel_count: 50,
              panel_model: 'Trina Solar TSM-400DE15',
              orientation: {
                tilt: 25,
                azimuth: 180,
                face: 'south'
              }
            }
          ],
          inverters: [
            {
              manufacturer: 'Fronius',
              model: 'Symo 20.0-3-M',
              type: 'solar' as const,
              capacity: 20.0
            }
          ]
        },
        hybrid: {
          name: 'Hybrid Solar + Battery System',
          description: 'Solar system with battery backup',
          systemType: 'hybrid' as const,
          panels: [
            {
              model: 'LG NeON 2 LG400N2K-V5',
              count: 16,
              watt_per_module: 400,
              manufacturer: 'LG'
            }
          ],
          arrays: [
            {
              name: 'Hybrid Array',
              panel_count: 16,
              panel_model: 'LG NeON 2 LG400N2K-V5',
              orientation: {
                tilt: 35,
                azimuth: 180,
                face: 'south'
              }
            }
          ],
          batteries: [
            {
              manufacturer: 'Tesla',
              model: 'Powerwall 2',
              capacity: 13.5,
              voltage: 400
            }
          ],
          inverters: [
            {
              manufacturer: 'Tesla',
              model: 'Powerwall 2 Gateway',
              type: 'hybrid' as const,
              capacity: 5.0
            }
          ]
        }
      };

      return {
        success: true,
        data: templates,
        message: 'Solar system templates retrieved successfully'
      };
    } catch (error: any) {
      this.logger.error('❌ Error fetching templates:', error.message);
      throw new HttpException(
        {
          success: false,
          message: 'Failed to fetch templates',
          error: error.message
        },
        HttpStatus.INTERNAL_SERVER_ERROR
      );
    }
  }
}

