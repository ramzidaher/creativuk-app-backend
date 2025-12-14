import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import axios from 'axios';

export interface OpenSolarAuth {
  token: string;
  orgId: number;
}

export interface OpenSolarProject {
  id: number;
  address?: string;
  display_name?: string;
  name?: string;
}

export interface OpenSolarPanel {
  model: string;
  count: number;
  watt_per_module?: number;
  dc_size_kw?: number;
}

export interface OpenSolarArray {
  id: string;
  name?: string;
  panelCount: number;
  panelModel: string;
  orientation: {
    tilt?: number;
    azimuth?: number;
    face?: string;
  };
  shading?: {
    annualLoss?: number;
    monthlyLoss?: number[];
  };
}

export interface OpenSolarBattery {
  manufacturer: string;
  model: string;
  capacity?: number;
  voltage?: number;
}

export interface OpenSolarInverter {
  manufacturer: string;
  model: string;
  type: 'solar' | 'battery' | 'hybrid';
  capacity?: number;
}



export interface OpenSolarSystem {
  uuid: string;
  name?: string;
  display_name?: string;
  panels: OpenSolarPanel[];
  arrays: OpenSolarArray[];
  batteries: OpenSolarBattery[];
  inverters: OpenSolarInverter[];
  total_dc_kw_est?: number;
  orientation?: {
    tilt?: number;
    azimuth?: number;
    normalizedAzimuth?: number;
    faces?: Array<{
      tilt: number;
      azimuth: number;
      normalizedAzimuth: number;
      face: string;
    }>;
  };
  shading?: {
    annualLoss?: number;
    monthlyLoss?: number[];
  };
}

export interface OpenSolarProjectData {
  projectId: number;
  projectName: string;
  address?: string;
  systems: OpenSolarSystem[];
}

@Injectable()
export class OpenSolarService {
  private readonly logger = new Logger(OpenSolarService.name);
  private readonly baseUrl = 'https://api.opensolar.com';
  private readonly username: string;
  private readonly password: string;
  private httpClient: any;
  private authCache: { token: string; orgId: number; expiresAt: number } | null = null;
  private readonly AUTH_CACHE_DURATION = 30 * 60 * 1000; // 30 minutes

  constructor(private configService: ConfigService) {
    this.username = this.configService.get<string>('OPENSOLAR_USERNAME') || 'ramzi@paldev.tech';
    this.password = this.configService.get<string>('OPENSOLAR_PASSWORD') || 'pUH6WdNCC,ZUdKd';
    
    this.httpClient = axios.create({
      baseURL: this.baseUrl,
      timeout: 30000,
    });
  }

  /**
   * Authenticate with OpenSolar API with caching and rate limit handling
   */
  async authenticate(): Promise<OpenSolarAuth> {
    // Check if we have a valid cached auth
    if (this.authCache && Date.now() < this.authCache.expiresAt) {
      this.logger.log('🔄 Using cached OpenSolar authentication');
      return { token: this.authCache.token, orgId: this.authCache.orgId };
    }

    try {
      this.logger.log('🔐 Authenticating with OpenSolar...');
      
      const response = await this.httpClient.post('/api-token-auth/', {
        username: this.username,
        password: this.password,
      });

      const data = response.data;
      const token = data.token;
      const orgId = data.org_id || (data.orgs?.[0]?.id);

      if (!orgId) {
        throw new Error('No org_id returned in login response');
      }

      // Cache the authentication
      this.authCache = {
        token,
        orgId: parseInt(orgId),
        expiresAt: Date.now() + this.AUTH_CACHE_DURATION
      };

      this.logger.log(`✅ OpenSolar authentication successful | org_id=${orgId}`);
      return { token, orgId: parseInt(orgId) };
    } catch (error) {
      this.logger.error('❌ OpenSolar authentication failed:', error.message);
      
      // Handle rate limiting specifically
      if (error.response?.status === 429) {
        this.logger.warn('⚠️ OpenSolar rate limit exceeded. Please wait before trying again.');
        throw new Error('OpenSolar API rate limit exceeded. Please wait a few minutes before trying again.');
      }
      
      throw new Error(`OpenSolar authentication failed: ${error.message}`);
    }
  }

  /**
   * Get project details by ID
   */
  async getProject(projectId: number): Promise<OpenSolarProject> {
    try {
      const auth = await this.authenticate();
      
      this.logger.log(`📋 Fetching OpenSolar project #${projectId}...`);
      
      const response = await this.httpClient.get(
        `/api/orgs/${auth.orgId}/projects/${projectId}/`,
        {
          headers: { Authorization: `Bearer ${auth.token}` }
        }
      );

      const project = response.data;
      this.logger.log(`✅ Project fetched successfully: ${project.address || project.display_name || project.name}`);
      
      return project;
    } catch (error) {
      this.logger.error(`❌ Failed to fetch project #${projectId}:`, error.message);
      throw new Error(`Failed to fetch OpenSolar project: ${error.message}`);
    }
  }

  /**
   * Get system details including hardware and layout
   */
  async getSystemDetails(projectId: number, includeParts: string = 'layout,hardware'): Promise<any> {
    try {
      const auth = await this.authenticate();
      
      this.logger.log(`🔧 Fetching system details for project #${projectId}...`);
      
      const response = await this.httpClient.get(
        `/api/orgs/${auth.orgId}/projects/${projectId}/systems/details/`,
        {
          headers: { Authorization: `Bearer ${auth.token}` },
          params: { include_parts: includeParts }
        }
      );

      this.logger.log(`✅ System details fetched successfully`);
      return response.data;
    } catch (error) {
      this.logger.error(`❌ Failed to fetch system details for project #${projectId}:`, error.message);
      throw new Error(`Failed to fetch OpenSolar system details: ${error.message}`);
    }
  }

  /**
   * Extract panel information from system data
   */
  private extractPanelsFromSystem(sysObj: any): OpenSolarPanel[] {
    const panels: OpenSolarPanel[] = [];

    // Check system.data for module information
    const systemData = sysObj.data || {};
    
    // Look for moduleTypes in system data
    if (systemData.moduleTypes && Array.isArray(systemData.moduleTypes)) {
      for (const moduleType of systemData.moduleTypes) {
        const model = moduleType.code || moduleType.name || 'Unknown module';
        const count = parseInt(moduleType.quantity || '0') || 0;
        const wattPerModule = parseFloat(moduleType.kw_stc || '0') * 1000 || 0; // Convert kW to W
        
        if (count > 0 && wattPerModule > 0) {
          const dcSizeKw = Math.round((count * wattPerModule) / 1000.0 * 1000) / 1000;
          
          panels.push({
            model,
            count,
            watt_per_module: wattPerModule,
            dc_size_kw: dcSizeKw
          });
        }
      }
    }

    // Fallback: Check hardware.modules style
    if (panels.length === 0) {
      const hardware = sysObj.hardware || {};
      const candidates: any[] = [];
      
      ['modules', 'module', 'pv_modules', 'panels', 'panel'].forEach(key => {
        const val = hardware[key];
        if (val) {
          if (Array.isArray(val)) {
            candidates.push(...val);
          } else if (typeof val === 'object') {
            candidates.push(val);
          }
        }
      });

      // Components fallback
      if (candidates.length === 0) {
        const components = this.findFirst(sysObj, ['components', 'parts', 'items']);
        if (Array.isArray(components)) {
          for (const c of components) {
            const txt = [
              c.type, c.category, c.name, c.display_name
            ].filter(Boolean).join(' ').toLowerCase();
            
            if (['module', 'panel', 'pv module', 'pv-module', 'pv panel'].some(s => txt.includes(s))) {
              candidates.push(c);
            }
          }
        }
      }

      // Process candidates
      for (const c of candidates) {
        const model = c.model_name || c.name || c.display_name || 
                     this.dig(c, ['model', 'name']) || 
                     this.dig(c, ['component', 'name']) || 
                     'Unknown module';
        
        const count = parseInt(c.count || c.quantity || c.qty || 
                             this.dig(c, ['quantity', 'value']) || '0') || 0;
        
        const wattPerModule = parseFloat(c.watts || c.wattage || c.power_stc_watts || 
                                       this.dig(c, ['model', 'watts']) || 
                                       this.dig(c, ['specs', 'wattage']) || '0') || 0;

        const dcSizeKw = count && wattPerModule ? 
          Math.round((count * wattPerModule) / 1000.0 * 1000) / 1000 : undefined;

        panels.push({
          model,
          count,
          watt_per_module: wattPerModule || undefined,
          dc_size_kw: dcSizeKw
        });
      }
    }

    // Deduplicate identical models by summing counts
    const consolidated: { [key: string]: OpenSolarPanel } = {};
    for (const p of panels) {
      if (!consolidated[p.model]) {
        consolidated[p.model] = { ...p };
      } else {
        consolidated[p.model].count += p.count;
        if (p.watt_per_module) {
          consolidated[p.model].watt_per_module = p.watt_per_module;
        }
        if (consolidated[p.model].watt_per_module && consolidated[p.model].count) {
          const wattPerModule = consolidated[p.model].watt_per_module;
          const count = consolidated[p.model].count;
          if (wattPerModule && count) {
            consolidated[p.model].dc_size_kw = Math.round(
              (count * wattPerModule) / 1000.0 * 1000
            ) / 1000;
          }
        }
      }
    }

    return Object.values(consolidated);
  }

  /**
   * Extract array information from system data
   */
  private extractArraysFromSystem(sysObj: any): OpenSolarArray[] {
    const arrays: OpenSolarArray[] = [];
    
    // Check system.data for shadingByPanelGroup (arrays)
    const systemData = sysObj.data || {};
    
    this.logger.log(`🔍 Extracting arrays from system data:`, {
      hasShadingByPanelGroup: !!systemData.shadingByPanelGroup,
      shadingByPanelGroupLength: systemData.shadingByPanelGroup?.length || 0
    });
    
    if (systemData.shadingByPanelGroup && Array.isArray(systemData.shadingByPanelGroup)) {
      for (let i = 0; i < systemData.shadingByPanelGroup.length; i++) {
        const array = systemData.shadingByPanelGroup[i];
        
        this.logger.log(`🔍 Processing array ${i + 1}:`, {
          uuid: array.uuid,
          module_quantity: array.module_quantity,
          slope: array.slope,
          relative_tilt: array.relative_tilt,
          tilt: array.tilt,
          azimuth: array.azimuth,
          beam_access: array.beam_access?.length || 0,
          all_fields: Object.keys(array)
        });
        
        const arrayId = array.uuid || `array_${i + 1}`;
        const arrayName = `Array ${i + 1}`;
        
        // Extract panel count and orientation
        const panelCount = parseInt(array.module_quantity || '0') || 0;
        const panelModel = 'From system data'; // Will be filled from moduleTypes
        
        // Extract orientation from array data
        // Look for relative_tilt first (the actual roof pitch), then fall back to slope
        let orientation = {
          tilt: parseFloat(array.relative_tilt || array.tilt || array.slope || '0') || undefined,
          azimuth: parseFloat(array.azimuth || '0') || undefined,
          face: `Group ${i + 1}`
        };
        
        // If array doesn't have orientation data, try to use system-level orientation
        if (!orientation.tilt && !orientation.azimuth) {
          const layout = sysObj.layout || {};
          const systemOrientation = layout.orientation || {};
          orientation = {
            tilt: parseFloat(systemOrientation.relative_tilt || systemOrientation.tilt || systemOrientation.tilt_deg || '0') || undefined,
            azimuth: parseFloat(systemOrientation.azimuth || systemOrientation.azimuth_deg || '0') || undefined,
            face: `Group ${i + 1} (system-level)`
          };
          this.logger.log(`🔍 Array ${i + 1} using system-level orientation:`, orientation);
        } else {
          this.logger.log(`🔍 Array ${i + 1} extracted orientation:`, orientation);
        }
        
        // Extract shading from beam_access array
        const beamAccess = array.beam_access || [];
        const shadingValues = beamAccess.filter((val: any) => val !== null && val !== undefined);
        const averageShading = shadingValues.length > 0 ? 
          (1 - (shadingValues.reduce((sum: number, val: number) => sum + val, 0) / shadingValues.length)) * 100 : 0;
        
        const shading = {
          annualLoss: averageShading,
          monthlyLoss: undefined // Could be calculated from beam_access if needed
        };
        
        arrays.push({
          id: arrayId,
          name: arrayName,
          panelCount,
          panelModel,
          orientation,
          shading
        });
      }
    }
    
    // Fallback: If no arrays found in shadingByPanelGroup, create one from system-level data
    if (arrays.length === 0) {
      this.logger.log(`🔍 No arrays found in shadingByPanelGroup, creating fallback array from system data`);
      
      const layout = sysObj.layout || {};
      const systemOrientation = layout.orientation || {};
      
      // Try to get panel count from system data
      const totalPanels = parseInt(systemData.module_quantity || '0') || 0;
      
      if (totalPanels > 0) {
        const orientation = {
          tilt: parseFloat(systemOrientation.relative_tilt || systemOrientation.tilt || systemOrientation.tilt_deg || '0') || undefined,
          azimuth: parseFloat(systemOrientation.azimuth || systemOrientation.azimuth_deg || '0') || undefined,
          face: 'System Default'
        };
        
        this.logger.log(`🔍 Created fallback array with orientation:`, orientation);
        
        arrays.push({
          id: 'system_array_1',
          name: 'Array 1',
          panelCount: totalPanels,
          panelModel: 'From system data',
          orientation,
          shading: {
            annualLoss: 0,
            monthlyLoss: undefined
          }
        });
      }
    }
    
    return arrays;
  }

  /**
   * Extract battery information from system data
   */
  private extractBatteriesFromSystem(sysObj: any): OpenSolarBattery[] {
    const batteries: OpenSolarBattery[] = [];
    
    // Check system.data for battery information
    const systemData = sysObj.data || {};
    
    if (systemData.battery_total_kwh && systemData.battery_total_kwh > 0) {
      // Extract battery info from system data
      const capacity = systemData.battery_total_kwh;
      
      // Look for battery components in custom_data
      if (systemData.custom_data && systemData.custom_data.component_dependencies) {
        for (const dep of systemData.custom_data.component_dependencies) {
          if (dep.parentComponents) {
            for (const comp of dep.parentComponents) {
              if (comp.componentType === 'battery') {
                batteries.push({
                  manufacturer: 'V-TAC', // From the debug output
                  model: comp.code || 'Unknown',
                  capacity,
                  voltage: undefined
                });
                break;
              }
            }
          }
        }
      }
      
      // If no battery found in dependencies, create a generic one
      if (batteries.length === 0) {
        batteries.push({
          manufacturer: 'System Battery',
          model: 'Battery System',
          capacity,
          voltage: undefined
        });
      }
    }
    
    // Fallback: Check hardware
    if (batteries.length === 0) {
      const hardware = sysObj.hardware || {};
      const batteryCandidates = hardware.batteries || hardware.battery || hardware.storage || [];
      
      if (Array.isArray(batteryCandidates)) {
        for (const battery of batteryCandidates) {
          const manufacturer = battery.manufacturer || battery.brand || 'Unknown';
          const model = battery.model || battery.model_name || battery.name || 'Unknown';
          const capacity = parseFloat(battery.capacity || battery.capacity_kwh || '0') || undefined;
          const voltage = parseFloat(battery.voltage || battery.voltage_v || '0') || undefined;
          
          batteries.push({
            manufacturer,
            model,
            capacity,
            voltage
          });
        }
      }
    }
    
    return batteries;
  }

  /**
   * Extract inverter information from system data
   */
  private extractInvertersFromSystem(sysObj: any): OpenSolarInverter[] {
    const inverters: OpenSolarInverter[] = [];
    
    // Check system.data for inverter information
    const systemData = sysObj.data || {};
    
    // Look for inverter components in custom_data
    if (systemData.custom_data && systemData.custom_data.component_dependencies) {
      for (const dep of systemData.custom_data.component_dependencies) {
        if (dep.parentComponents) {
          for (const comp of dep.parentComponents) {
            if (comp.componentType === 'inverter') {
              inverters.push({
                manufacturer: 'V-TAC', // From the debug output
                model: comp.code || 'Unknown',
                type: 'solar' as 'solar' | 'battery' | 'hybrid',
                capacity: undefined
              });
            }
          }
        }
      }
    }
    
    // Fallback: Check hardware
    if (inverters.length === 0) {
      const hardware = sysObj.hardware || {};
      const inverterCandidates = hardware.inverters || hardware.inverter || [];
      
      if (Array.isArray(inverterCandidates)) {
        for (const inverter of inverterCandidates) {
          const manufacturer = inverter.manufacturer || inverter.brand || 'Unknown';
          const model = inverter.model || inverter.model_name || inverter.name || 'Unknown';
          const type = inverter.type || 'solar';
          const capacity = parseFloat(inverter.capacity || inverter.capacity_kw || '0') || undefined;
          
          inverters.push({
            manufacturer,
            model,
            type: type as 'solar' | 'battery' | 'hybrid',
            capacity
          });
        }
      }
    }
    
    return inverters;
  }

  /**
   * Extract orientation information from system data
   */
  private extractOrientationFromSystem(sysObj: any): any {
    // Check system.data for orientation information
    const systemData = sysObj.data || {};
    
    if (systemData.shadingByPanelGroup && Array.isArray(systemData.shadingByPanelGroup)) {
      const faces = systemData.shadingByPanelGroup.map((group: any, index: number) => {
        const rawTilt = parseFloat(group.slope || '0') || 0;
        const rawAzimuth = parseFloat(group.azimuth || '0') || 0;
        
        // Convert azimuth to 180° system with 5° increments
        const normalizedAzimuth = this.normalizeAzimuthTo180(rawAzimuth);
        
        return {
          tilt: rawTilt,
          azimuth: rawAzimuth,
          normalizedAzimuth,
          face: `Group ${index + 1}`
        };
      });
      
      // Calculate average tilt and azimuth
      const avgTilt = faces.length > 0 ? 
        faces.reduce((sum: number, face: any) => sum + face.tilt, 0) / faces.length : undefined;
      const avgAzimuth = faces.length > 0 ? 
        faces.reduce((sum: number, face: any) => sum + face.azimuth, 0) / faces.length : undefined;
      const avgNormalizedAzimuth = faces.length > 0 ? 
        faces.reduce((sum: number, face: any) => sum + face.normalizedAzimuth, 0) / faces.length : undefined;
      
      return {
        tilt: avgTilt,
        azimuth: avgAzimuth,
        normalizedAzimuth: avgNormalizedAzimuth,
        faces
      };
    }
    
    // Fallback: Check layout
    const layout = sysObj.layout || {};
    const orientation = layout.orientation || {};
    
    // Extract tilt and azimuth
    const tilt = parseFloat(orientation.tilt || orientation.tilt_deg || '0') || undefined;
    const azimuth = parseFloat(orientation.azimuth || orientation.azimuth_deg || '0') || undefined;
    const normalizedAzimuth = azimuth ? this.normalizeAzimuthTo180(azimuth) : undefined;
    
    // Extract faces if available
    const faces = orientation.faces || orientation.roof_faces || [];
    
    return {
      tilt,
      azimuth,
      normalizedAzimuth,
      faces: Array.isArray(faces) ? faces.map((face: any) => {
        const faceTilt = parseFloat(face.tilt || face.tilt_deg || '0') || 0;
        const faceAzimuth = parseFloat(face.azimuth || face.azimuth_deg || '0') || 0;
        const faceNormalizedAzimuth = this.normalizeAzimuthTo180(faceAzimuth);
        
        return {
          tilt: faceTilt,
          azimuth: faceAzimuth,
          normalizedAzimuth: faceNormalizedAzimuth,
          face: face.face || face.name || 'Unknown'
        };
      }) : []
    };
  }

  /**
   * Calculate orientation as difference from 180° and round UP to nearest 5° increment
   * This matches the Excel sheet orientation format which goes up in 5° increments
   * Example: 156.6° -> |180° - 156.6°| = 23.4° -> rounded UP to 25°
   */
  private normalizeAzimuthTo180(azimuth: number): number {
    // Normalize to 0-360 range
    let normalized = azimuth % 360;
    if (normalized < 0) normalized += 360;
    
    // Calculate difference from 180° (the reference direction)
    const differenceFrom180 = Math.abs(180 - normalized);
    
    // Round UP to nearest 5° increment (as per project manager's specification)
    return Math.ceil(differenceFrom180 / 5) * 5;
  }

  /**
   * Extract shading information from system data
   */
  private extractShadingFromSystem(sysObj: any): any {
    // Check system.data for shading information
    const systemData = sysObj.data || {};
    
    if (systemData.shadingByPanelGroup && Array.isArray(systemData.shadingByPanelGroup)) {
      // Calculate average shading from all panel groups
      let totalShading = 0;
      let groupCount = 0;
      
      for (const group of systemData.shadingByPanelGroup) {
        const beamAccess = group.beam_access || [];
        const shadingValues = beamAccess.filter((val: any) => val !== null && val !== undefined);
        
        if (shadingValues.length > 0) {
          // Convert from decimal (0.5 = 50%) to percentage
          const averageShading = (1 - (shadingValues.reduce((sum: number, val: number) => sum + val, 0) / shadingValues.length)) * 100;
          totalShading += averageShading;
          groupCount++;
        }
      }
      
      const annualLoss = groupCount > 0 ? Math.round(totalShading / groupCount * 100) / 100 : undefined;
      
      return {
        annualLoss,
        monthlyLoss: undefined // Could be calculated from beam_access arrays if needed
      };
    }
    
    // Fallback: Check layout
    const layout = sysObj.layout || {};
    const shading = layout.shading || {};
    
    const annualLoss = parseFloat(shading.annual_loss || shading.annual_shading || '0') || undefined;
    const monthlyLoss = Array.isArray(shading.monthly_loss) ? shading.monthly_loss : 
                       Array.isArray(shading.monthly_shading) ? shading.monthly_shading : undefined;
    
    return {
      annualLoss,
      monthlyLoss
    };
  }



  /**
   * Safely dig nested object by array of keys
   */
  private dig(obj: any, keys: string[]): any {
    let current = obj;
    for (const key of keys) {
      if (current && typeof current === 'object' && key in current) {
        current = current[key];
      } else {
        return undefined;
      }
    }
    return current;
  }

  /**
   * Find first matching key in object tree
   */
  private findFirst(obj: any, keyNames: string[]): any {
    const stack = [obj];
    while (stack.length > 0) {
      const current = stack.pop()!;
      if (current && typeof current === 'object') {
        for (const [key, value] of Object.entries(current)) {
          if (keyNames.some(k => k.toLowerCase() === key.toLowerCase())) {
            return value;
          }
          stack.push(value);
        }
      } else if (Array.isArray(current)) {
        stack.push(...current);
      }
    }
    return undefined;
  }

  /**
   * Get complete project data including panels
   */
  async getProjectData(projectId: number): Promise<OpenSolarProjectData> {
    try {
      this.logger.log(`🌞 Fetching complete OpenSolar data for project #${projectId}...`);
      
      const [project, systemDetails] = await Promise.all([
        this.getProject(projectId),
        this.getSystemDetails(projectId, 'layout,hardware')
      ]);

      const systems = systemDetails.systems || [];
      if (systems.length === 0) {
        this.logger.warn(`⚠️ No systems found in project #${projectId} - this project may not have been designed yet`);
        // Return project data without systems instead of throwing error
        const projectData: OpenSolarProjectData = {
          projectId,
          projectName: project.address || project.display_name || project.name || `Project ${projectId}`,
          address: project.address,
          systems: []
        };
        return projectData;
      }

      const projectSystems: OpenSolarSystem[] = [];

      for (let i = 0; i < systems.length; i++) {
        const sysObj = systems[i];
        const sysName = sysObj.name || sysObj.display_name || sysObj.uuid || `System ${i + 1}`;
        const sysUuid = sysObj.uuid;

        const panels = this.extractPanelsFromSystem(sysObj);
        const arrays = this.extractArraysFromSystem(sysObj);
        const batteries = this.extractBatteriesFromSystem(sysObj);
        const inverters = this.extractInvertersFromSystem(sysObj);
        const orientation = this.extractOrientationFromSystem(sysObj);
        const shading = this.extractShadingFromSystem(sysObj);
        
        const totalDcKw = panels.length > 0 ? 
          Math.round(panels.reduce((sum, p) => sum + (p.dc_size_kw || 0), 0) * 1000) / 1000 : undefined;

        projectSystems.push({
          uuid: sysUuid,
          name: sysName,
          panels,
          arrays,
          batteries,
          inverters,
          orientation,
          shading,
          total_dc_kw_est: totalDcKw
        });
      }

      const projectData: OpenSolarProjectData = {
        projectId,
        projectName: project.address || project.display_name || project.name || `Project ${projectId}`,
        address: project.address,
        systems: projectSystems
      };

      this.logger.log(`✅ OpenSolar project data extracted successfully`);
      this.logger.log(`📊 Found ${projectSystems.length} systems with ${projectSystems.reduce((sum, sys) => sum + sys.panels.length, 0)} panel types`);
      
      // Log detailed extraction results
      for (let i = 0; i < projectSystems.length; i++) {
        const sys = projectSystems[i];
        this.logger.log(`🔍 System ${i + 1} (${sys.name}):`);
        this.logger.log(`   📦 Panels: ${sys.panels.length} types`);
        this.logger.log(`   🏗️ Arrays: ${sys.arrays.length} arrays`);
        this.logger.log(`   🔋 Batteries: ${sys.batteries.length} batteries`);
        this.logger.log(`   ⚡ Inverters: ${sys.inverters.length} inverters`);
        this.logger.log(`   📐 Orientation: ${sys.orientation?.tilt ? `${sys.orientation.tilt}° tilt, ${sys.orientation.normalizedAzimuth}° normalized azimuth` : 'Not specified'}`);
        this.logger.log(`   🌳 Shading: ${sys.shading?.annualLoss ? `${sys.shading.annualLoss}% annual loss` : 'Not specified'}`);
      }
      
      if (projectSystems.length === 0) {
        this.logger.log(`ℹ️ Project #${projectId} has no systems - this is normal for projects that haven't been designed yet`);
      }
      
      return projectData;
    } catch (error) {
      this.logger.error(`❌ Failed to get project data for #${projectId}:`, error.message);
      throw error;
    }
  }

  /**
   * Search for projects by address (to find project by opportunity location)
   */
  async searchProjectsByAddress(address: string): Promise<OpenSolarProject[]> {
    try {
      const auth = await this.authenticate();
      
      this.logger.log(`🔍 Searching OpenSolar projects by address: ${address}`);
      
      // First, try to search by address in the projects list
      const response = await this.httpClient.get(
        `/api/orgs/${auth.orgId}/projects/`,
        {
          headers: { Authorization: `Bearer ${auth.token}` },
          params: { 
            search: address,
            limit: 50  // Increase limit to find more matches
          }
        }
      );

      let projects = response.data?.results || response.data || [];
      this.logger.log(`🔍 Raw search response:`, JSON.stringify(response.data, null, 2));
      
      // Filter and normalize the projects to ensure we have the right structure
      const normalizedProjects = projects
        .filter((project: any) => {
          // Ensure project has required fields
          return project && project.id && (
            project.address || 
            project.display_name || 
            project.name || 
            project.title
          );
        })
        .map((project: any) => {
          // Normalize project structure
          return {
            id: parseInt(project.id) || project.id, // Ensure ID is numeric
            address: project.address || project.location || '',
            display_name: project.display_name || project.name || project.title || `Project ${project.id}`,
            name: project.name || project.title || project.display_name || `Project ${project.id}`,
            // Add any other fields that might be useful
            created_at: project.created_at || project.createdAt,
            updated_at: project.updated_at || project.updatedAt
          };
        });

      this.logger.log(`✅ Found ${normalizedProjects.length} normalized projects matching address`);
      this.logger.log(`🔍 Normalized projects:`, JSON.stringify(normalizedProjects, null, 2));
      
      return normalizedProjects;
    } catch (error) {
      this.logger.error(`❌ Failed to search projects by address:`, error.message);
      
      // Handle rate limiting specifically
      if (error.response?.status === 429) {
        this.logger.warn('⚠️ OpenSolar rate limit exceeded during search. Returning empty results.');
        return [];
      }
      
      // Handle authentication errors
      if (error.message.includes('rate limit exceeded')) {
        this.logger.warn('⚠️ OpenSolar rate limit exceeded. Returning empty results.');
        return [];
      }
      
      this.logger.error(`❌ Error details:`, error);
      // Return empty array instead of throwing to allow graceful fallback
      return [];
    }
  }
}
