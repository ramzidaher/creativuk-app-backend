import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { ExcelAutomationService } from '../excel-automation/excel-automation.service';
import { EPVSAutomationService } from '../epvs-automation/epvs-automation.service';

export interface CalculatorProgressData {
  opportunityId: string;
  userId: string;
  calculatorType: 'off-peak' | 'flux' | 'epvs';
  currentStep: 'template-selection' | 'radio-buttons' | 'dynamic-inputs' | 'arrays' | 'pricing' | 'completed';
  
  // Template Selection Data
  templateSelection?: {
    selectedOptions: {
      solar: boolean;
      battery: boolean;
      solarHybrid: boolean;
      batteryInverter: boolean;
    };
    templateFileName: string;
  };
  
  // Calculator Screen Data
  radioButtonSelections?: Record<string, string>;
  
  // Dynamic Inputs Data
  dynamicInputs?: Record<string, string>;
  
  // Arrays Data
  arraysData?: {
    arrayRows: Array<{
      id: number;
      enabled: boolean;
      numberOfPanels?: string;
      panelSizeWp?: string; // Panel size in watts (array_X_panel_size_wp)
      arraySizeKwp?: string; // Array size in kWp (array_X_array_size_kwp)
      orientationDeg?: string;
      pitchDeg?: string;
      irradianceKk?: string; // Irradiance value (array_X_irradiance_kk)
      shadingFactor?: string;
      source?: 'opensolar' | 'manual';
      overrideOpenSolar?: boolean;
    }>;
    enabledCount: number;
  };
  
  // Pricing Data
  pricingData?: {
    selectedBatteryType: '5kW' | '10kW';
    selectedNumberOfPanels: number;
    additionalItemQuantities: Record<string, number>;
    paymentMethod: 'Cash' | 'Hometree' | 'New Finance' | null;
    totalSystemCost?: string; // Total system cost (total_system_cost)
    deposit: string;
    interestRate: string;
    interestRateType: string;
    paymentTerm: string;
  };
  
  // Customer Details
  customerDetails?: {
    customerName: string;
    address: string;
    postcode: string;
  };
  
  // Metadata
  lastSavedAt: string;
  completedSteps: Record<string, boolean>;
  dataHash?: string; // For change detection
}

@Injectable()
export class CalculatorProgressService {
  private readonly logger = new Logger(CalculatorProgressService.name);

  constructor(
    private prisma: PrismaService,
    private excelAutomationService: ExcelAutomationService,
    private epvsAutomationService: EPVSAutomationService
  ) {}

  /**
   * Save calculator progress data
   */
  async saveProgress(
    userId: string,
    opportunityId: string,
    calculatorType: 'off-peak' | 'flux' | 'epvs',
    progressData: Partial<CalculatorProgressData>
  ): Promise<{ success: boolean; message: string; dataHash?: string }> {
    try {
      this.logger.log(`Saving calculator progress for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);

      // Get existing progress
      const existingProgress = await this.getProgress(userId, opportunityId, calculatorType);
      
      // Create data hash for change detection
      const dataHash = this.generateDataHash(progressData);
      
      // Check if data has actually changed
      if (existingProgress && existingProgress.dataHash === dataHash) {
        this.logger.log(`No changes detected for opportunity ${opportunityId}, skipping save`);
        return { 
          success: true, 
          message: 'No changes detected, data not saved',
          dataHash: existingProgress.dataHash
        };
      }

      // Merge with existing progress data to preserve all previous selections
      this.logger.log(`Merging progress data for opportunity ${opportunityId}:`);
      this.logger.log(`Existing progress:`, existingProgress);
      this.logger.log(`New progress data:`, progressData);
      
      const mergedProgressData: CalculatorProgressData = {
        opportunityId,
        userId,
        calculatorType,
        currentStep: progressData.currentStep || existingProgress?.currentStep || 'template-selection',
        lastSavedAt: new Date().toISOString(),
        completedSteps: { ...existingProgress?.completedSteps, ...progressData.completedSteps },
        dataHash,
        // Preserve existing data and only update provided fields
        // Use hasOwnProperty to check if the field was explicitly provided
        templateSelection: progressData.hasOwnProperty('templateSelection') 
          ? progressData.templateSelection 
          : existingProgress?.templateSelection,
        radioButtonSelections: progressData.hasOwnProperty('radioButtonSelections')
          ? { ...existingProgress?.radioButtonSelections, ...progressData.radioButtonSelections }
          : existingProgress?.radioButtonSelections,
        dynamicInputs: progressData.hasOwnProperty('dynamicInputs')
          ? { ...existingProgress?.dynamicInputs, ...progressData.dynamicInputs }
          : existingProgress?.dynamicInputs,
        arraysData: progressData.hasOwnProperty('arraysData')
          ? progressData.arraysData
          : existingProgress?.arraysData,
        pricingData: progressData.hasOwnProperty('pricingData')
          ? progressData.pricingData
          : existingProgress?.pricingData,
        customerDetails: progressData.hasOwnProperty('customerDetails')
          ? progressData.customerDetails
          : existingProgress?.customerDetails,
      };
      
      this.logger.log(`Merged progress data:`, mergedProgressData);

      // Upsert the progress data
      await this.prisma.calculatorProgress.upsert({
        where: {
          userId_opportunityId_calculatorType: {
            userId,
            opportunityId,
            calculatorType,
          },
        },
        update: {
          data: mergedProgressData as any,
          updatedAt: new Date(),
        },
        create: {
          userId,
          opportunityId,
          calculatorType,
          data: mergedProgressData as any,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      });

      this.logger.log(`Successfully saved calculator progress for opportunity ${opportunityId}`);
      return { 
        success: true, 
        message: 'Progress saved successfully',
        dataHash
      };
    } catch (error) {
      this.logger.error(`Failed to save calculator progress: ${error.message}`);
      return { 
        success: false, 
        message: `Failed to save progress: ${error.message}` 
      };
    }
  }

  /**
   * Get calculator progress data
   */
  async getProgress(
    userId: string,
    opportunityId: string,
    calculatorType: 'off-peak' | 'flux' | 'epvs'
  ): Promise<CalculatorProgressData | null> {
    try {
      this.logger.log(`Getting progress for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
      
      const progress = await this.prisma.calculatorProgress.findUnique({
        where: {
          userId_opportunityId_calculatorType: {
            userId,
            opportunityId,
            calculatorType,
          },
        },
      });

      this.logger.log(`Database query result:`, progress);

      if (!progress) {
        this.logger.log(`No progress found in database for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);
        return null;
      }

      this.logger.log(`Progress data from database:`, progress.data);
      const result = progress.data as unknown as CalculatorProgressData;
      this.logger.log(`Converted progress data:`, result);
      
      return result;
    } catch (error) {
      this.logger.error(`Failed to get calculator progress: ${error.message}`);
      this.logger.error(`Error details:`, error);
      return null;
    }
  }

  /**
   * Check if data has changed since last save
   */
  async hasDataChanged(
    userId: string,
    opportunityId: string,
    calculatorType: 'off-peak' | 'flux' | 'epvs',
    newData: Partial<CalculatorProgressData>
  ): Promise<{ hasChanged: boolean; currentHash?: string; newHash: string }> {
    try {
      const existingProgress = await this.getProgress(userId, opportunityId, calculatorType);
      const newHash = this.generateDataHash(newData);
      
      return {
        hasChanged: !existingProgress || existingProgress.dataHash !== newHash,
        currentHash: existingProgress?.dataHash,
        newHash,
      };
    } catch (error) {
      this.logger.error(`Failed to check data changes: ${error.message}`);
      return {
        hasChanged: true,
        newHash: this.generateDataHash(newData),
      };
    }
  }

  /**
   * Clear calculator progress
   */
  async clearProgress(
    userId: string,
    opportunityId: string,
    calculatorType?: 'off-peak' | 'flux' | 'epvs'
  ): Promise<{ success: boolean; message: string }> {
    try {
      if (calculatorType) {
        await this.prisma.calculatorProgress.deleteMany({
          where: {
            userId,
            opportunityId,
            calculatorType,
          },
        });
      } else {
        await this.prisma.calculatorProgress.deleteMany({
          where: {
            userId,
            opportunityId,
          },
        });
      }

      this.logger.log(`Cleared calculator progress for user ${userId}, opportunity ${opportunityId}`);
      return { 
        success: true, 
        message: 'Progress cleared successfully' 
      };
    } catch (error) {
      this.logger.error(`Failed to clear calculator progress: ${error.message}`);
      return { 
        success: false, 
        message: `Failed to clear progress: ${error.message}` 
      };
    }
  }

  /**
   * Get progress summary
   */
  async getProgressSummary(
    userId: string,
    opportunityId: string,
    calculatorType: 'off-peak' | 'flux' | 'epvs'
  ): Promise<{
    hasProgress: boolean;
    currentStep: string;
    completedSteps: string[];
    lastSavedAt?: string;
    progressPercentage: number;
  }> {
    try {
      const progress = await this.getProgress(userId, opportunityId, calculatorType);
      
      if (!progress) {
        return {
          hasProgress: false,
          currentStep: 'template-selection',
          completedSteps: [],
          progressPercentage: 0,
        };
      }

      const totalSteps = 5; // template-selection, radio-buttons, dynamic-inputs, arrays, pricing
      const completedCount = Object.keys(progress.completedSteps || {}).length;
      const progressPercentage = Math.round((completedCount / totalSteps) * 100);

      return {
        hasProgress: true,
        currentStep: progress.currentStep,
        completedSteps: Object.keys(progress.completedSteps || {}),
        lastSavedAt: progress.lastSavedAt,
        progressPercentage,
      };
    } catch (error) {
      this.logger.error(`Failed to get progress summary: ${error.message}`);
      return {
        hasProgress: false,
        currentStep: 'template-selection',
        completedSteps: [],
        progressPercentage: 0,
      };
    }
  }

  /**
   * Generate a hash for data change detection
   */
  private generateDataHash(data: Partial<CalculatorProgressData>): string {
    // Create a deterministic hash of the data
    const dataString = JSON.stringify(data, Object.keys(data).sort());
    return this.simpleHash(dataString);
  }

  /**
   * Simple hash function for data comparison
   */
  private simpleHash(str: string): string {
    let hash = 0;
    if (str.length === 0) return hash.toString();
    
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32-bit integer
    }
    
    return Math.abs(hash).toString(36);
  }

  /**
   * Submit complete calculator data to Excel via COM
   * Takes all saved JSON progress data and inputs it into Excel in one batch operation
   */
  async submitCalculator(
    userId: string,
    opportunityId: string,
    calculatorType: 'off-peak' | 'flux' | 'epvs',
    existingFileName?: string
  ): Promise<{
    success: boolean;
    message: string;
    filePath?: string;
    error?: string;
  }> {
    this.logger.log(`📥 submitCalculator called with existingFileName: ${existingFileName || 'undefined'}`);
    try {
      this.logger.log(`🚀 SUBMIT CALCULATOR SERVICE CALLED`);
      this.logger.log(`📋 Parameters: userId=${userId}, opportunityId=${opportunityId}, calculatorType=${calculatorType}`);
      this.logger.log(`Submitting calculator data for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);

      // Get complete progress data
      const progressData = await this.getProgress(userId, opportunityId, calculatorType);
      
      if (!progressData) {
        return {
          success: false,
          message: 'No progress data found. Please complete the calculator steps first.',
        };
      }

      // Validate required fields
      if (!progressData.customerDetails) {
        return {
          success: false,
          message: 'Customer details are required. Please complete the customer details step first.',
        };
      }

      // Transform radioButtonSelections from Record<string, string> to string[]
      // The Record format is: { 'Category Name': 'ShapeName', ... }
      // We need to extract the VALUES (shape names), not the keys
      const radioButtonSelections: string[] = [];
      if (progressData.radioButtonSelections) {
        // Extract all values (shape names) from the Record
        radioButtonSelections.push(...Object.values(progressData.radioButtonSelections));
      }
      
      this.logger.log(`Extracted ${radioButtonSelections.length} radio button selections: ${JSON.stringify(radioButtonSelections)}`);

      // Build dynamic inputs from multiple sources
      const dynamicInputs: Record<string, string> = {
        ...(progressData.dynamicInputs || {}),
      };

      // Add arrays data to dynamicInputs
      if (progressData.arraysData?.arrayRows) {
        const isFlux = calculatorType === 'flux';
        const enabledArrays = progressData.arraysData.arrayRows.filter(row => row.enabled);
        
        // Add no_of_arrays
        dynamicInputs['no_of_arrays'] = String(enabledArrays.length);

        // Add each array's data
        enabledArrays.forEach((arrayRow) => {
          const arrayIndex = arrayRow.id;
          
          if (isFlux) {
            // Flux format
            if (arrayRow.numberOfPanels) {
              dynamicInputs[`array${arrayIndex}_panels`] = arrayRow.numberOfPanels;
            }
            if (arrayRow.orientationDeg) {
              dynamicInputs[`array${arrayIndex}_orientation`] = arrayRow.orientationDeg;
            }
            if (arrayRow.pitchDeg) {
              dynamicInputs[`array${arrayIndex}_pitch`] = arrayRow.pitchDeg;
            }
            if (arrayRow.shadingFactor) {
              dynamicInputs[`array${arrayIndex}_shading`] = arrayRow.shadingFactor;
            }
            if (arrayRow.panelSizeWp) {
              dynamicInputs[`array${arrayIndex}_panel_size_wp`] = arrayRow.panelSizeWp;
            }
            if (arrayRow.arraySizeKwp) {
              dynamicInputs[`array${arrayIndex}_array_size_kwp`] = arrayRow.arraySizeKwp;
            }
            if (arrayRow.irradianceKk) {
              dynamicInputs[`array${arrayIndex}_irradiance_kk`] = arrayRow.irradianceKk;
            }
          } else {
            // Off-peak format
            if (arrayRow.numberOfPanels) {
              dynamicInputs[`array_${arrayIndex}_num_panels`] = arrayRow.numberOfPanels;
            }
            if (arrayRow.orientationDeg) {
              dynamicInputs[`array_${arrayIndex}_orientation_deg_from_south`] = arrayRow.orientationDeg;
            }
            if (arrayRow.pitchDeg) {
              dynamicInputs[`array_${arrayIndex}_pitch_deg_from_flat`] = arrayRow.pitchDeg;
            }
            if (arrayRow.shadingFactor) {
              dynamicInputs[`array_${arrayIndex}_shading_factor`] = arrayRow.shadingFactor;
            }
            if (arrayRow.panelSizeWp) {
              dynamicInputs[`array_${arrayIndex}_panel_size_wp`] = arrayRow.panelSizeWp;
            }
            if (arrayRow.arraySizeKwp) {
              dynamicInputs[`array_${arrayIndex}_array_size_kwp`] = arrayRow.arraySizeKwp;
            }
            if (arrayRow.irradianceKk) {
              dynamicInputs[`array_${arrayIndex}_irradiance_kk`] = arrayRow.irradianceKk;
            }
          }
        });
      }

      // Add pricing data to dynamicInputs
      if (progressData.pricingData) {
        const pricing = progressData.pricingData;
        
        this.logger.log(`📋 Processing pricing data: ${JSON.stringify(pricing)}`);
        
        if (pricing.paymentMethod) {
          dynamicInputs['payment_method'] = pricing.paymentMethod;
          this.logger.log(`  ✓ Added payment_method: ${pricing.paymentMethod}`);
        }
        if (pricing.totalSystemCost !== undefined && pricing.totalSystemCost !== null && pricing.totalSystemCost !== '') {
          dynamicInputs['total_system_cost'] = pricing.totalSystemCost;
          this.logger.log(`  ✓ Added total_system_cost: ${pricing.totalSystemCost}`);
        } else {
          this.logger.warn(`  ⚠️ totalSystemCost is missing or empty in pricingData`);
        }
        if (pricing.deposit) {
          dynamicInputs['deposit'] = pricing.deposit;
          this.logger.log(`  ✓ Added deposit: ${pricing.deposit}`);
        }
        if (pricing.interestRate) {
          dynamicInputs['interest_rate'] = pricing.interestRate;
          this.logger.log(`  ✓ Added interest_rate: ${pricing.interestRate}`);
        }
        if (pricing.interestRateType) {
          dynamicInputs['interest_rate_type'] = pricing.interestRateType;
          this.logger.log(`  ✓ Added interest_rate_type: ${pricing.interestRateType}`);
        }
        if (pricing.paymentTerm) {
          dynamicInputs['payment_term'] = pricing.paymentTerm;
          this.logger.log(`  ✓ Added payment_term: ${pricing.paymentTerm}`);
        }
        
        this.logger.log(`📊 Total payment fields added to dynamicInputs: ${Object.keys(dynamicInputs).filter(k => k === 'payment_method' || k === 'total_system_cost' || k === 'deposit' || k === 'interest_rate' || k === 'interest_rate_type' || k === 'payment_term').length}`);
      } else {
        this.logger.warn(`⚠️ No pricingData found in progressData`);
      }

      // Extract template file name
      const templateFileName = progressData.templateSelection?.templateFileName;

      // For Flux/EPVS calculators: Fetch Flux rates from Octopus API and add to dynamicInputs
      if (calculatorType === 'flux' || calculatorType === 'epvs') {
        const postcode = progressData.customerDetails?.postcode;
        if (postcode) {
          try {
            this.logger.log(`🔌 Fetching Flux rates from Octopus API for postcode: ${postcode}`);
            
            // Fetch Flux rates from Octopus API
            const fluxRatesResult = await this.epvsAutomationService.getOctopusFluxRates(postcode);
            
            if (fluxRatesResult.success && fluxRatesResult.rates?.parsed_rates) {
              const rates = fluxRatesResult.rates.parsed_rates;
              this.logger.log(`✅ Successfully fetched Flux rates for ${postcode}`);
              
              // Add Flux rates to dynamicInputs (using named field keys that match EPVS cell mappings)
              // Import rates: H22, H23, H24
              // Export rates: J22, J23, J24
              if (rates.import) {
                dynamicInputs['import_day_rate'] = (Math.round(rates.import.day * 100) / 100).toString(); // Import Day Rate (H22)
                dynamicInputs['import_flux_rate'] = (Math.round(rates.import.flux * 100) / 100).toString();  // Import Flux Rate (H23)
                dynamicInputs['import_peak_rate'] = (Math.round(rates.import.peak * 100) / 100).toString(); // Import Peak Rate (H24)
              }
              if (rates.export) {
                dynamicInputs['export_day_rate'] = (Math.round(rates.export.day * 100) / 100).toString(); // Export Day Rate (J22)
                dynamicInputs['export_flux_rate'] = (Math.round(rates.export.flux * 100) / 100).toString(); // Export Flux Rate (J23)
                dynamicInputs['export_peak_rate'] = (Math.round(rates.export.peak * 100) / 100).toString(); // Export Peak Rate (J24)
              }
              
              this.logger.log(`📊 Added Flux rates to dynamicInputs: import_day_rate=${dynamicInputs['import_day_rate']}, import_flux_rate=${dynamicInputs['import_flux_rate']}, import_peak_rate=${dynamicInputs['import_peak_rate']}, export_day_rate=${dynamicInputs['export_day_rate']}, export_flux_rate=${dynamicInputs['export_flux_rate']}, export_peak_rate=${dynamicInputs['export_peak_rate']}`);
            } else {
              this.logger.warn(`⚠️ Failed to fetch Flux rates: ${fluxRatesResult.error || 'Unknown error'}`);
              // Continue without Flux rates - user can enter them manually
            }
          } catch (fluxError) {
            this.logger.warn(`⚠️ Error fetching Flux rates from Octopus API: ${fluxError.message}`);
            // Continue without Flux rates - user can enter them manually
          }
        } else {
          this.logger.warn(`⚠️ No postcode provided for Flux calculator - skipping Flux rates fetch`);
        }
      }

      // Log the data being submitted for debugging
      this.logger.log(`📤 Submitting calculator with the following data:`);
      this.logger.log(`   - Opportunity ID: ${opportunityId}`);
      this.logger.log(`   - Calculator Type: ${calculatorType}`);
      this.logger.log(`   - Template File: ${templateFileName || 'N/A'}`);
      this.logger.log(`   - Customer Details: ${JSON.stringify(progressData.customerDetails)}`);
      this.logger.log(`   - Radio Button Selections (${radioButtonSelections.length}): ${JSON.stringify(radioButtonSelections)}`);
      this.logger.log(`   - Dynamic Inputs (${Object.keys(dynamicInputs).length}): ${JSON.stringify(Object.keys(dynamicInputs))}`);
      this.logger.log(`   - Arrays Enabled: ${progressData.arraysData?.enabledCount || 0}`);

      // Route to the correct automation service based on calculator type
      // This performs ALL COM operations in ONE batch: create file, customer details, radio buttons, inputs, arrays, pricing, Flux rates
      let result;
      if (calculatorType === 'flux' || calculatorType === 'epvs') {
        // Use EPVS automation service for Flux/EPVS calculators
        this.logger.log(`🚀 Using EPVS automation service for ${calculatorType} calculator`);
        this.logger.log(`📋 Calling performCompleteCalculation with ${radioButtonSelections.length} radio buttons and ${Object.keys(dynamicInputs).length} dynamic inputs`);
        if (existingFileName) {
          this.logger.log(`📝 Editing existing file: ${existingFileName}`);
        }
        
        try {
          result = await this.epvsAutomationService.performCompleteCalculation(
            opportunityId,
            progressData.customerDetails,
            radioButtonSelections,
            dynamicInputs,
            templateFileName,
            existingFileName
          );
          this.logger.log(`✅ EPVS automation service returned: ${JSON.stringify(result)}`);
        } catch (error) {
          this.logger.error(`❌ EPVS automation service error: ${error.message}`);
          this.logger.error(`❌ Error stack: ${error.stack}`);
          throw error;
        }
      } else {
        // Use Excel automation service for Off-Peak calculator
        this.logger.log(`🚀 Using Excel automation service for ${calculatorType} calculator`);
        this.logger.log(`📋 Calling performCompleteCalculation with ${radioButtonSelections.length} radio buttons and ${Object.keys(dynamicInputs).length} dynamic inputs`);
        if (existingFileName) {
          this.logger.log(`📝 Editing existing file: ${existingFileName}`);
        }
        
        try {
          result = await this.excelAutomationService.performCompleteCalculation(
            opportunityId,
            progressData.customerDetails,
            radioButtonSelections,
            dynamicInputs,
            templateFileName,
            existingFileName
          );
          this.logger.log(`✅ Excel automation service returned: ${JSON.stringify(result)}`);
        } catch (error) {
          this.logger.error(`❌ Excel automation service error: ${error.message}`);
          this.logger.error(`❌ Error stack: ${error.stack}`);
          throw error;
        }
      }

      this.logger.log(`🎉 Calculator submission completed: ${JSON.stringify(result)}`);

      return result;
    } catch (error) {
      this.logger.error(`Error submitting calculator data: ${error.message}`);
      this.logger.error(`Error stack: ${error.stack}`);
      return {
        success: false,
        message: `Error submitting calculator data: ${error.message}`,
        error: error.message,
      };
    }
  }

  /**
   * Get submission status - checks if Excel file was created
   */
  async getSubmissionStatus(
    userId: string,
    opportunityId: string,
    calculatorType: 'off-peak' | 'flux' | 'epvs'
  ): Promise<{
    success: boolean;
    submitted: boolean;
    message: string;
    filePath?: string;
    error?: string;
  }> {
    try {
      this.logger.log(`Checking submission status for user ${userId}, opportunity ${opportunityId}, type ${calculatorType}`);

      // Check if Excel file exists by using ExcelAutomationService's check method
      // We need to map calculatorType to what ExcelAutomationService expects
      const excelCalculatorType = calculatorType === 'epvs' ? 'flux' : calculatorType;
      const result = await this.excelAutomationService.checkOpportunityFileExists(opportunityId);

      if (!result.success) {
        return {
          success: false,
          submitted: false,
          message: result.message || 'Error checking submission status',
          error: result.message,
        };
      }

      return {
        success: true,
        submitted: result.exists,
        message: result.exists 
          ? `Calculator submission found for ${opportunityId}` 
          : `No submission found for ${opportunityId}`,
        filePath: result.filePath,
      };
    } catch (error) {
      this.logger.error(`Error checking submission status: ${error.message}`);
      return {
        success: false,
        submitted: false,
        message: `Error checking submission status: ${error.message}`,
        error: error.message,
      };
    }
  }
}
