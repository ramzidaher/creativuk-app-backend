import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import * as path from 'path';
import * as fs from 'fs';

export interface UserSession {
  userId: string;
  sessionId: string;
  startTime: Date;
  lastActivity: Date;
  activeOperations: Set<string>;
  comProcesses: {
    excel?: number;
    powerpoint?: number;
  };
  workingDirectory: string;
}

export interface QueuedRequest {
  id: string;
  userId: string;
  operation: string;
  operationType: 'com' | 'non-com' | 'database' | 'api';
  priority: number;
  queuedAt: Date;
  startedAt?: Date;
  completedAt?: Date;
  status: 'queued' | 'processing' | 'completed' | 'failed';
  data: any;
  resolve: (value: any) => void;
  reject: (error: any) => void;
}

@Injectable()
export class SessionManagementService {
  private readonly logger = new Logger(SessionManagementService.name);
  private readonly activeSessions = new Map<string, UserSession>();
  private readonly requestQueue: QueuedRequest[] = [];
  private readonly processingRequests = new Map<string, QueuedRequest>();
  
  // Optimized concurrent operation limits
  private readonly maxConcurrentComOperations = 5; // Increased for COM operations
  private readonly maxConcurrentNonComOperations = 15; // Separate limit for non-COM operations
  private readonly maxConcurrentDatabaseOperations = 20; // Database operations
  private readonly maxConcurrentApiOperations = 10; // API operations
  
  private readonly sessionTimeout = 30 * 60 * 1000; // 30 minutes
  private readonly baseWorkingDir = process.cwd();
  
  // Operation type counters
  private readonly activeComOperations = new Set<string>();
  private readonly activeNonComOperations = new Set<string>();
  private readonly activeDatabaseOperations = new Set<string>();
  private readonly activeApiOperations = new Set<string>();

  constructor(private readonly prisma: PrismaService) {
    // Start cleanup timer
    setInterval(() => this.cleanupExpiredSessions(), 5 * 60 * 1000); // Every 5 minutes
    setInterval(() => this.processQueue(), 500); // Process queue every 500ms for faster processing
  }

  /**
   * Create or get user session
   */
  async createOrGetSession(userId: string): Promise<UserSession> {
    const existingSession = this.activeSessions.get(userId);
    
    if (existingSession && this.isSessionValid(existingSession)) {
      existingSession.lastActivity = new Date();
      return existingSession;
    }

    // Create new session
    const sessionId = this.generateSessionId();
    const workingDirectory = this.createUserWorkingDirectory(userId, sessionId);
    
    const session: UserSession = {
      userId,
      sessionId,
      startTime: new Date(),
      lastActivity: new Date(),
      activeOperations: new Set(),
      comProcesses: {},
      workingDirectory
    };

    this.activeSessions.set(userId, session);
    this.logger.log(`Created new session for user ${userId}: ${sessionId}`);
    
    return session;
  }

  /**
   * Queue an operation request with optimized handling
   */
  async queueRequest(
    userId: string,
    operation: string,
    operationType: 'com' | 'non-com' | 'database' | 'api',
    data: any,
    priority: number = 1
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      const requestId = this.generateRequestId();
      
      const request: QueuedRequest = {
        id: requestId,
        userId,
        operation,
        operationType,
        priority,
        queuedAt: new Date(),
        status: 'queued',
        data,
        resolve,
        reject
      };

      // For non-COM operations, try to process immediately if capacity allows
      if (operationType !== 'com' && this.canProcessOperationType(operationType)) {
        this.processRequestImmediately(request);
        return;
      }

      // Insert request in priority order
      const insertIndex = this.requestQueue.findIndex(r => r.priority < priority);
      if (insertIndex === -1) {
        this.requestQueue.push(request);
      } else {
        this.requestQueue.splice(insertIndex, 0, request);
      }

      this.logger.log(`Queued request ${requestId} for user ${userId}, operation: ${operation}, type: ${operationType}, priority: ${priority}`);
    });
  }

  /**
   * Get user session
   */
  getUserSession(userId: string): UserSession | null {
    const session = this.activeSessions.get(userId);
    return session && this.isSessionValid(session) ? session : null;
  }

  /**
   * Check if we can process an operation type based on current capacity
   */
  private canProcessOperationType(operationType: 'com' | 'non-com' | 'database' | 'api'): boolean {
    switch (operationType) {
      case 'com':
        return this.activeComOperations.size < this.maxConcurrentComOperations;
      case 'non-com':
        return this.activeNonComOperations.size < this.maxConcurrentNonComOperations;
      case 'database':
        return this.activeDatabaseOperations.size < this.maxConcurrentDatabaseOperations;
      case 'api':
        return this.activeApiOperations.size < this.maxConcurrentApiOperations;
      default:
        return false;
    }
  }

  /**
   * Process a request immediately without queuing
   */
  private async processRequestImmediately(request: QueuedRequest): Promise<void> {
    const session = this.getUserSession(request.userId);
    if (!session) {
      request.reject(new Error('User session expired'));
      return;
    }

    // Add to active operations
    this.addToActiveOperations(request);

    request.status = 'processing';
    request.startedAt = new Date();
    this.processingRequests.set(request.id, request);

    this.logger.log(`Processing request ${request.id} immediately for user ${request.userId}, type: ${request.operationType}`);

    try {
      const result = await this.executeOperation(request, session);
      request.status = 'completed';
      request.completedAt = new Date();
      request.resolve(result);
    } catch (error) {
      request.status = 'failed';
      request.completedAt = new Date();
      request.reject(error);
    } finally {
      this.removeFromActiveOperations(request);
      this.processingRequests.delete(request.id);
    }
  }

  /**
   * Add request to active operations tracking
   */
  private addToActiveOperations(request: QueuedRequest): void {
    switch (request.operationType) {
      case 'com':
        this.activeComOperations.add(request.id);
        break;
      case 'non-com':
        this.activeNonComOperations.add(request.id);
        break;
      case 'database':
        this.activeDatabaseOperations.add(request.id);
        break;
      case 'api':
        this.activeApiOperations.add(request.id);
        break;
    }
  }

  /**
   * Remove request from active operations tracking
   */
  private removeFromActiveOperations(request: QueuedRequest): void {
    switch (request.operationType) {
      case 'com':
        this.activeComOperations.delete(request.id);
        break;
      case 'non-com':
        this.activeNonComOperations.delete(request.id);
        break;
      case 'database':
        this.activeDatabaseOperations.delete(request.id);
        break;
      case 'api':
        this.activeApiOperations.delete(request.id);
        break;
    }
  }

  /**
   * Check if user has active operations
   */
  hasActiveOperations(userId: string): boolean {
    const session = this.getUserSession(userId);
    return session ? session.activeOperations.size > 0 : false;
  }

  /**
   * Start operation for user
   */
  startOperation(userId: string, operationId: string): boolean {
    const session = this.getUserSession(userId);
    if (!session) {
      this.logger.warn(`No valid session found for user ${userId}`);
      return false;
    }

    session.activeOperations.add(operationId);
    session.lastActivity = new Date();
    this.logger.log(`Started operation ${operationId} for user ${userId}`);
    return true;
  }

  /**
   * Complete operation for user
   */
  completeOperation(userId: string, operationId: string): void {
    const session = this.getUserSession(userId);
    if (session) {
      session.activeOperations.delete(operationId);
      session.lastActivity = new Date();
      this.logger.log(`Completed operation ${operationId} for user ${userId}`);
    }
  }

  /**
   * Get user's working directory
   */
  getUserWorkingDirectory(userId: string): string | null {
    const session = this.getUserSession(userId);
    return session ? session.workingDirectory : null;
  }

  /**
   * Clean up user session
   */
  async cleanupUserSession(userId: string): Promise<void> {
    const session = this.activeSessions.get(userId);
    if (session) {
      // Kill any active COM processes
      await this.killUserComProcesses(session);
      
      // Clean up working directory
      await this.cleanupUserWorkingDirectory(session.workingDirectory);
      
      // Remove session
      this.activeSessions.delete(userId);
      this.logger.log(`Cleaned up session for user ${userId}`);
    }
  }

  /**
   * Get queue status with detailed operation type breakdown
   */
  getQueueStatus(): {
    queued: number;
    processing: number;
    totalActiveSessions: number;
    operationTypes: {
      com: { queued: number; active: number; maxConcurrent: number };
      nonCom: { queued: number; active: number; maxConcurrent: number };
      database: { queued: number; active: number; maxConcurrent: number };
      api: { queued: number; active: number; maxConcurrent: number };
    };
  } {
    const queuedByType = {
      com: 0,
      nonCom: 0,
      database: 0,
      api: 0
    };

    this.requestQueue.forEach(request => {
      switch (request.operationType) {
        case 'com':
          queuedByType.com++;
          break;
        case 'non-com':
          queuedByType.nonCom++;
          break;
        case 'database':
          queuedByType.database++;
          break;
        case 'api':
          queuedByType.api++;
          break;
      }
    });

    return {
      queued: this.requestQueue.length,
      processing: this.processingRequests.size,
      totalActiveSessions: this.activeSessions.size,
      operationTypes: {
        com: {
          queued: queuedByType.com,
          active: this.activeComOperations.size,
          maxConcurrent: this.maxConcurrentComOperations
        },
        nonCom: {
          queued: queuedByType.nonCom,
          active: this.activeNonComOperations.size,
          maxConcurrent: this.maxConcurrentNonComOperations
        },
        database: {
          queued: queuedByType.database,
          active: this.activeDatabaseOperations.size,
          maxConcurrent: this.maxConcurrentDatabaseOperations
        },
        api: {
          queued: queuedByType.api,
          active: this.activeApiOperations.size,
          maxConcurrent: this.maxConcurrentApiOperations
        }
      }
    };
  }

  /**
   * Process the request queue with optimized operation type handling
   */
  private async processQueue(): Promise<void> {
    // Find the next request that can be processed based on operation type capacity
    let nextRequestIndex = -1;
    let nextRequest: QueuedRequest | null = null;

    for (let i = 0; i < this.requestQueue.length; i++) {
      const request = this.requestQueue[i];
      if (this.canProcessOperationType(request.operationType)) {
        nextRequestIndex = i;
        nextRequest = request;
        break;
      }
    }

    if (!nextRequest || nextRequestIndex === -1) {
      return; // No processable requests in queue
    }

    // Remove the request from queue
    this.requestQueue.splice(nextRequestIndex, 1);

    // Check if user session is still valid
    const session = this.getUserSession(nextRequest.userId);
    if (!session) {
      nextRequest.reject(new Error('User session expired'));
      return;
    }

    // Add to active operations tracking
    this.addToActiveOperations(nextRequest);

    // Start processing
    nextRequest.status = 'processing';
    nextRequest.startedAt = new Date();
    this.processingRequests.set(nextRequest.id, nextRequest);

    this.logger.log(`Processing request ${nextRequest.id} for user ${nextRequest.userId}, type: ${nextRequest.operationType}`);

    try {
      // Execute the operation
      const result = await this.executeOperation(nextRequest, session);
      
      nextRequest.status = 'completed';
      nextRequest.completedAt = new Date();
      nextRequest.resolve(result);
    } catch (error) {
      nextRequest.status = 'failed';
      nextRequest.completedAt = new Date();
      nextRequest.reject(error);
    } finally {
      this.removeFromActiveOperations(nextRequest);
      this.processingRequests.delete(nextRequest.id);
    }
  }

  /**
   * Execute the actual operation with operation type handling
   */
  private async executeOperation(request: QueuedRequest, session: UserSession): Promise<any> {
    const { operation, operationType, data } = request;
    
    // Start operation tracking
    this.startOperation(session.userId, request.id);

    try {
      // Handle different operation types
      switch (operationType) {
        case 'com':
          return await this.executeComOperation(operation, data, session);
        case 'non-com':
          return await this.executeNonComOperation(operation, data, session);
        case 'database':
          return await this.executeDatabaseOperation(operation, data, session);
        case 'api':
          return await this.executeApiOperation(operation, data, session);
        default:
          throw new Error(`Unknown operation type: ${operationType}`);
      }
    } finally {
      this.completeOperation(session.userId, request.id);
    }
  }

  /**
   * Execute COM operations (Excel, PowerPoint, etc.)
   */
  private async executeComOperation(operation: string, data: any, session: UserSession): Promise<any> {
    switch (operation) {
      case 'excel_calculation':
        return await this.executeExcelCalculation(data, session);
      case 'epvs_calculation':
        return await this.executeEPVSCalculation(data, session);
      case 'powerpoint_generation':
        return await this.executePowerPointGeneration(data, session);
      case 'pdf_conversion':
        return await this.executePdfConversion(data, session);
      default:
        throw new Error(`Unknown COM operation: ${operation}`);
    }
  }

  /**
   * Execute non-COM operations (file operations, data processing, etc.)
   */
  private async executeNonComOperation(operation: string, data: any, session: UserSession): Promise<any> {
    switch (operation) {
      case 'file_processing':
        return await this.executeFileProcessing(data, session);
      case 'data_processing':
        return await this.executeDataProcessing(data, session);
      case 'image_processing':
        return await this.executeImageProcessing(data, session);
      default:
        throw new Error(`Unknown non-COM operation: ${operation}`);
    }
  }

  /**
   * Execute database operations
   */
  private async executeDatabaseOperation(operation: string, data: any, session: UserSession): Promise<any> {
    switch (operation) {
      case 'pricing_save':
        return await this.executePricingSave(data, session);
      case 'survey_save':
        return await this.executeSurveySave(data, session);
      case 'opportunity_update':
        return await this.executeOpportunityUpdate(data, session);
      default:
        throw new Error(`Unknown database operation: ${operation}`);
    }
  }

  /**
   * Execute API operations (external API calls, etc.)
   */
  private async executeApiOperation(operation: string, data: any, session: UserSession): Promise<any> {
    switch (operation) {
      case 'ghl_api_call':
        return await this.executeGhlApiCall(data, session);
      case 'external_api_call':
        return await this.executeExternalApiCall(data, session);
      default:
        throw new Error(`Unknown API operation: ${operation}`);
    }
  }

  /**
   * Execute file processing operations
   */
  private async executeFileProcessing(data: any, session: UserSession): Promise<any> {
    this.logger.log(`Executing file processing for user ${session.userId}`);
    // Placeholder implementation - will be implemented with actual file processing logic
    return { success: true, message: 'File processing completed' };
  }

  /**
   * Execute data processing operations
   */
  private async executeDataProcessing(data: any, session: UserSession): Promise<any> {
    this.logger.log(`Executing data processing for user ${session.userId}`);
    // Placeholder implementation - will be implemented with actual data processing logic
    return { success: true, message: 'Data processing completed' };
  }

  /**
   * Execute image processing operations
   */
  private async executeImageProcessing(data: any, session: UserSession): Promise<any> {
    this.logger.log(`Executing image processing for user ${session.userId}`);
    // Placeholder implementation - will be implemented with actual image processing logic
    return { success: true, message: 'Image processing completed' };
  }

  /**
   * Execute survey save operations
   */
  private async executeSurveySave(data: any, session: UserSession): Promise<any> {
    this.logger.log(`Executing survey save for user ${session.userId}`);
    // Placeholder implementation - will be implemented with actual survey save logic
    return { success: true, message: 'Survey save completed' };
  }

  /**
   * Execute opportunity update operations
   */
  private async executeOpportunityUpdate(data: any, session: UserSession): Promise<any> {
    this.logger.log(`Executing opportunity update for user ${session.userId}`);
    // Placeholder implementation - will be implemented with actual opportunity update logic
    return { success: true, message: 'Opportunity update completed' };
  }

  /**
   * Execute GHL API call operations
   */
  private async executeGhlApiCall(data: any, session: UserSession): Promise<any> {
    this.logger.log(`Executing GHL API call for user ${session.userId}`);
    // Placeholder implementation - will be implemented with actual GHL API logic
    return { success: true, message: 'GHL API call completed' };
  }

  /**
   * Execute external API call operations
   */
  private async executeExternalApiCall(data: any, session: UserSession): Promise<any> {
    this.logger.log(`Executing external API call for user ${session.userId}`);
    // Placeholder implementation - will be implemented with actual external API logic
    return { success: true, message: 'External API call completed' };
  }

  /**
   * Execute Excel calculation with user isolation
   */
  private async executeExcelCalculation(data: any, session: UserSession): Promise<any> {
    // This will be implemented by injecting the Excel automation service
    this.logger.log(`Executing Excel calculation for user ${session.userId}`);
    
    // For now, return a placeholder - this will be properly implemented
    // when we inject the ExcelAutomationService
    return { 
      success: true, 
      message: 'Excel calculation completed with user isolation',
      filePath: path.join(session.workingDirectory, 'excel', `calculation_${data.opportunityId}.xlsm`)
    };
  }

  /**
   * Execute EPVS calculation with user isolation
   */
  private async executeEPVSCalculation(data: any, session: UserSession): Promise<any> {
    // This will be implemented by injecting the EPVS automation service
    this.logger.log(`Executing EPVS calculation for user ${session.userId}`);
    
    // For now, return a placeholder - this will be properly implemented
    // when we inject the EPVSAutomationService
    return { 
      success: true, 
      message: 'EPVS calculation completed with user isolation',
      filePath: path.join(session.workingDirectory, 'excel', `epvs_calculation_${data.opportunityId}.xlsm`)
    };
  }

  /**
   * Execute PowerPoint generation with user isolation
   */
  private async executePowerPointGeneration(data: any, session: UserSession): Promise<any> {
    // Implementation will be added by the presentation service
    this.logger.log(`Executing PowerPoint generation for user ${session.userId}`);
    return { success: true, message: 'PowerPoint generation completed' };
  }

  /**
   * Execute pricing save with user isolation
   */
  private async executePricingSave(data: any, session: UserSession): Promise<any> {
    // This will be implemented by injecting the pricing service
    this.logger.log(`Executing pricing save for user ${session.userId}`);
    
    // For now, return a placeholder - this will be properly implemented
    // when we inject the PricingService
    return { 
      success: true, 
      message: 'Pricing save completed with user isolation'
    };
  }

  /**
   * Execute PDF conversion with user isolation
   */
  private async executePdfConversion(data: any, session: UserSession): Promise<any> {
    // Implementation will be added by the PDF service
    this.logger.log(`Executing PDF conversion for user ${session.userId}`);
    return { success: true, message: 'PDF conversion completed' };
  }

  /**
   * Create user-specific working directory
   */
  private createUserWorkingDirectory(userId: string, sessionId: string): string {
    const userDir = path.join(this.baseWorkingDir, 'user-sessions', userId, sessionId);
    
    if (!fs.existsSync(userDir)) {
      fs.mkdirSync(userDir, { recursive: true });
    }

    // Create subdirectories
    const subdirs = ['excel', 'powerpoint', 'pdf', 'temp'];
    subdirs.forEach(subdir => {
      const subdirPath = path.join(userDir, subdir);
      if (!fs.existsSync(subdirPath)) {
        fs.mkdirSync(subdirPath, { recursive: true });
      }
    });

    return userDir;
  }

  /**
   * Clean up user working directory
   */
  private async cleanupUserWorkingDirectory(workingDir: string): Promise<void> {
    try {
      if (fs.existsSync(workingDir)) {
        fs.rmSync(workingDir, { recursive: true, force: true });
        this.logger.log(`Cleaned up working directory: ${workingDir}`);
      }
    } catch (error) {
      this.logger.error(`Failed to clean up working directory ${workingDir}:`, error);
    }
  }

  /**
   * Kill user's COM processes
   */
  private async killUserComProcesses(session: UserSession): Promise<void> {
    const { exec } = require('child_process');
    const { promisify } = require('util');
    const execAsync = promisify(exec);

    try {
      // Kill Excel processes (only if they belong to this user)
      if (session.comProcesses.excel) {
        await execAsync(`taskkill /PID ${session.comProcesses.excel} /F`);
        this.logger.log(`Killed Excel process ${session.comProcesses.excel} for user ${session.userId}`);
      }

      // Kill PowerPoint processes (only if they belong to this user)
      if (session.comProcesses.powerpoint) {
        await execAsync(`taskkill /PID ${session.comProcesses.powerpoint} /F`);
        this.logger.log(`Killed PowerPoint process ${session.comProcesses.powerpoint} for user ${session.userId}`);
      }
    } catch (error) {
      this.logger.warn(`Failed to kill COM processes for user ${session.userId}:`, error);
    }
  }

  /**
   * Clean up expired sessions
   */
  private async cleanupExpiredSessions(): Promise<void> {
    const now = new Date();
    const expiredUsers: string[] = [];

    for (const [userId, session] of this.activeSessions.entries()) {
      if (now.getTime() - session.lastActivity.getTime() > this.sessionTimeout) {
        expiredUsers.push(userId);
      }
    }

    for (const userId of expiredUsers) {
      await this.cleanupUserSession(userId);
    }

    if (expiredUsers.length > 0) {
      this.logger.log(`Cleaned up ${expiredUsers.length} expired sessions`);
    }
  }

  /**
   * Check if session is valid
   */
  private isSessionValid(session: UserSession): boolean {
    const now = new Date();
    return now.getTime() - session.lastActivity.getTime() < this.sessionTimeout;
  }

  /**
   * Generate unique session ID
   */
  private generateSessionId(): string {
    return `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Generate unique request ID
   */
  private generateRequestId(): string {
    return `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }
}
