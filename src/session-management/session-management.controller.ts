import { Controller, Post, Get, Body, Param, Logger } from '@nestjs/common';
import { SessionManagementService } from './session-management.service';

@Controller('session')
export class SessionManagementController {
  private readonly logger = new Logger(SessionManagementController.name);

  constructor(private readonly sessionManagementService: SessionManagementService) {}

  /**
   * Create or get user session
   */
  @Post('create')
  async createSession(@Body() body: { userId: string }) {
    try {
      const session = await this.sessionManagementService.createOrGetSession(body.userId);
      return {
        success: true,
        session,
        message: 'Session created successfully'
      };
    } catch (error) {
      this.logger.error('Failed to create session:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Queue a COM operation (Excel, PowerPoint, etc.)
   */
  @Post('queue/com')
  async queueComOperation(@Body() body: {
    userId: string;
    operation: string;
    data: any;
    priority?: number;
  }) {
    try {
      const result = await this.sessionManagementService.queueRequest(
        body.userId,
        body.operation,
        'com',
        body.data,
        body.priority || 1
      );
      return {
        success: true,
        result,
        message: 'COM operation queued successfully'
      };
    } catch (error) {
      this.logger.error('Failed to queue COM operation:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Queue a non-COM operation (file processing, data processing, etc.)
   */
  @Post('queue/non-com')
  async queueNonComOperation(@Body() body: {
    userId: string;
    operation: string;
    data: any;
    priority?: number;
  }) {
    try {
      const result = await this.sessionManagementService.queueRequest(
        body.userId,
        body.operation,
        'non-com',
        body.data,
        body.priority || 2
      );
      return {
        success: true,
        result,
        message: 'Non-COM operation queued successfully'
      };
    } catch (error) {
      this.logger.error('Failed to queue non-COM operation:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Queue a database operation
   */
  @Post('queue/database')
  async queueDatabaseOperation(@Body() body: {
    userId: string;
    operation: string;
    data: any;
    priority?: number;
  }) {
    try {
      const result = await this.sessionManagementService.queueRequest(
        body.userId,
        body.operation,
        'database',
        body.data,
        body.priority || 3
      );
      return {
        success: true,
        result,
        message: 'Database operation queued successfully'
      };
    } catch (error) {
      this.logger.error('Failed to queue database operation:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Queue an API operation
   */
  @Post('queue/api')
  async queueApiOperation(@Body() body: {
    userId: string;
    operation: string;
    data: any;
    priority?: number;
  }) {
    try {
      const result = await this.sessionManagementService.queueRequest(
        body.userId,
        body.operation,
        'api',
        body.data,
        body.priority || 4
      );
      return {
        success: true,
        result,
        message: 'API operation queued successfully'
      };
    } catch (error) {
      this.logger.error('Failed to queue API operation:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Get detailed queue status
   */
  @Get('queue/status')
  async getQueueStatus() {
    try {
      const status = this.sessionManagementService.getQueueStatus();
      return {
        success: true,
        status,
        message: 'Queue status retrieved successfully'
      };
    } catch (error) {
      this.logger.error('Failed to get queue status:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Clean up user session
   */
  @Post('cleanup/:userId')
  async cleanupUserSession(@Param('userId') userId: string) {
    try {
      await this.sessionManagementService.cleanupUserSession(userId);
      return {
        success: true,
        message: `Session cleaned up for user ${userId}`
      };
    } catch (error) {
      this.logger.error('Failed to cleanup user session:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Test endpoint to demonstrate optimized queuing
   */
  @Post('test/performance')
  async testPerformance(@Body() body: {
    userId: string;
    operations: Array<{
      type: 'com' | 'non-com' | 'database' | 'api';
      operation: string;
      data: any;
      priority?: number;
    }>;
  }) {
    try {
      const startTime = Date.now();
      const results: any[] = [];

      // Queue all operations simultaneously
      const promises = body.operations.map(op => 
        this.sessionManagementService.queueRequest(
          body.userId,
          op.operation,
          op.type,
          op.data,
          op.priority || 1
        )
      );

      // Wait for all operations to complete
      const operationResults = await Promise.all(promises);
      const endTime = Date.now();

      results.push(...operationResults);

      return {
        success: true,
        results,
        performance: {
          totalOperations: body.operations.length,
          totalTime: endTime - startTime,
          averageTimePerOperation: (endTime - startTime) / body.operations.length,
          operationsByType: body.operations.reduce((acc, op) => {
            acc[op.type] = (acc[op.type] || 0) + 1;
            return acc;
          }, {} as Record<string, number>)
        },
        message: 'Performance test completed successfully'
      };
    } catch (error) {
      this.logger.error('Performance test failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
}