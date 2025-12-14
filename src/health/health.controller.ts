import { Controller, Get, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { SessionManagementService } from '../session-management/session-management.service';

@Controller('health')
export class HealthController {
  private readonly logger = new Logger(HealthController.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly sessionManagement: SessionManagementService
  ) {}

  /**
   * Get overall system health
   */
  @Get()
  async getHealth() {
    try {
      const startTime = Date.now();
      
      // Check database health
      const dbStatus = await this.prisma.getConnectionStatus();
      
      // Check session management health
      const queueStatus = this.sessionManagement.getQueueStatus();
      
      // Check memory usage
      const memoryUsage = process.memoryUsage();
      
      const responseTime = Date.now() - startTime;
      
      return {
        status: 'healthy',
        timestamp: new Date().toISOString(),
        responseTime: `${responseTime}ms`,
        services: {
          database: {
            status: dbStatus.connected ? 'healthy' : 'unhealthy',
            connected: dbStatus.connected,
            uptime: `${Math.round(dbStatus.uptime)}s`
          },
          sessionManagement: {
            status: 'healthy',
            activeSessions: queueStatus.totalActiveSessions,
            queuedOperations: queueStatus.queued,
            processingOperations: queueStatus.processing,
            operationTypes: queueStatus.operationTypes
          },
          memory: {
            used: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
            total: `${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
            external: `${Math.round(memoryUsage.external / 1024 / 1024)}MB`,
            rss: `${Math.round(memoryUsage.rss / 1024 / 1024)}MB`
          }
        }
      };
    } catch (error) {
      this.logger.error('Health check failed:', error);
      return {
        status: 'unhealthy',
        timestamp: new Date().toISOString(),
        error: error.message
      };
    }
  }

  /**
   * Get database health specifically
   */
  @Get('database')
  async getDatabaseHealth() {
    try {
      const startTime = Date.now();
      const dbStatus = await this.prisma.getConnectionStatus();
      const responseTime = Date.now() - startTime;
      
      return {
        status: dbStatus.connected ? 'healthy' : 'unhealthy',
        timestamp: new Date().toISOString(),
        responseTime: `${responseTime}ms`,
        connected: dbStatus.connected,
        uptime: `${Math.round(dbStatus.uptime)}s`,
        connectionCount: dbStatus.connectionCount
      };
    } catch (error) {
      this.logger.error('Database health check failed:', error);
      return {
        status: 'unhealthy',
        timestamp: new Date().toISOString(),
        error: error.message
      };
    }
  }

  /**
   * Get session management health
   */
  @Get('sessions')
  async getSessionHealth() {
    try {
      const queueStatus = this.sessionManagement.getQueueStatus();
      
      return {
        status: 'healthy',
        timestamp: new Date().toISOString(),
        activeSessions: queueStatus.totalActiveSessions,
        queuedOperations: queueStatus.queued,
        processingOperations: queueStatus.processing,
        operationTypes: queueStatus.operationTypes,
        performance: {
          comUtilization: `${Math.round((queueStatus.operationTypes.com.active / queueStatus.operationTypes.com.maxConcurrent) * 100)}%`,
          nonComUtilization: `${Math.round((queueStatus.operationTypes.nonCom.active / queueStatus.operationTypes.nonCom.maxConcurrent) * 100)}%`,
          databaseUtilization: `${Math.round((queueStatus.operationTypes.database.active / queueStatus.operationTypes.database.maxConcurrent) * 100)}%`,
          apiUtilization: `${Math.round((queueStatus.operationTypes.api.active / queueStatus.operationTypes.api.maxConcurrent) * 100)}%`
        }
      };
    } catch (error) {
      this.logger.error('Session health check failed:', error);
      return {
        status: 'unhealthy',
        timestamp: new Date().toISOString(),
        error: error.message
      };
    }
  }

  /**
   * Get performance metrics
   */
  @Get('performance')
  async getPerformanceMetrics() {
    try {
      const memoryUsage = process.memoryUsage();
      const cpuUsage = process.cpuUsage();
      const queueStatus = this.sessionManagement.getQueueStatus();
      
      return {
        timestamp: new Date().toISOString(),
        memory: {
          used: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
          total: `${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
          external: `${Math.round(memoryUsage.external / 1024 / 1024)}MB`,
          rss: `${Math.round(memoryUsage.rss / 1024 / 1024)}MB`
        },
        cpu: {
          user: `${Math.round(cpuUsage.user / 1000)}ms`,
          system: `${Math.round(cpuUsage.system / 1000)}ms`
        },
        queue: {
          totalQueued: queueStatus.queued,
          totalProcessing: queueStatus.processing,
          totalActiveSessions: queueStatus.totalActiveSessions,
          operationTypes: queueStatus.operationTypes
        },
        uptime: `${Math.round(process.uptime())}s`
      };
    } catch (error) {
      this.logger.error('Performance metrics failed:', error);
      return {
        status: 'error',
        timestamp: new Date().toISOString(),
        error: error.message
      };
    }
  }
}





