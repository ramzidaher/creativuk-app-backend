import { Injectable, OnModuleInit, OnModuleDestroy, Logger } from '@nestjs/common';
import { PrismaClient } from '@prisma/client';

@Injectable()
export class PrismaService extends PrismaClient implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(PrismaService.name);

  constructor() {
    super({
      datasources: {
        db: {
          url: process.env.DATABASE_URL,
        },
      },
      // Basic logging configuration
      log: ['error', 'warn'],
    });
  }

  async onModuleInit() {
    try {
      await this.$connect();
      this.logger.log('Database connected successfully');
      
      // Test connection with a simple query
      await this.$queryRaw`SELECT 1`;
      this.logger.log('Database connection test successful');
    } catch (error) {
      this.logger.error('Failed to connect to database:', error);
      throw error;
    }
  }

  async onModuleDestroy() {
    try {
      await this.$disconnect();
      this.logger.log('Database disconnected successfully');
    } catch (error) {
      this.logger.error('Error disconnecting from database:', error);
    }
  }

  /**
   * Get database connection status
   */
  async getConnectionStatus(): Promise<{
    connected: boolean;
    connectionCount: number;
    uptime: number;
  }> {
    try {
      const startTime = Date.now();
      await this.$queryRaw`SELECT 1`;
      const responseTime = Date.now() - startTime;
      
      return {
        connected: true,
        connectionCount: 0, // Prisma doesn't expose this directly
        uptime: process.uptime(),
      };
    } catch (error) {
      return {
        connected: false,
        connectionCount: 0,
        uptime: process.uptime(),
      };
    }
  }

  /**
   * Execute a transaction with retry logic
   */
  async executeWithRetry<T>(
    operation: (prisma: PrismaClient) => Promise<T>,
    maxRetries: number = 3,
    delay: number = 1000
  ): Promise<T> {
    let lastError: Error;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await operation(this);
      } catch (error) {
        lastError = error as Error;
        
        if (attempt === maxRetries) {
          this.logger.error(`Operation failed after ${maxRetries} attempts:`, error);
          throw error;
        }
        
        this.logger.warn(`Operation attempt ${attempt} failed, retrying in ${delay}ms:`, error);
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2; // Exponential backoff
      }
    }
    
    throw lastError!;
  }

  /**
   * Batch operations for better performance
   */
  async batchOperations<T>(
    operations: Array<(prisma: PrismaClient) => Promise<T>>
  ): Promise<T[]> {
    const results: T[] = [];
    
    // Process operations in batches of 10 to avoid overwhelming the database
    const batchSize = 10;
    for (let i = 0; i < operations.length; i += batchSize) {
      const batch = operations.slice(i, i + batchSize);
      const batchResults = await Promise.all(
        batch.map(operation => operation(this))
      );
      results.push(...batchResults);
    }
    
    return results;
  }
}
