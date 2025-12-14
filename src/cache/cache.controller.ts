import { Controller, Get, Post, Delete, Body, Param, Logger } from '@nestjs/common';
import { CacheService } from './cache.service';

@Controller('cache')
export class CacheController {
  private readonly logger = new Logger(CacheController.name);

  constructor(private readonly cacheService: CacheService) {}

  /**
   * Get cache statistics
   */
  @Get('stats')
  async getCacheStats() {
    try {
      const stats = this.cacheService.getStats();
      return {
        success: true,
        stats,
        message: 'Cache statistics retrieved successfully'
      };
    } catch (error) {
      this.logger.error('Failed to get cache stats:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Clear all cache entries
   */
  @Post('clear')
  async clearCache() {
    try {
      this.cacheService.clear();
      return {
        success: true,
        message: 'Cache cleared successfully'
      };
    } catch (error) {
      this.logger.error('Failed to clear cache:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Delete a specific cache entry
   */
  @Delete(':key')
  async deleteCacheEntry(@Param('key') key: string) {
    try {
      const deleted = this.cacheService.delete(key);
      return {
        success: true,
        deleted,
        message: deleted ? 'Cache entry deleted successfully' : 'Cache entry not found'
      };
    } catch (error) {
      this.logger.error('Failed to delete cache entry:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Invalidate user-specific cache entries
   */
  @Post('invalidate/user/:userId')
  async invalidateUserCache(@Param('userId') userId: string) {
    try {
      const invalidatedCount = this.cacheService.invalidateUser(userId);
      return {
        success: true,
        invalidatedCount,
        message: `Invalidated ${invalidatedCount} cache entries for user ${userId}`
      };
    } catch (error) {
      this.logger.error('Failed to invalidate user cache:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Invalidate opportunity-specific cache entries
   */
  @Post('invalidate/opportunity/:opportunityId')
  async invalidateOpportunityCache(@Param('opportunityId') opportunityId: string) {
    try {
      const invalidatedCount = this.cacheService.invalidateOpportunity(opportunityId);
      return {
        success: true,
        invalidatedCount,
        message: `Invalidated ${invalidatedCount} cache entries for opportunity ${opportunityId}`
      };
    } catch (error) {
      this.logger.error('Failed to invalidate opportunity cache:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Invalidate cache entries by pattern
   */
  @Post('invalidate/pattern')
  async invalidatePattern(@Body() body: { pattern: string }) {
    try {
      const invalidatedCount = this.cacheService.invalidatePattern(body.pattern);
      return {
        success: true,
        invalidatedCount,
        message: `Invalidated ${invalidatedCount} cache entries matching pattern: ${body.pattern}`
      };
    } catch (error) {
      this.logger.error('Failed to invalidate pattern cache:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Test cache functionality
   */
  @Post('test')
  async testCache(@Body() body: { key: string; data: any; ttl?: number }) {
    try {
      const { key, data, ttl } = body;
      
      // Set cache entry
      this.cacheService.set(key, data, ttl);
      
      // Get cache entry
      const retrieved = this.cacheService.get(key);
      
      return {
        success: true,
        test: {
          key,
          originalData: data,
          retrievedData: retrieved,
          match: JSON.stringify(data) === JSON.stringify(retrieved),
          ttl: ttl || 300000 // 5 minutes default
        },
        message: 'Cache test completed successfully'
      };
    } catch (error) {
      this.logger.error('Cache test failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
}





