import { Injectable, Logger, OnModuleDestroy } from '@nestjs/common';

interface CacheItem<T> {
  data: T;
  timestamp: number;
  ttl: number;
}

@Injectable()
export class CacheService implements OnModuleDestroy {
  private readonly logger = new Logger(CacheService.name);
  private readonly cache = new Map<string, CacheItem<any>>();
  private readonly defaultTTL = 5 * 60 * 1000; // 5 minutes
  private cleanupInterval: NodeJS.Timeout;

  constructor() {
    // Clean up expired cache entries every minute
    this.cleanupInterval = setInterval(() => {
      this.cleanupExpiredEntries();
    }, 60 * 1000);
  }

  onModuleDestroy() {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
    }
  }

  /**
   * Set a cache entry
   */
  set<T>(key: string, data: T, ttl?: number): void {
    const item: CacheItem<T> = {
      data,
      timestamp: Date.now(),
      ttl: ttl || this.defaultTTL
    };
    
    this.cache.set(key, item);
    this.logger.debug(`Cache set: ${key} (TTL: ${item.ttl}ms)`);
  }

  /**
   * Get a cache entry
   */
  get<T>(key: string): T | null {
    const item = this.cache.get(key);
    
    if (!item) {
      return null;
    }
    
    // Check if expired
    if (Date.now() - item.timestamp > item.ttl) {
      this.cache.delete(key);
      this.logger.debug(`Cache expired: ${key}`);
      return null;
    }
    
    this.logger.debug(`Cache hit: ${key}`);
    return item.data as T;
  }

  /**
   * Get or set a cache entry
   */
  async getOrSet<T>(
    key: string, 
    factory: () => Promise<T>, 
    ttl?: number
  ): Promise<T> {
    const cached = this.get<T>(key);
    
    if (cached !== null) {
      return cached;
    }
    
    this.logger.debug(`Cache miss: ${key}, fetching data...`);
    const data = await factory();
    this.set(key, data, ttl);
    return data;
  }

  /**
   * Delete a cache entry
   */
  delete(key: string): boolean {
    const deleted = this.cache.delete(key);
    if (deleted) {
      this.logger.debug(`Cache deleted: ${key}`);
    }
    return deleted;
  }

  /**
   * Clear all cache entries
   */
  clear(): void {
    this.cache.clear();
    this.logger.log('Cache cleared');
  }

  /**
   * Get cache statistics
   */
  getStats(): {
    size: number;
    entries: Array<{
      key: string;
      age: number;
      ttl: number;
      expired: boolean;
    }>;
  } {
    const now = Date.now();
    const entries = Array.from(this.cache.entries()).map(([key, item]) => ({
      key,
      age: now - item.timestamp,
      ttl: item.ttl,
      expired: now - item.timestamp > item.ttl
    }));

    return {
      size: this.cache.size,
      entries
    };
  }

  /**
   * Clean up expired entries
   */
  private cleanupExpiredEntries(): void {
    const now = Date.now();
    let cleanedCount = 0;
    
    for (const [key, item] of this.cache.entries()) {
      if (now - item.timestamp > item.ttl) {
        this.cache.delete(key);
        cleanedCount++;
      }
    }
    
    if (cleanedCount > 0) {
      this.logger.debug(`Cleaned up ${cleanedCount} expired cache entries`);
    }
  }

  /**
   * Generate cache key for user-specific data
   */
  generateUserKey(userId: string, operation: string, ...params: any[]): string {
    const paramString = params.length > 0 ? `_${params.join('_')}` : '';
    return `user_${userId}_${operation}${paramString}`;
  }

  /**
   * Generate cache key for opportunity-specific data
   */
  generateOpportunityKey(opportunityId: string, operation: string, ...params: any[]): string {
    const paramString = params.length > 0 ? `_${params.join('_')}` : '';
    return `opportunity_${opportunityId}_${operation}${paramString}`;
  }

  /**
   * Generate cache key for API responses
   */
  generateApiKey(endpoint: string, params: Record<string, any> = {}): string {
    const paramString = Object.keys(params).length > 0 
      ? `_${JSON.stringify(params)}` 
      : '';
    return `api_${endpoint}${paramString}`;
  }

  /**
   * Invalidate cache entries by pattern
   */
  invalidatePattern(pattern: string): number {
    const regex = new RegExp(pattern);
    let invalidatedCount = 0;
    
    for (const key of this.cache.keys()) {
      if (regex.test(key)) {
        this.cache.delete(key);
        invalidatedCount++;
      }
    }
    
    if (invalidatedCount > 0) {
      this.logger.debug(`Invalidated ${invalidatedCount} cache entries matching pattern: ${pattern}`);
    }
    
    return invalidatedCount;
  }

  /**
   * Invalidate user-specific cache entries
   */
  invalidateUser(userId: string): number {
    return this.invalidatePattern(`^user_${userId}_`);
  }

  /**
   * Invalidate opportunity-specific cache entries
   */
  invalidateOpportunity(opportunityId: string): number {
    return this.invalidatePattern(`^opportunity_${opportunityId}_`);
  }
}





