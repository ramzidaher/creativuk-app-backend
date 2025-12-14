import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { ConfigService } from '@nestjs/config';
import * as express from 'express';
import { join } from 'path';
import * as fs from 'fs';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  const configService = app.get(ConfigService);
  
  const apiBaseUrl = configService.get('API_BASE_URL');
  if (!apiBaseUrl) {
    throw new Error('API_BASE_URL environment variable is required');
  }

  // Configure body parser to handle larger requests (for file uploads and auto-save)
  app.use(express.json({ limit: '200mb' }));
  app.use(express.urlencoded({ limit: '200mb', extended: true }));
  
  // Add request logging middleware
  app.use((req, res, next) => {
    console.log(`📥 ${new Date().toISOString()} - ${req.method} ${req.url} from ${req.ip}`);
    console.log(`📥 Origin: ${req.headers.origin || 'none'}`);
    console.log(`📥 User-Agent: ${req.headers['user-agent'] || 'none'}`);
    next();
  });
  
  // Serve static files from public directory with CORS headers
  app.use('/videos', (req, res, next) => {
    // Add CORS headers for video files
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');
    
    // Handle preflight requests
    if (req.method === 'OPTIONS') {
      res.sendStatus(200);
    } else {
      next();
    }
  }, express.static(join(process.cwd(), 'public', 'videos')));
  
  // Serve static image files from public directory with CORS headers
  // Images are organized by opportunity ID: /images/{opportunityId}/{filename}
  app.use('/images', (req, res, next) => {
    // Add CORS headers for image files
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');
    
    // Handle preflight requests
    if (req.method === 'OPTIONS') {
      res.sendStatus(200);
    } else {
      next();
    }
  }, express.static(join(process.cwd(), 'public', 'images')));
  

  
  // Check if we're in development mode
  const isDevelopment = process.env.NODE_ENV === 'development' || !process.env.NODE_ENV || process.env.NODE_ENV === 'local' || process.env.NODE_ENV === 'dev';
  
  console.log('Environment check:', {
    NODE_ENV: process.env.NODE_ENV,
    isDevelopment: isDevelopment,
    apiBaseUrl: apiBaseUrl
  });
  
  if (isDevelopment) {
    // In development, allow all origins for easier testing
    console.log('🔧 Using development CORS configuration (allowing all origins)');
    app.enableCors({
      origin: true, // Allow all origins in development
      credentials: true,
      methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
      allowedHeaders: [
        'Content-Type', 
        'Authorization', 
        'Accept',
        'ngrok-skip-browser-warning',
        'X-Requested-With',
        'Origin',
        'Access-Control-Request-Method',
        'Access-Control-Request-Headers'
      ],
    });
  } else {
    console.log('🔧 Using production CORS configuration');
    // In production, use strict CORS but allow mobile apps
    app.enableCors({
      origin: (origin, callback) => {
        // Allow requests with no origin (like mobile apps, Postman, etc.)
        if (!origin) {
          console.log('🔧 CORS: Allowing request with no origin (mobile app)');
          return callback(null, true);
        }
        
        const allowedOrigins = [
          'http://localhost:8081', // Frontend local development
          'http://localhost:19006', // Expo dev server
          'http://localhost:19000', // Expo dev server
          'http://localhost:19001', // Expo dev server
          'http://localhost:19002', // Expo dev server
          'http://localhost:3000', // Local frontend
          'http://localhost:3001', // Local frontend
          apiBaseUrl, // Dynamic API base URL from environment
        ];
        
        // Check if origin is in allowed list
        if (allowedOrigins.includes(origin)) {
          console.log('🔧 CORS: Allowing origin:', origin);
          return callback(null, true);
        }
        
        // Check regex patterns
        const regexPatterns = [
          /^http:\/\/localhost:\d+$/, // Allow all localhost origins
        ];
        
        for (const pattern of regexPatterns) {
          if (pattern.test(origin)) {
            console.log('🔧 CORS: Allowing origin (regex match):', origin);
            return callback(null, true);
          }
        }
        
        console.log('❌ CORS: Blocking origin:', origin);
        callback(new Error('Not allowed by CORS'));
      },
      credentials: true,
      methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
      allowedHeaders: [
        'Content-Type', 
        'Authorization', 
        'Accept',
        'ngrok-skip-browser-warning',
        'X-Requested-With',
        'Origin',
        'Access-Control-Request-Method',
        'Access-Control-Request-Headers'
      ],
    });
  }
  
  // Start HTTP server on localhost
  const port = parseInt(process.env.PORT ?? '3000');
  
  await app.listen(port, '127.0.0.1');
  console.log(`🚀 HTTP Server running on http://localhost:${port}`);
  console.log(`📱 Frontend should connect to: http://localhost:${port}`);
  console.log(`🌐 Server listening on localhost:${port}`);
  console.log(`🔍 Test with: curl http://localhost:${port}/health`);
}
bootstrap();
