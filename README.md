# CreativUK App Backend

A comprehensive NestJS backend application for managing solar energy opportunities, workflows, and customer interactions.

## 🚀 Features

- **Opportunity Management**: Complete workflow management for solar energy opportunities
- **Excel Automation**: Automated Excel file processing and calculations
- **PDF Signing**: Integration with DocuSign and DocuSeal for document signing
- **Email Services**: Automated email sending for surveys, confirmations, and notifications
- **Calendar Integration**: Appointment booking and calendar management
- **Survey Management**: Dynamic survey creation and report generation
- **Video Generation**: Automated video generation from presentations
- **Third-Party Integrations**:
  - GoHighLevel (CRM)
  - OpenSolar (Solar project management)
  - DocuSign & DocuSeal (Document signing)
  - Cloudinary (Media management)
  - OneDrive (File storage)
- **Admin Dashboard**: Analytics and opportunity management
- **User Authentication**: JWT-based authentication with role-based access control
- **Auto-Save**: Automatic saving of opportunity data

## 📋 Prerequisites

- Node.js (v18 or higher)
- PostgreSQL database
- npm or yarn package manager

## 🛠️ Installation

1. Clone the repository:
```bash
git clone https://github.com/ramzidaher/creativuk-app-backend.git
cd creativuk-app-backend
```

2. Install dependencies:
```bash
npm install
```

3. Set up environment variables:
```bash
cp .env.example .env
```

4. Configure your `.env` file with the following variables:
   - `DATABASE_URL`: PostgreSQL connection string
   - `JWT_SECRET`: Secret key for JWT tokens
   - `API_BASE_URL`: Base URL for your API
   - `PORT`: Server port (default: 3000)
   - GoHighLevel credentials (if using GHL integration)
   - SMTP credentials (for email services)
   - DocuSeal/DocuSign API keys (if using document signing)
   - Cloudinary credentials (if using media storage)

5. Set up the database:
```bash
npx prisma migrate dev
npx prisma generate
```

## 🏃 Running the Application

### Development
```bash
npm run start:dev
```

### Production
```bash
npm run build
npm run start:prod
```

### Debug Mode
```bash
npm run start:debug
```

## 📝 Available Scripts

- `npm run build` - Build the application
- `npm run start` - Start the application
- `npm run start:dev` - Start in development mode with hot reload
- `npm run start:debug` - Start in debug mode
- `npm run start:prod` - Start in production mode
- `npm run lint` - Run ESLint
- `npm run format` - Format code with Prettier
- `npm test` - Run unit tests
- `npm run test:e2e` - Run end-to-end tests
- `npm run test:cov` - Run tests with coverage

## 🏗️ Project Structure

```
src/
├── admin/              # Admin dashboard controllers and services
├── appointment/        # Appointment booking functionality
├── auth/              # Authentication and authorization
├── calendar/          # Calendar integration
├── contracts/         # Contract generation and management
├── email/             # Email service
├── excel-automation/  # Excel file automation
├── excel-file-calculator/ # Excel calculation engine
├── integrations/      # Third-party integrations (GHL, OpenSolar, etc.)
├── opportunities/     # Opportunity management
├── pdf-signing/       # PDF signing services
├── signatures/        # Signature management
├── user/              # User management
└── main.ts           # Application entry point
```

## 🔐 Authentication

The application uses JWT-based authentication. Include the token in the Authorization header:

```
Authorization: Bearer <your-jwt-token>
```

## 📚 API Documentation

The API base URL is configured via the `API_BASE_URL` environment variable. Key endpoints include:

- `/api/auth/*` - Authentication endpoints
- `/api/opportunities/*` - Opportunity management
- `/api/appointments/*` - Appointment booking
- `/api/admin/*` - Admin dashboard endpoints
- `/api/integrations/*` - Third-party integrations
- `/health` - Health check endpoint

## 🗄️ Database

This project uses Prisma ORM with PostgreSQL. Database migrations are managed through Prisma:

```bash
# Create a new migration
npx prisma migrate dev --name migration_name

# Apply migrations
npx prisma migrate deploy

# Generate Prisma Client
npx prisma generate
```

## 🔧 Configuration

Key configuration files:
- `.env` - Environment variables
- `prisma/schema.prisma` - Database schema
- `nest-cli.json` - NestJS CLI configuration
- `tsconfig.json` - TypeScript configuration

## 🧪 Testing

```bash
# Unit tests
npm run test

# E2E tests
npm run test:e2e

# Test coverage
npm run test:cov
```

## 📦 Dependencies

Key dependencies:
- **@nestjs/core** - NestJS framework
- **@prisma/client** - Prisma ORM
- **@nestjs/jwt** - JWT authentication
- **passport** - Authentication middleware
- **axios** - HTTP client
- **puppeteer** - Browser automation
- **pdf-lib** - PDF manipulation
- **xlsx** - Excel file processing

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is private and proprietary.

## 👥 Authors

- CreativUK Development Team

## 🆘 Support

For support, please contact the development team or open an issue in the repository.

---

Built with ❤️ using NestJS
