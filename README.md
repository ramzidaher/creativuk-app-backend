# CreativUK App Backend

A comprehensive NestJS backend application for managing solar energy opportunities, workflows, and customer interactions.

## 🚀 Features

- **Opportunity Management**: Complete workflow management for solar energy opportunities with multi-step progress tracking
- **Excel Automation**: Automated Excel file processing, calculations, and cell manipulation
- **EPVS Automation**: Specialized automation for EPVS (Energy Performance Verification System) calculations
- **PDF Signing**: Integration with DocuSign, DocuSeal, and Adobe Sign for document signing
- **Email Services**: Automated email sending for surveys, confirmations, and notifications
- **Calendar Integration**: Appointment booking and calendar management
- **Survey Management**: Dynamic survey creation and report generation
- **Video Generation**: Automated video generation from presentations (PowerPoint to MP4)
- **Solar Projections**: Detailed solar projection data and financial analysis
- **Contract Generation**: Automated contract and proposal document generation
- **Pricing Management**: Pricing calculations and management
- **Calculator Progress Tracking**: Real-time progress tracking for calculation workflows
- **Session Management**: User session and workflow state management
- **Auto-Save**: Automatic saving of opportunity data
- **Caching**: Application-level caching for improved performance
- **System Settings**: Configurable system-wide settings
- **Opportunity Outcomes**: Tracking and management of opportunity outcomes
- **Disclaimer Management**: Dynamic disclaimer handling
- **Email Confirmation**: Automated email confirmation workflows
- **Admin Dashboard**: Analytics and opportunity management
- **User Authentication**: JWT-based authentication with role-based access control
- **Third-Party Integrations**:
  - GoHighLevel (CRM) - OAuth integration and user sync
  - OpenSolar (Solar project management) - Public and authenticated APIs
  - DocuSign & DocuSeal (Document signing) - Multiple signing providers
  - Adobe Sign (Document signing)
  - Cloudinary (Media management)
  - OneDrive (File storage)
  - Webhooks support for external integrations

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

   **Required:**
   - `DATABASE_URL`: PostgreSQL connection string (e.g., `postgresql://user:password@localhost:5432/solar_app`)
   - `JWT_SECRET`: Secret key for JWT tokens
   - `API_BASE_URL`: Base URL for your API
   - `PORT`: Server port (default: 3000)

   **GoHighLevel Integration:**
   - `GOHIGHLEVEL_API_TOKEN`: Private Integration token (v2, recommended for internal use)
   - `GHL_LOCATION_ID`: GoHighLevel location (sub-account) ID
   - `GHL_CLIENT_ID`: GoHighLevel OAuth client ID (optional, OAuth flow)
   - `GHL_CLIENT_SECRET`: GoHighLevel OAuth client secret (optional, OAuth flow)
   - `GHL_REDIRECT_URI`: OAuth redirect URI (optional, OAuth flow)
   - `GHL_ACCESS_TOKEN`: Legacy access token (v1, avoid for new setups)

   **Email Configuration:**
   - `SMTP_HOST`: SMTP server host (e.g., `smtp.gmail.com`)
   - `SMTP_PORT`: SMTP server port (e.g., `587`)
   - `SMTP_USER`: SMTP username/email
   - `SMTP_PASS`: SMTP password/app password
   - `SMTP_FROM`: From email address

   **DocuSeal Configuration:**
   - `DOCUSEAL_BASE_URL`: DocuSeal API base URL (e.g., `https://api.docuseal.eu` or `https://api.docuseal.com`)
   - `DOCUSEAL_API_KEY`: DocuSeal API key
   - `DOCUSEAL_USER_EMAIL`: DocuSeal user email
   - `DOCUSEAL_WEBHOOK_SECRET`: Webhook secret for security (optional but recommended)

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

**Core Scripts:**
- `npm run build` - Build the application
- `npm run start` - Start the application
- `npm run start:dev` - Start in development mode with hot reload
- `npm run start:debug` - Start in debug mode
- `npm run start:prod` - Start in production mode

**Code Quality:**
- `npm run lint` - Run ESLint
- `npm run format` - Format code with Prettier
- `npm run typecheck` - Type check without emitting files
- `npm run check` - Run ESLint and type checking

**Testing:**
- `npm test` - Run unit tests
- `npm run test:watch` - Run tests in watch mode
- `npm run test:e2e` - Run end-to-end tests
- `npm run test:cov` - Run tests with coverage
- `npm run test:debug` - Run tests in debug mode

**Utility Scripts:**
- `npm run create:test-admin` - Create a test admin user
- `npm run debug:appointments` - Debug appointment issues
- `npm run performance:test` - Run performance tests
- `npm run performance:compare` - Compare performance metrics

**Excel Testing Scripts:**
- `npm run clear:excel` - Clear Excel cells
- `npm run clear:excel:safe` - Safe Excel cell clearing
- `npm run clear:excel:backup` - Clear Excel with backup
- `npm run test:radio-button` - Test radio button functionality
- `npm run test:customer-file` - Test customer file creation
- `npm run test:complete-calculation` - Test complete calculation flow
- `npm run test:dynamic-signatures` - Test dynamic signature placement

## 🏗️ Project Structure

```
src/
├── admin/                    # Admin dashboard (analytics, opportunity details)
├── appointment/              # Appointment booking functionality
├── auth/                     # Authentication and authorization (JWT, roles)
├── cache/                    # Application caching
├── calculator-progress/      # Calculator progress tracking
├── calendar/                 # Calendar integration
├── cloudinary/               # Cloudinary media management
├── contracts/                # Contract generation and management
├── disclaimer/               # Disclaimer management
├── docusign/                 # DocuSign integration
├── email/                    # Email service
├── email_confirmation/       # Email confirmation workflows
├── epvs-automation/          # EPVS automation
├── excel-automation/         # Excel file automation
├── excel-file-calculator/    # Excel calculation engine
├── health/                   # Health check endpoints
├── integrations/             # Third-party integrations
│   ├── ghl-auth/            # GoHighLevel OAuth
│   ├── opensolar.module.ts  # OpenSolar integration
│   └── webhooks.controller.ts # Webhook handlers
├── onedrive/                 # OneDrive file storage
├── opportunities/            # Opportunity management
│   ├── auto-save.controller.ts
│   ├── opportunity-workflow.controller.ts
│   └── survey.controller.ts
├── opportunity-outcomes/     # Opportunity outcomes tracking
├── pdf-signature/            # PDF signature functionality
├── pdf-signing/              # PDF signing services
├── pricing/                  # Pricing management
├── prisma/                   # Prisma ORM setup
├── session-management/       # Session and workflow state management
├── signatures/               # Signature management (free signatures)
├── solar-projection/         # Solar projection calculations
├── system-settings/          # System-wide settings
├── user/                     # User management
├── video-generation/         # Video generation from presentations
├── app.module.ts            # Root application module
├── app.controller.ts        # Root controller
├── main.ts                  # Application entry point
└── test.controller.ts       # Test endpoints
```

## 🔐 Authentication

The application uses JWT-based authentication with role-based access control (RBAC). Include the token in the Authorization header:

```
Authorization: Bearer <your-jwt-token>
```

**User Roles:**
- `ADMIN` - Full system access
- `SURVEYOR` - Survey and opportunity management
- `USER` - Basic user access

**Protected Routes:**
Most endpoints require authentication. Use the `@UseGuards(JwtAuthGuard)` decorator for authentication and `@Roles()` decorator for role-based access.

## 📚 API Documentation

### Swagger/OpenAPI Documentation

The easiest way to explore and test the API is through the interactive Swagger documentation:

**Access Swagger UI:** `http://localhost:3000/api-docs`

The Swagger UI provides:
- Complete API endpoint documentation
- Request/response schemas
- Interactive API testing
- Authentication support (JWT Bearer tokens)

### API Endpoints

**Authentication:**
- `POST /auth/login` - User login
- `POST /auth/register` - User registration
- `POST /auth/refresh` - Refresh JWT token
- `GET /auth/profile` - Get current user profile
- `PUT /auth/profile` - Update user profile
- `POST /auth/change-password` - Change password
- `POST /auth/reset-password` - Request password reset
- `POST /auth/reset-password/confirm` - Confirm password reset

**Opportunities:**
- `GET /opportunities` - List opportunities
- `POST /opportunities` - Create opportunity
- `GET /opportunities/:id` - Get opportunity details
- `PUT /opportunities/:id` - Update opportunity
- `DELETE /opportunities/:id` - Delete opportunity
- `GET /opportunity-workflow/:id` - Get workflow status
- `POST /opportunity-workflow/:id/complete-step` - Complete workflow step
- `POST /auto-save` - Auto-save opportunity data

**Appointments:**
- `GET /appointments` - List appointments
- `POST /appointments` - Create appointment
- `PUT /appointments/:id` - Update appointment
- `DELETE /appointments/:id` - Delete appointment

**Surveys:**
- `GET /surveys` - List surveys
- `POST /surveys` - Create survey
- `GET /surveys/:id` - Get survey details
- `PUT /surveys/:id` - Update survey

**Excel Automation:**
- `POST /excel-automation/session/calculate` - Calculate Excel session
- `POST /excel-automation/select-radio-button` - Select radio button
- `POST /excel-automation/energy-use/single-rate` - Single rate energy use
- `POST /excel-automation/energy-use/dual-rate` - Dual rate energy use
- `POST /excel-automation/battery/*` - Battery configuration endpoints
- `POST /excel-automation/existing-solar/*` - Existing solar configuration
- `POST /excel-automation/payment/*` - Payment method configuration

**EPVS Automation:**
- `POST /epvs-automation/*` - EPVS automation endpoints

**Calculations:**
- `GET /calculator-progress/:id` - Get calculation progress
- `POST /calculator-progress` - Create calculation progress
- `PUT /calculator-progress/:id` - Update calculation progress

**Presentations:**
- `GET /presentation/:id` - Get presentation
- `POST /presentation` - Create presentation
- `GET /public/presentation/:id` - Public presentation access

**Video Generation:**
- `POST /video/generate` - Generate video from presentation
- `GET /video/:id` - Get video status

**PDF Signing:**
- `POST /pdf-signing/create` - Create PDF signing request
- `POST /pdf-signing/sign` - Sign PDF
- `GET /pdf-signing/:id` - Get signing status
- `POST /docusign/*` - DocuSign endpoints
- `POST /docuseal/*` - DocuSeal endpoints
- `POST /adobe-sign/*` - Adobe Sign endpoints
- `POST /free-signatures/*` - Free signature endpoints
- `POST /digital-signature/*` - Digital signature endpoints

**Contracts:**
- `GET /contracts` - List contracts
- `POST /contracts` - Generate contract
- `GET /contracts/:id` - Get contract
- `POST /contracts/:id/sign` - Sign contract

**Email:**
- `POST /email/send` - Send email
- `POST /email/survey` - Send survey email
- `POST /email-confirmation/*` - Email confirmation endpoints

**Integrations:**
- `GET /creativ-crm/*` - GoHighLevel OAuth endpoints
- `GET /opensolar/*` - OpenSolar endpoints
- `GET /opensolar-public/*` - Public OpenSolar endpoints
- `POST /hooks/*` - Webhook endpoints

**Calendar:**
- `GET /calendar/events` - Get calendar events
- `POST /calendar/events` - Create calendar event

**OneDrive:**
- `GET /onedrive/files` - List OneDrive files
- `POST /onedrive/upload` - Upload file to OneDrive
- `GET /onedrive/files/:id` - Get OneDrive file

**Admin:**
- `GET /admin/analytics` - Get analytics data
- `GET /admin/opportunities` - Admin opportunity management
- `GET /admin/opportunities/:id` - Get opportunity details (admin)

**System:**
- `GET /system-settings` - Get system settings
- `PUT /system-settings` - Update system settings
- `GET /health` - Health check
- `GET /cache/*` - Cache management
- `GET /session/*` - Session management
- `GET /pricing/*` - Pricing endpoints
- `GET /disclaimer/*` - Disclaimer endpoints
- `GET /opportunity-outcomes/*` - Opportunity outcomes
- `GET /solar-projection/*` - Solar projection data

**User Management:**
- `GET /user` - Get current user
- `PUT /user` - Update user
- `GET /user/ghl-sync` - Sync with GoHighLevel

## 🗄️ Database

This project uses Prisma ORM with PostgreSQL. Database migrations are managed through Prisma:

```bash
# Create a new migration
npx prisma migrate dev --name migration_name

# Apply migrations
npx prisma migrate deploy

# Generate Prisma Client
npx prisma generate

# View database in Prisma Studio
npx prisma studio
```

**Key Models:**
- `User` - User accounts with authentication and GHL integration
- `Appointment` - Appointment scheduling
- `OpportunityProgress` - Opportunity workflow tracking
- `OpportunityStep` - Individual workflow steps
- `Survey` - Survey data and reports
- `AutoSave` - Auto-save data
- `CalculatorProgress` - Calculator workflow progress
- `OpportunityOutcome` - Opportunity outcomes tracking

## 🔧 Configuration

Key configuration files:
- `.env` - Environment variables (see Installation section)
- `prisma/schema.prisma` - Database schema
- `nest-cli.json` - NestJS CLI configuration
- `tsconfig.json` - TypeScript configuration
- `tsconfig.build.json` - TypeScript build configuration
- `eslint.config.mjs` - ESLint configuration
- `.prettierrc` - Prettier configuration

## 🧪 Testing

```bash
# Unit tests
npm run test

# Watch mode
npm run test:watch

# E2E tests
npm run test:e2e

# Test coverage
npm run test:cov

# Debug tests
npm run test:debug
```

## 📦 Dependencies

**Core Framework:**
- **@nestjs/core** - NestJS framework
- **@nestjs/common** - NestJS common utilities
- **@nestjs/config** - Configuration management
- **@nestjs/platform-express** - Express platform adapter

**Database:**
- **@prisma/client** - Prisma ORM client
- **prisma** - Prisma CLI and tools

**Authentication:**
- **@nestjs/jwt** - JWT authentication
- **@nestjs/passport** - Passport integration
- **passport** - Authentication middleware
- **passport-jwt** - JWT strategy for Passport
- **jsonwebtoken** - JWT token handling
- **bcrypt** - Password hashing

**HTTP & API:**
- **axios** - HTTP client
- **@nestjs/swagger** - Swagger/OpenAPI documentation

**File Processing:**
- **xlsx** - Excel file processing
- **pdf-lib** - PDF manipulation
- **pdf-parse** - PDF parsing
- **pdfjs-dist** - PDF.js for PDF processing
- **node-signpdf** - PDF signing
- **docxtemplater** - Word document templating
- **pizzip** - ZIP file handling
- **libreoffice-convert** - Document conversion

**Media & Cloud:**
- **cloudinary** - Cloudinary media management
- **canvas** - Canvas rendering
- **puppeteer** - Browser automation (for PDF/video generation)

**Email:**
- **nodemailer** - Email sending

**Document Signing:**
- **docusign-esign** - DocuSign integration
- **@nutrient-sdk/node** - Additional SDK

**Cloud Storage:**
- **@azure/msal-node** - Microsoft Azure authentication (for OneDrive)

**Utilities:**
- **luxon** - Date/time handling
- **class-validator** - Validation decorators
- **class-transformer** - Object transformation
- **multer** - File upload handling
- **form-data** - Form data handling
- **node-fetch** - Fetch API for Node.js
- **node-forge** - Cryptographic utilities
- **dotenv** - Environment variable management

**Development:**
- **typescript** - TypeScript compiler
- **@nestjs/cli** - NestJS CLI
- **ts-node** - TypeScript execution
- **jest** - Testing framework
- **ts-jest** - TypeScript Jest preset
- **supertest** - HTTP assertion library
- **eslint** - Linting
- **prettier** - Code formatting

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





















