require('dotenv').config();
const { PrismaClient } = require('@prisma/client');
const fs = require('fs');
const path = require('path');

// Parse DATABASE_URL correctly (handle cases where .env has multiple entries)
let databaseUrl = process.env.DATABASE_URL;
if (databaseUrl && databaseUrl.includes('postgresql://')) {
  // Extract the first valid postgresql:// URL
  const match = databaseUrl.match(/postgresql:\/\/[^\s#]+/);
  if (match) {
    databaseUrl = match[0];
  }
}

// Ensure DATABASE_URL is set correctly
if (!databaseUrl || !databaseUrl.startsWith('postgresql://')) {
  console.error('❌ DATABASE_URL is not set correctly in .env file');
  console.error('Current value:', process.env.DATABASE_URL);
  process.exit(1);
}

console.log('Connecting to database...');

const prisma = new PrismaClient({
  datasources: {
    db: {
      url: databaseUrl
    }
  }
});

async function backupDatabase() {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const backupDir = path.join(__dirname, 'backups');
  
  // Create backups directory if it doesn't exist
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }

  const backupFile = path.join(backupDir, `backup_${timestamp}.json`);
  
  console.log('Starting database backup...');
  console.log(`Backup file: ${backupFile}`);

  try {
    // Get all data from each table
    const backup = {
      timestamp: new Date().toISOString(),
      database: 'solar_app',
      tables: {}
    };

    // Backup Users
    console.log('Backing up Users...');
    backup.tables.users = await prisma.user.findMany();
    console.log(`  - ${backup.tables.users.length} users found`);

    // Backup Appointments
    console.log('Backing up Appointments...');
    backup.tables.appointments = await prisma.appointment.findMany();
    console.log(`  - ${backup.tables.appointments.length} appointments found`);

    // Backup Surveys
    console.log('Backing up Surveys...');
    backup.tables.surveys = await prisma.survey.findMany();
    console.log(`  - ${backup.tables.surveys.length} surveys found`);

    // Backup OpportunityProgress
    console.log('Backing up OpportunityProgress...');
    backup.tables.opportunityProgress = await prisma.opportunityProgress.findMany();
    console.log(`  - ${backup.tables.opportunityProgress.length} opportunity progress records found`);

    // Backup OpportunityStep
    console.log('Backing up OpportunityStep...');
    backup.tables.opportunitySteps = await prisma.opportunityStep.findMany();
    console.log(`  - ${backup.tables.opportunitySteps.length} opportunity steps found`);

    // Backup OpportunityOutcome
    console.log('Backing up OpportunityOutcome...');
    backup.tables.opportunityOutcomes = await prisma.opportunityOutcome.findMany();
    console.log(`  - ${backup.tables.opportunityOutcomes.length} opportunity outcomes found`);

    // Backup AutoSave
    console.log('Backing up AutoSave...');
    backup.tables.autoSaves = await prisma.autoSave.findMany();
    console.log(`  - ${backup.tables.autoSaves.length} auto saves found`);

    // Backup CalculatorProgress
    console.log('Backing up CalculatorProgress...');
    backup.tables.calculatorProgress = await prisma.calculatorProgress.findMany();
    console.log(`  - ${backup.tables.calculatorProgress.length} calculator progress records found`);

    // Backup SurveyImage
    console.log('Backing up SurveyImage...');
    backup.tables.surveyImages = await prisma.surveyImage.findMany();
    console.log(`  - ${backup.tables.surveyImages.length} survey images found`);

    // Backup Signature
    console.log('Backing up Signature...');
    backup.tables.signatures = await prisma.signature.findMany();
    console.log(`  - ${backup.tables.signatures.length} signatures found`);

    // Backup OpenSolarProject
    console.log('Backing up OpenSolarProject...');
    backup.tables.openSolarProjects = await prisma.openSolarProject.findMany();
    console.log(`  - ${backup.tables.openSolarProjects.length} open solar projects found`);

    // Backup DocuSealSubmission
    console.log('Backing up DocuSealSubmission...');
    backup.tables.docuSealSubmissions = await prisma.docuSealSubmission.findMany();
    console.log(`  - ${backup.tables.docuSealSubmissions.length} docu seal submissions found`);

    // Backup SystemSettings
    console.log('Backing up SystemSettings...');
    backup.tables.systemSettings = await prisma.systemSettings.findMany();
    console.log(`  - ${backup.tables.systemSettings.length} system settings found`);

    // Try to backup Form tables (they might not exist yet)
    try {
      console.log('Backing up Forms...');
      backup.tables.forms = await prisma.form.findMany();
      console.log(`  - ${backup.tables.forms.length} forms found`);
    } catch (e) {
      console.log(`  - Forms table doesn't exist yet: ${e.message}`);
      backup.tables.forms = [];
    }

    try {
      console.log('Backing up FormFields...');
      backup.tables.formFields = await prisma.formField.findMany();
      console.log(`  - ${backup.tables.formFields.length} form fields found`);
    } catch (e) {
      console.log(`  - FormFields table doesn't exist yet: ${e.message}`);
      backup.tables.formFields = [];
    }

    try {
      console.log('Backing up FormSubmissions...');
      backup.tables.formSubmissions = await prisma.formSubmission.findMany();
      console.log(`  - ${backup.tables.formSubmissions.length} form submissions found`);
    } catch (e) {
      console.log(`  - FormSubmissions table doesn't exist yet: ${e.message}`);
      backup.tables.formSubmissions = [];
    }

    try {
      console.log('Backing up FormSubmissionFields...');
      backup.tables.formSubmissionFields = await prisma.formSubmissionField.findMany();
      console.log(`  - ${backup.tables.formSubmissionFields.length} form submission fields found`);
    } catch (e) {
      console.log(`  - FormSubmissionFields table doesn't exist yet: ${e.message}`);
      backup.tables.formSubmissionFields = [];
    }

    try {
      console.log('Backing up FormSubmissionImages...');
      backup.tables.formSubmissionImages = await prisma.formSubmissionImage.findMany();
      console.log(`  - ${backup.tables.formSubmissionImages.length} form submission images found`);
    } catch (e) {
      console.log(`  - FormSubmissionImages table doesn't exist yet: ${e.message}`);
      backup.tables.formSubmissionImages = [];
    }

    // Write backup to file
    fs.writeFileSync(backupFile, JSON.stringify(backup, null, 2));
    
    console.log('\n✅ Backup completed successfully!');
    console.log(`📁 Backup saved to: ${backupFile}`);
    
    // Calculate total records
    const totalRecords = Object.values(backup.tables).reduce((sum, records) => sum + records.length, 0);
    console.log(`📊 Total records backed up: ${totalRecords}`);

  } catch (error) {
    console.error('❌ Error during backup:', error);
    throw error;
  } finally {
    await prisma.$disconnect();
  }
}

backupDatabase()
  .catch((error) => {
    console.error('Backup failed:', error);
    process.exit(1);
  });

