require('dotenv').config();
const { PrismaClient } = require('@prisma/client');
const fs = require('fs');
const path = require('path');

// Parse DATABASE_URL correctly
let databaseUrl = process.env.DATABASE_URL;
if (databaseUrl && databaseUrl.includes('postgresql://')) {
  const match = databaseUrl.match(/postgresql:\/\/[^\s#]+/);
  if (match) {
    databaseUrl = match[0];
  }
}

const prisma = new PrismaClient({
  datasources: {
    db: {
      url: databaseUrl
    }
  }
});

async function restoreUser() {
  console.log('Loading backup file...\n');
  
  // Find the latest backup
  const backupDir = path.join(__dirname, 'backups');
  const files = fs.readdirSync(backupDir)
    .filter(f => f.startsWith('backup_') && f.endsWith('.json'))
    .sort()
    .reverse();
  
  if (files.length === 0) {
    console.error('❌ No backup files found!');
    process.exit(1);
  }
  
  const latestBackup = path.join(backupDir, files[0]);
  console.log(`📁 Using backup: ${files[0]}\n`);
  
  const backup = JSON.parse(fs.readFileSync(latestBackup, 'utf8'));
  
  if (!backup.tables || !backup.tables.users) {
    console.error('❌ No users found in backup!');
    process.exit(1);
  }
  
  const users = backup.tables.users;
  console.log(`Found ${users.length} users in backup:\n`);
  
  // Display all users
  users.forEach((user, index) => {
    console.log(`${index + 1}. ${user.username || user.email} (${user.role}) - ID: ${user.id}`);
    if (user.name) console.log(`   Name: ${user.name}`);
    if (user.email) console.log(`   Email: ${user.email}`);
    console.log('');
  });
  
  // Get command line argument for user ID or username
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log('\nUsage: node restore-user.js <user-id-or-username>');
    console.log('Example: node restore-user.js kemberly.willocks');
    console.log('Example: node restore-user.js 9e86e1c1-e551-42ca-b462-b052d5346441\n');
    process.exit(0);
  }
  
  const searchTerm = args[0].toLowerCase();
  const userToRestore = users.find(u => 
    u.id.toLowerCase() === searchTerm || 
    u.username?.toLowerCase() === searchTerm ||
    u.email?.toLowerCase() === searchTerm
  );
  
  if (!userToRestore) {
    console.error(`❌ User not found: ${args[0]}`);
    process.exit(1);
  }
  
  console.log(`\n🔄 Restoring user: ${userToRestore.username || userToRestore.email}\n`);
  
  try {
    // Check if user already exists
    const existingUser = await prisma.user.findUnique({
      where: { id: userToRestore.id }
    });
    
    if (existingUser) {
      console.log('⚠️  User already exists in database!');
      console.log('Options:');
      console.log('1. Update existing user with backup data');
      console.log('2. Skip (user already exists)');
      
      // For now, we'll update it
      console.log('\n📝 Updating existing user...');
      
      const { id, createdAt, ...updateData } = userToRestore;
      
      const updated = await prisma.user.update({
        where: { id: userToRestore.id },
        data: updateData
      });
      
      console.log('✅ User updated successfully!');
      console.log(`   Username: ${updated.username}`);
      console.log(`   Email: ${updated.email}`);
      console.log(`   Role: ${updated.role}`);
      
    } else {
      console.log('📝 Creating new user...');
      
      const { id, createdAt, ...createData } = userToRestore;
      
      const created = await prisma.user.create({
        data: {
          ...createData,
          id: userToRestore.id, // Keep original ID
          createdAt: new Date(userToRestore.createdAt)
        }
      });
      
      console.log('✅ User restored successfully!');
      console.log(`   Username: ${created.username}`);
      console.log(`   Email: ${created.email}`);
      console.log(`   Role: ${created.role}`);
      console.log(`   ID: ${created.id}`);
    }
    
  } catch (error) {
    console.error('❌ Error restoring user:', error.message);
    if (error.code === 'P2002') {
      console.error('   This user already exists (unique constraint violation)');
    }
    throw error;
  } finally {
    await prisma.$disconnect();
  }
}

restoreUser()
  .catch((error) => {
    console.error('Restore failed:', error);
    process.exit(1);
  });










