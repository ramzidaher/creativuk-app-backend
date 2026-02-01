require('dotenv').config();
const { PrismaClient } = require('@prisma/client');

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

async function listUsers() {
  try {
    const users = await prisma.user.findMany({
      orderBy: { username: 'asc' }
    });
    
    console.log(`\nCurrent users in database (${users.length} total):\n`);
    users.forEach((u, i) => {
      console.log(`${i + 1}. ${u.username}`);
      console.log(`   Name: ${u.name || 'No name'}`);
      console.log(`   Email: ${u.email}`);
      console.log(`   Role: ${u.role}`);
      console.log(`   Status: ${u.status}`);
      console.log(`   ID: ${u.id}`);
      console.log('');
    });
  } catch (error) {
    console.error('Error:', error.message);
  } finally {
    await prisma.$disconnect();
  }
}

listUsers();






