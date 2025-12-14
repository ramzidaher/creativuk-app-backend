import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function fixGhlIds() {
  try {
    console.log('🔍 Checking database for users with null/default GHL user IDs...');
    
    // Get all users
    const users = await prisma.user.findMany({
      select: {
        id: true,
        username: true,
        name: true,
        ghlUserId: true,
        role: true
      }
    });

    console.log(`📋 Found ${users.length} total users in database`);

    // Check for users with null, undefined, or 'default' ghlUserId
    const usersWithInvalidGhlId = users.filter(user => 
      !user.ghlUserId || 
      user.ghlUserId === 'default' || 
      user.ghlUserId === 'null'
    );

    const usersWithValidGhlId = users.filter(user => 
      user.ghlUserId && 
      user.ghlUserId !== 'default' && 
      user.ghlUserId !== 'null'
    );

    console.log(`✅ Users with valid GHL ID: ${usersWithValidGhlId.length}`);
    console.log(`❌ Users with invalid/null GHL ID: ${usersWithInvalidGhlId.length}`);

    if (usersWithInvalidGhlId.length > 0) {
      console.log('\n📝 Users that need GHL user IDs:');
      usersWithInvalidGhlId.forEach(user => {
        console.log(`  - ${user.username} (${user.name || 'No name'}) - Current: ${user.ghlUserId || 'null'}`);
      });

      console.log('\n🔧 To fix this, you need to:');
      console.log('1. Get real GHL user IDs from your GHL dashboard');
      console.log('2. Update the users in the database');
      console.log('3. Or use the API endpoint: GET /user/ghl-sync');
      
      // For immediate testing, let's assign some test IDs
      console.log('\n🚀 Assigning test GHL user IDs for immediate testing...');
      
      const testAssignments = [
        { username: 'admin', ghlUserId: 'p3b0OE6h38soxorDNNi9' },
        { username: 'andrew.hughes', ghlUserId: 'test_user_001' },
        { username: 'ion.zacon', ghlUserId: 'test_user_002' },
        { username: 'jordan.stewart', ghlUserId: 'test_user_003' },
        { username: 'onur.saliah', ghlUserId: 'test_user_004' },
        { username: 'james.barnett', ghlUserId: 'test_user_005' },
        { username: 'kenji.omachi', ghlUserId: 'test_user_006' },
        { username: 'alexandru.iuzu', ghlUserId: 'test_user_007' },
      ];

      let assigned = 0;
      for (const assignment of testAssignments) {
        const user = usersWithInvalidGhlId.find(u => u.username === assignment.username);
        if (user) {
          try {
            await prisma.user.update({
              where: { id: user.id },
              data: { ghlUserId: assignment.ghlUserId }
            });
            console.log(`✅ Assigned GHL ID to ${assignment.username}: ${assignment.ghlUserId}`);
            assigned++;
          } catch (error) {
            console.error(`❌ Error assigning to ${assignment.username}: ${error.message}`);
          }
        }
      }

      if (assigned > 0) {
        console.log(`\n🎉 Successfully assigned ${assigned} GHL user IDs!`);
        console.log('⚠️  Note: These are test IDs. Replace with real GHL user IDs for production.');
      }
    } else {
      console.log('🎉 All users already have valid GHL user IDs!');
    }

    // Show final status
    console.log('\n📊 Final Status:');
    const finalUsers = await prisma.user.findMany({
      select: {
        username: true,
        name: true,
        ghlUserId: true
      }
    });

    finalUsers.forEach(user => {
      const status = user.ghlUserId ? '✅' : '❌';
      console.log(`${status} ${user.username} (${user.name || 'No name'}): ${user.ghlUserId || 'NO GHL ID'}`);
    });

  } catch (error) {
    console.error('❌ Error fixing GHL user IDs:', error);
  } finally {
    await prisma.$disconnect();
  }
}

fixGhlIds();
