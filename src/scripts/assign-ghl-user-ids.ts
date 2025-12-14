import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function assignGhlUserIds() {
  try {
    console.log('Assigning GHL user IDs to test users...');
    
    // Get all users from database
    const dbUsers = await prisma.user.findMany({
      select: {
        id: true,
        username: true,
        name: true,
        ghlUserId: true,
        role: true
      }
    });

    console.log(`📋 Found ${dbUsers.length} users in database`);

    // Check which users need GHL user IDs
    const usersWithoutGhlId = dbUsers.filter(user => !user.ghlUserId);
    const usersWithGhlId = dbUsers.filter(user => user.ghlUserId);

    console.log(`✅ Users with GHL ID: ${usersWithGhlId.length}`);
    console.log(`❌ Users without GHL ID: ${usersWithoutGhlId.length}`);

    if (usersWithoutGhlId.length === 0) {
      console.log('🎉 All users already have GHL user IDs assigned!');
      return;
    }

    console.log('\n📝 Users without GHL user IDs:');
    usersWithoutGhlId.forEach(user => {
      console.log(`  - ${user.username} (${user.name || 'No name'})`);
    });

    console.log('\n⚠️  IMPORTANT: You need to manually assign GHL user IDs to these users.');
    console.log('   You can do this by:');
    console.log('   1. Going to your GHL dashboard');
    console.log('   2. Finding the user IDs for each team member');
    console.log('   3. Using the API endpoint: POST /user/assign-ghl-id/:userId');
    console.log('   4. Or running the sync endpoint: GET /user/ghl-sync');

    // Show current GHL user IDs for reference
    if (usersWithGhlId.length > 0) {
      console.log('\n📋 Current GHL user IDs for reference:');
      usersWithGhlId.forEach(user => {
        console.log(`  - ${user.username} (${user.name || 'No name'}): ${user.ghlUserId}`);
      });
    }

    // For testing purposes, you can uncomment and modify this section to assign test IDs
    /*
    const testAssignments = [
      { username: 'admin', ghlUserId: 'p3b0OE6h38soxorDNNi9' },
      { username: 'andrew.hughes', ghlUserId: 'test_user_001' },
      { username: 'ion.zacon', ghlUserId: 'test_user_002' },
      // Add more as needed
    ];

    for (const assignment of testAssignments) {
      const user = dbUsers.find(u => u.username === assignment.username);
      if (user && !user.ghlUserId) {
        await prisma.user.update({
          where: { id: user.id },
          data: { ghlUserId: assignment.ghlUserId }
        });
        console.log(`✅ Assigned GHL user ID to ${assignment.username}: ${assignment.ghlUserId}`);
      }
    }
    */

  } catch (error) {
    console.error('❌ Error assigning GHL user IDs:', error);
  } finally {
    await prisma.$disconnect();
  }
}

assignGhlUserIds();
