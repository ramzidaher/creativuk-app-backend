import { PrismaClient } from '@prisma/client';
import axios from 'axios';

const prisma = new PrismaClient();

// GHL API configuration - Updated to use the correct v1 API endpoint
const GHL_BASE_URL = 'https://rest.gohighlevel.com/v1';

interface GHLUser {
  id: string;
  firstName: string;
  lastName: string;
  email: string;
}

async function getAllGHLUsers(accessToken: string, locationId: string): Promise<GHLUser[]> {
  try {
    // Updated to use the correct GHL API v1 endpoint
    const url = `${GHL_BASE_URL}/users/`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
      params: {
        locationId: locationId,
        limit: 1000
      },
      timeout: 10000,
    });
    
    console.log(`✅ Successfully fetched ${(response.data as any)?.users?.length || 0} users from GHL`);
    return (response.data as any)?.users || [];
  } catch (error) {
    console.error(`❌ Error fetching GHL users: ${error.message}`);
    if (error.response) {
      console.error(`GHL API error status: ${error.response.status}`);
      console.error(`GHL API error data: ${JSON.stringify(error.response.data)}`);
    }
    return [];
  }
}

async function findGHLUserByName(ghlUsers: GHLUser[], userName: string): Promise<GHLUser | null> {
  // Try to find user by name (case insensitive)
  const user = ghlUsers.find((u: GHLUser) => {
    const fullName = `${u.firstName || ''} ${u.lastName || ''}`.trim();
    return fullName.toLowerCase().includes(userName.toLowerCase()) ||
           u.firstName?.toLowerCase().includes(userName.toLowerCase()) ||
           u.lastName?.toLowerCase().includes(userName.toLowerCase());
  });
  
  if (user) {
    console.log(`✅ Found GHL user: ${user.firstName} ${user.lastName} (ID: ${user.id})`);
    return user;
  } else {
    console.log(`❌ User not found in GHL: ${userName}`);
    return null;
  }
}

async function quickGhlAssign() {
  try {
    console.log('🚀 Quick GHL User ID Assignment with Real API Data...');
    
    // Get GHL credentials from environment - Updated to use correct variable names
    const accessToken = process.env.GOHIGHLEVEL_API_TOKEN;
    const locationId = process.env.GHL_LOCATION_ID; // This might need to be set
    
    console.log('🔧 Environment check:');
    console.log(`  - GHL Base URL: ${GHL_BASE_URL}`);
    console.log(`  - GHL Token: ${accessToken ? '✅ Set' : '❌ Missing'}`);
    console.log(`  - GHL Location ID: ${locationId ? '✅ Set' : '❌ Missing'}`);
    
    if (!accessToken) {
      console.error('❌ GHL API token not configured. Please set GOHIGHLEVEL_API_TOKEN environment variable.');
      return;
    }

    if (!locationId) {
      console.error('❌ GHL Location ID not configured. Please set GHL_LOCATION_ID environment variable.');
      console.log('💡 You can find your Location ID in your GHL dashboard or API settings.');
      return;
    }

    // Fetch all GHL users first
    console.log('📡 Fetching users from GHL API...');
    const ghlUsers = await getAllGHLUsers(accessToken, locationId);
    
    if (ghlUsers.length === 0) {
      console.error('❌ No users found in GHL or API error occurred');
      return;
    }

    // List of usernames to assign GHL IDs to
    const usernamesToAssign = [
      'admin',
      'andrew.hughes',
      'ion.zacon',
      'jordan.stewart',
      'onur.saliah',
      'james.barnett',
      'kenji.omachi',
      'alexandru.iuzu',
    ];

    let assigned = 0;
    let skipped = 0;
    let errors = 0;
    let notFound = 0;

    console.log('\n🔍 Processing users...\n');

    for (const username of usernamesToAssign) {
      try {
        // Find user in database
        const user = await prisma.user.findUnique({
          where: { username: username }
        });

        if (!user) {
          console.log(`❌ User not found in database: ${username}`);
          errors++;
          continue;
        }

        // Check if user already has GHL ID
        if (user.ghlUserId) {
          console.log(`ℹ️  ${username} already has GHL ID: ${user.ghlUserId}`);
          skipped++;
          continue;
        }

        // Try to find matching GHL user by name
        const ghlUser = await findGHLUserByName(ghlUsers, user.name || username);
        
        if (!ghlUser) {
          console.log(`❌ No matching GHL user found for: ${username} (${user.name})`);
          notFound++;
          continue;
        }

        // Assign the GHL user ID
        await prisma.user.update({
          where: { id: user.id },
          data: { ghlUserId: ghlUser.id }
        });

        console.log(`✅ Assigned GHL ID to ${username}: ${ghlUser.id} (${ghlUser.firstName} ${ghlUser.lastName})`);
        assigned++;

      } catch (error) {
        console.error(`❌ Error processing ${username}: ${error.message}`);
        errors++;
      }
    }

    console.log('\n📊 Summary:');
    console.log(`  ✅ Assigned: ${assigned}`);
    console.log(`  ℹ️  Skipped (already assigned): ${skipped}`);
    console.log(`  ❌ Not found in GHL: ${notFound}`);
    console.log(`  ❌ Errors: ${errors}`);

    if (assigned > 0) {
      console.log('\n🎉 Quick assignment completed!');
      console.log('\n💡 Available GHL users for reference:');
      ghlUsers.forEach(user => {
        console.log(`   - ${user.firstName} ${user.lastName} (ID: ${user.id})`);
      });
    } else {
      console.log('\n⚠️  No new assignments made. Check the summary above for details.');
    }

  } catch (error) {
    console.error('❌ Script error:', error);
  } finally {
    await prisma.$disconnect();
  }
}

quickGhlAssign();
