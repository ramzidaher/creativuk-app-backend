import { PrismaClient } from '@prisma/client';
import * as bcrypt from 'bcrypt';

const prisma = new PrismaClient();

async function updateExistingUserPasswords() {
  try {
    console.log('Updating existing user passwords...');

    // Define the users that need password updates
    const usersToUpdate = [
      {
        username: 'miles.kent',
        email: 'miles.kent@creativuk.co.uk',
        newPassword: 'Miles@2025',
      },
      {
        username: 'andrew.hughes',
        email: 'andrew.hughes@creativuk.co.uk',
        newPassword: 'Andrew@2025',
      },
      {
        username: 'james.barnett',
        email: 'james.barnett@creativuk.co.uk',
        newPassword: 'James@2025',
      },
      {
        username: 'ion.zacon',
        email: 'ion.zacon@creativuk.co.uk',
        newPassword: 'Ion@2025',
      },
      {
        username: 'karl.gedney',
        email: 'karl.gedney@creativuk.co.uk',
        newPassword: 'Karl@2025',
      },
      {
        username: 'onur.saliah',
        email: 'onur.saliah@creativuk.co.uk',
        newPassword: 'Onur@2025',
      },
      {
        username: 'kenji.omachi',
        email: 'kenji.omachi@creativuk.co.uk',
        newPassword: 'Kenji@2025',
      },
      {
        username: 'jordan.stewart',
        email: 'jordan.stewart@creativuk.co.uk',
        newPassword: 'Jordan@2025',
      },
      {
        username: 'alexandru.iuzu',
        email: 'alexandru.iuzu@creativuk.co.uk',
        newPassword: 'Alexandru@2025',
      },
    ];

    for (const userData of usersToUpdate) {
      // Find the existing user
      const existingUser = await prisma.user.findFirst({
        where: {
          OR: [
            { username: userData.username },
            { email: userData.email }
          ]
        }
      });

      if (existingUser) {
        // Hash the new password
        const hashedPassword = await bcrypt.hash(userData.newPassword, 12);

        // Update the user's password
        await prisma.user.update({
          where: { id: existingUser.id },
          data: { password: hashedPassword }
        });

        console.log(`Updated password for user: ${existingUser.username} (${existingUser.email})`);
      } else {
        console.log(`User not found: ${userData.username} (${userData.email})`);
      }
    }

    console.log('Password updates completed successfully!');
  } catch (error) {
    console.error('Error updating user passwords:', error);
  } finally {
    await prisma.$disconnect();
  }
}

updateExistingUserPasswords();
