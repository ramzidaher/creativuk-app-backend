import { PrismaClient, UserRole, UserStatus } from '@prisma/client';
import * as bcrypt from 'bcrypt';

const prisma = new PrismaClient();

async function createTestUsers() {
  try {
    console.log('Creating test users...');

    const testUsers = [
      {
        username: 'miles.kent',
        email: 'miles.kent@creativuk.co.uk',
        name: 'Miles Kent',
        password: 'Miles@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'hamzah.islam',
        email: 'hamzah.islam@creativuk.co.uk',
        name: 'Hamzah Islam',
        password: 'Hamzah@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'andrew.hughes',
        email: 'andrew.hughes@creativuk.co.uk',
        name: 'Andrew Hughes',
        password: 'Andrew@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'james.barnett',
        email: 'james.barnett@creativuk.co.uk',
        name: 'James Barnett',
        password: 'James@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'ion.zacon',
        email: 'ion.zacon@creativuk.co.uk',
        name: 'Ion Zacon',
        password: 'Ion@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'iuzu.alexandru',
        email: 'iuzu.alexandru@creativuk.co.uk',
        name: 'Iuzu Alexandru',
        password: 'Iuzu@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'support',
        email: 'support@creativuk.co.uk',
        name: 'Support Admin',
        password: 'Support@2025',
        role: UserRole.ADMIN,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'karl.gedney',
        email: 'karl.gedney@creativuk.co.uk',
        name: 'Karl Gedney',
        password: 'Karl@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
      {
        username: 'kemberly.willocks',
        email: 'kemberly.willocks@creativuk.co.uk',
        name: 'Kemberly Willocks',
        password: 'Kemberly@2025',
        role: UserRole.SURVEYOR,
        status: UserStatus.ACTIVE,
      },
    ];

    for (const userData of testUsers) {
      // Check if user already exists
      const existingUser = await prisma.user.findFirst({
        where: {
          OR: [
            { username: userData.username },
            { email: userData.email }
          ]
        }
      });

      if (existingUser) {
        console.log(`User ${userData.username} already exists, skipping...`);
        continue;
      }

      // Hash password
      const hashedPassword = await bcrypt.hash(userData.password, 12);

      // Create user
      const user = await prisma.user.create({
        data: {
          username: userData.username,
          email: userData.email,
          name: userData.name,
          password: hashedPassword,
          role: userData.role,
          status: userData.status,
          isEmailVerified: true,
        }
      });

      console.log(`Created user: ${user.username} (${user.role})`);
    }

    console.log('Test users created successfully!');
  } catch (error) {
    console.error('Error creating test users:', error);
  } finally {
    await prisma.$disconnect();
  }
}

createTestUsers(); 