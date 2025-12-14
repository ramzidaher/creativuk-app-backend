import { PrismaClient } from '@prisma/client';
import * as bcrypt from 'bcrypt';

const prisma = new PrismaClient();

async function createAdminUser() {
  try {
    // Check if admin already exists
    const existingAdmin = await prisma.user.findFirst({
      where: {
        OR: [
          { username: 'admin' },
          { email: 'admin@creativsolar.com' }
        ]
      }
    });

    if (existingAdmin) {
      console.log('Admin user already exists');
      
      // Update admin user with ghlUserId for testing
      if (!existingAdmin.ghlUserId) {
        await prisma.user.update({
          where: { id: existingAdmin.id },
          data: { ghlUserId: 'p3b0OE6h38soxorDNNi9' } // Use the GHL user ID from the opportunities
        });
        console.log('Updated admin user with ghlUserId: p3b0OE6h38soxorDNNi9');
      }
      
      return;
    }

    // Hash password
    const hashedPassword = await bcrypt.hash('admin123', 10);

    // Create admin user
    const adminUser = await prisma.user.create({
      data: {
        username: 'admin',
        email: 'admin@creativsolar.com',
        password: hashedPassword,
        name: 'System Administrator',
        role: 'ADMIN',
        status: 'ACTIVE',
        isEmailVerified: true,
        ghlUserId: 'p3b0OE6h38soxorDNNi9', // Set ghlUserId for workflow testing
      }
    });

    console.log('Admin user created successfully:', {
      id: adminUser.id,
      username: adminUser.username,
      email: adminUser.email,
      role: adminUser.role,
      ghlUserId: adminUser.ghlUserId
    });
  } catch (error) {
    console.error('Error creating admin user:', error);
  } finally {
    await prisma.$disconnect();
  }
}

createAdminUser(); 