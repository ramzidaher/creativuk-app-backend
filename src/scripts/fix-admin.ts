import { PrismaClient } from '@prisma/client';
import * as bcrypt from 'bcrypt';

const prisma = new PrismaClient();

async function fixAdminUser() {
  try {
    // Find the existing admin user
    const existingAdmin = await prisma.user.findFirst({
      where: {
        OR: [
          { username: 'admin' },
          { email: 'admin@creativsolar.com' }
        ]
      }
    });

    if (!existingAdmin) {
      console.log('No admin user found. Creating new admin user...');
      
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
        status: adminUser.status,
        ghlUserId: adminUser.ghlUserId
      });
    } else {
      console.log('Found existing admin user:', {
        id: existingAdmin.id,
        username: existingAdmin.username,
        email: existingAdmin.email,
        role: existingAdmin.role,
        status: existingAdmin.status,
        ghlUserId: existingAdmin.ghlUserId
      });

      // Check if there's a conflict with the GHL ID
      const conflictingUser = await prisma.user.findFirst({
        where: {
          ghlUserId: 'p3b0OE6h38soxorDNNi9',
          id: { not: existingAdmin.id }
        }
      });

      if (conflictingUser) {
        console.log('Found conflicting user with same GHL ID:', {
          id: conflictingUser.id,
          username: conflictingUser.username,
          name: conflictingUser.name
        });
        
        // Remove GHL ID from conflicting user
        await prisma.user.update({
          where: { id: conflictingUser.id },
          data: { ghlUserId: null }
        });
        console.log('Removed GHL ID from conflicting user');
      }

      // Update admin user to be active and set GHL ID
      const updatedAdmin = await prisma.user.update({
        where: { id: existingAdmin.id },
        data: {
          status: 'ACTIVE',
          ghlUserId: 'p3b0OE6h38soxorDNNi9',
          // Reset password to admin123
          password: await bcrypt.hash('admin123', 10)
        }
      });

      console.log('Admin user updated successfully:', {
        id: updatedAdmin.id,
        username: updatedAdmin.username,
        email: updatedAdmin.email,
        role: updatedAdmin.role,
        status: updatedAdmin.status,
        ghlUserId: updatedAdmin.ghlUserId
      });
    }
  } catch (error) {
    console.error('Error fixing admin user:', error);
  } finally {
    await prisma.$disconnect();
  }
}

fixAdminUser();

