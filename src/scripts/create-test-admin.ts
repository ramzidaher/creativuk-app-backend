import { NestFactory } from '@nestjs/core';
import { AppModule } from '../app.module';
import { UserService } from '../user/user.service';
import * as bcrypt from 'bcrypt';

async function createTestAdmin() {
  const app = await NestFactory.createApplicationContext(AppModule);
  const userService = app.get(UserService);

  try {
    // Check if test admin already exists
    const existingAdmins = await userService.findByRole('ADMIN');
    if (existingAdmins.length > 0) {
      console.log('✅ Admin user already exists:', existingAdmins[0].username);
      return;
    }

    // Create test admin user
    const hashedPassword = await bcrypt.hash('testadmin123', 10);
    
    const testAdmin = await userService.upsertByGhlUserId({
      ghlUserId: 'test-admin-ghl',
      name: 'Test Admin',
      email: 'testadmin@creativsolar.com',
      ghlAccessToken: 'test-token',
      ghlRefreshToken: 'test-refresh-token',
    });

    // Update with admin role and proper credentials
    await userService.update(testAdmin.id, {
      role: 'ADMIN',
      username: 'testadmin',
      password: hashedPassword,
      email: 'testadmin@creativsolar.com',
    });

    console.log('✅ Test admin user created successfully');
    console.log('Username: testadmin');
    console.log('Password: testadmin123');
    console.log('Role: ADMIN');

  } catch (error) {
    console.error('❌ Error creating test admin:', error.message);
  } finally {
    await app.close();
  }
}

// Run the script if this file is executed directly
if (require.main === module) {
  createTestAdmin().catch(console.error);
}

export { createTestAdmin };
