import { NestFactory } from '@nestjs/core';
import { AppModule } from '../app.module';
import { OpportunitiesService } from '../opportunities/opportunities.service';
import { UserService } from '../user/user.service';

async function performanceTest() {
  const app = await NestFactory.createApplicationContext(AppModule);
  const opportunitiesService = app.get(OpportunitiesService);
  const userService = app.get(UserService);

  try {
    // Get the first admin user for testing
    const adminUsers = await userService.findByRole('ADMIN');
    const adminUser = adminUsers[0];
    
    if (!adminUser) {
      console.log('No admin user found for testing. Please create an admin user in your database.');
      console.log('You can create a test admin user using the create-admin script.');
      return;
    }

    console.log(`🧪 Starting performance test with user: ${adminUser.name}`);

    // Test original method
    console.log('\n📊 Testing original method...');
    const originalStart = Date.now();
    const originalResult = await opportunitiesService.getOpportunitiesWithAppointments(adminUser.id);
    const originalTime = Date.now() - originalStart;

    console.log(`⏱️  Original method took: ${originalTime}ms`);
    console.log(`📈 Original results: ${originalResult.total} opportunities`);

    // Test optimized method
    console.log('\n🚀 Testing optimized method...');
    const optimizedStart = Date.now();
    const optimizedResult = await opportunitiesService.getOpportunitiesWithAppointmentsOptimized(adminUser.id);
    const optimizedTime = Date.now() - optimizedStart;

    console.log(`⏱️  Optimized method took: ${optimizedTime}ms`);
    console.log(`📈 Optimized results: ${optimizedResult.total} opportunities`);

    // Calculate improvement
    const timeSaved = originalTime - optimizedTime;
    const percentageFaster = Math.round((timeSaved / originalTime) * 100);

    console.log('\n📊 PERFORMANCE COMPARISON:');
    console.log(`Original method: ${originalTime}ms`);
    console.log(`Optimized method: ${optimizedTime}ms`);
    console.log(`Time saved: ${timeSaved}ms`);
    console.log(`Percentage faster: ${percentageFaster}%`);

    if (optimizedTime < originalTime) {
      console.log('✅ Optimized method is faster!');
    } else {
      console.log('⚠️  Original method is faster (this might be due to caching)');
    }

    // Log detailed performance metrics
    if (optimizedResult.performance) {
      console.log('\n🔍 OPTIMIZED METHOD DETAILS:');
      console.log(`Total contacts processed: ${optimizedResult.performance.totalContacts}`);
      console.log(`Total appointments retrieved: ${optimizedResult.performance.totalAppointments}`);
    }

  } catch (error) {
    console.error('❌ Performance test failed:', error.message);
  } finally {
    await app.close();
  }
}

// Run the test if this file is executed directly
if (require.main === module) {
  performanceTest().catch(console.error);
}

export { performanceTest };
