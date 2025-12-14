import { NestFactory } from '@nestjs/core';
import { AppModule } from '../app.module';
import { OpportunitiesService } from '../opportunities/opportunities.service';
import { UserService } from '../user/user.service';

async function performanceComparison() {
  const app = await NestFactory.createApplicationContext(AppModule);
  const opportunitiesService = app.get(OpportunitiesService);
  const userService = app.get(UserService);

  try {
    // Get the first admin user for testing
    const adminUsers = await userService.findByRole('ADMIN');
    const adminUser = adminUsers[0];
    
    if (!adminUser) {
      console.log('No admin user found for testing. Please create an admin user first.');
      console.log('Run: npm run create:test-admin');
      return;
    }

    console.log(`🧪 Starting comprehensive performance comparison with user: ${adminUser.name}`);
    console.log('='.repeat(60));

    // Test 1: Original method
    console.log('\n📊 Testing ORIGINAL method...');
    const originalStart = Date.now();
    const originalResult = await opportunitiesService.getOpportunitiesWithAppointments(adminUser.id);
    const originalTime = Date.now() - originalStart;

    console.log(`⏱️  Original method: ${originalTime}ms`);
    console.log(`📈 Results: ${originalResult.total} opportunities`);

    // Test 2: Optimized method
    console.log('\n🚀 Testing OPTIMIZED method...');
    const optimizedStart = Date.now();
    const optimizedResult = await opportunitiesService.getOpportunitiesWithAppointmentsOptimized(adminUser.id);
    const optimizedTime = Date.now() - optimizedStart;

    console.log(`⏱️  Optimized method: ${optimizedTime}ms`);
    console.log(`📈 Results: ${optimizedResult.total} opportunities`);

    // Test 3: Hybrid method (NEW - uses dashboard data)
    console.log('\n⚡ Testing HYBRID method (dashboard + appointments)...');
    const hybridStart = Date.now();
    const hybridResult = await opportunitiesService.getOpportunitiesWithAppointmentsHybrid(adminUser.id);
    const hybridTime = Date.now() - hybridStart;

    console.log(`⏱️  Hybrid method: ${hybridTime}ms`);
    console.log(`📈 Results: ${hybridResult.total} opportunities`);

    // Calculate improvements
    const originalVsOptimized = originalTime - optimizedTime;
    const originalVsHybrid = originalTime - hybridTime;
    const optimizedVsHybrid = optimizedTime - hybridTime;

    const originalVsOptimizedPercent = Math.round((originalVsOptimized / originalTime) * 100);
    const originalVsHybridPercent = Math.round((originalVsHybrid / originalTime) * 100);
    const optimizedVsHybridPercent = Math.round((optimizedVsHybrid / optimizedTime) * 100);

    console.log('\n' + '='.repeat(60));
    console.log('📊 COMPREHENSIVE PERFORMANCE COMPARISON');
    console.log('='.repeat(60));
    console.log(`Original method:     ${originalTime}ms`);
    console.log(`Optimized method:    ${optimizedTime}ms`);
    console.log(`Hybrid method:       ${hybridTime}ms`);
    console.log('');
    console.log('IMPROVEMENTS:');
    console.log(`Original → Optimized: ${originalVsOptimized}ms faster (${originalVsOptimizedPercent}%)`);
    console.log(`Original → Hybrid:    ${originalVsHybrid}ms faster (${originalVsHybridPercent}%)`);
    console.log(`Optimized → Hybrid:   ${optimizedVsHybrid}ms faster (${optimizedVsHybridPercent}%)`);
    console.log('');

    // Determine the best method
    const methods = [
      { name: 'Original', time: originalTime },
      { name: 'Optimized', time: optimizedTime },
      { name: 'Hybrid', time: hybridTime }
    ];
    
    const fastest = methods.reduce((prev, current) => 
      prev.time < current.time ? prev : current
    );

    console.log(`🏆 FASTEST METHOD: ${fastest.name} (${fastest.time}ms)`);
    console.log('');

    // Detailed metrics
    if (hybridResult.performance) {
      console.log('🔍 HYBRID METHOD DETAILS:');
      console.log(`Method: ${hybridResult.performance.method}`);
      console.log(`Dashboard data time: ${hybridResult.performance.dashboardDataTime}`);
      console.log(`Appointment processing: ${hybridResult.performance.appointmentProcessingTime}`);
      console.log(`Total contacts: ${hybridResult.performance.totalContacts}`);
      console.log(`Total appointments: ${hybridResult.performance.totalAppointments}`);
    }

    console.log('\n💡 RECOMMENDATION:');
    if (hybridTime < optimizedTime && hybridTime < originalTime) {
      console.log('✅ Use HYBRID method - it combines fast dashboard data with appointment info');
      console.log('   This should reduce loading time from 35s to ~8-12s');
    } else if (optimizedTime < originalTime) {
      console.log('✅ Use OPTIMIZED method - it batches API calls efficiently');
      console.log('   This should reduce loading time from 35s to ~15-20s');
    } else {
      console.log('⚠️  Use ORIGINAL method - other methods may have issues');
    }

  } catch (error) {
    console.error('❌ Performance comparison failed:', error.message);
  } finally {
    await app.close();
  }
}

// Run the test if this file is executed directly
if (require.main === module) {
  performanceComparison().catch(console.error);
}

export { performanceComparison };
