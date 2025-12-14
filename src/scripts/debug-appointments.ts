import { NestFactory } from '@nestjs/core';
import { AppModule } from '../app.module';
import { OpportunitiesService } from '../opportunities/opportunities.service';
import { UserService } from '../user/user.service';
import { PrismaService } from '../prisma/prisma.service';

async function debugAppointments() {
  const app = await NestFactory.createApplicationContext(AppModule);
  const opportunitiesService = app.get(OpportunitiesService);
  const userService = app.get(UserService);
  const prismaService = app.get(PrismaService);

  try {
    // Get the first admin user
    const users = await prismaService.user.findMany();
    const adminUser = users.find(user => user.role === 'ADMIN');
    
    if (!adminUser) {
      console.log('No admin user found');
      return;
    }

    console.log(`Testing appointment matching for user: ${adminUser.name}`);
    
    // Test the getOpportunitiesWithAppointments method
    const result = await opportunitiesService.getOpportunitiesWithAppointments(adminUser.id);
    
    console.log('=== OPPORTUNITIES WITH APPOINTMENTS DEBUG ===');
    console.log(`Total opportunities: ${result.total}`);
    console.log(`Classification:`, result.classification);
    
    const withAppointments = result.opportunities.filter((opp: any) => opp.hasAppointment);
    const withoutAppointments = result.opportunities.filter((opp: any) => !opp.hasAppointment);
    
    console.log(`\n=== OPPORTUNITIES WITH APPOINTMENTS (${withAppointments.length}) ===`);
    withAppointments.forEach((opp: any, index: number) => {
      console.log(`${index + 1}. "${opp.name}" (Contact: ${opp.contactName})`);
      console.log(`   Appointment: ${opp.appointmentDetails?.title}`);
      console.log(`   Appointment Contact: ${opp.appointmentDetails?.contact}`);
      console.log(`   Classification: ${opp.classification}`);
      console.log(`   Confidence: ${opp.confidence}`);
      console.log('');
    });
    
    console.log(`\n=== OPPORTUNITIES WITHOUT APPOINTMENTS (${withoutAppointments.length}) ===`);
    withoutAppointments.forEach((opp: any, index: number) => {
      console.log(`${index + 1}. "${opp.name}" (Contact: ${opp.contactName})`);
      console.log(`   Classification: ${opp.classification}`);
      console.log(`   Reason: ${opp.reason}`);
      console.log('');
    });
    
  } catch (error) {
    console.error('Error debugging appointments:', error);
  } finally {
    await app.close();
  }
}

debugAppointments().catch(console.error); 