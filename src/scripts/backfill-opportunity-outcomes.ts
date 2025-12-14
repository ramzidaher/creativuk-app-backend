import { PrismaClient } from '@prisma/client';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { ConfigService } from '@nestjs/config';
import { ConfigModule } from '@nestjs/config';

const prisma = new PrismaClient();

async function backfillOpportunityOutcomes() {
  console.log('🔄 Starting opportunity outcomes backfill...');

  try {
    // Initialize services
    const configService = new ConfigService();
    const ghlService = new GoHighLevelService();
    
    // Create a simple opportunity outcomes service for the script
    const opportunityOutcomesService = {
      recordOutcome: async (data: any) => {
        // Check if outcome already exists
        const existingOutcome = await prisma.opportunityOutcome.findUnique({
          where: { ghlOpportunityId: data.ghlOpportunityId },
        });

        if (existingOutcome) {
          // Update existing outcome
          return await prisma.opportunityOutcome.update({
            where: { ghlOpportunityId: data.ghlOpportunityId },
            data: {
              outcome: data.outcome,
              value: data.value,
              notes: data.notes,
              stageAtOutcome: data.stageAtOutcome,
              ghlUpdatedAt: new Date(),
              updatedAt: new Date(),
            },
          });
        } else {
          // Create new outcome
          return await prisma.opportunityOutcome.create({
            data: {
              ghlOpportunityId: data.ghlOpportunityId,
              userId: data.userId,
              outcome: data.outcome,
              value: data.value,
              notes: data.notes,
              stageAtOutcome: data.stageAtOutcome,
              ghlUpdatedAt: new Date(),
            },
          });
        }
      }
    };

    // Get GHL credentials
    const accessToken = configService.get<string>('GOHIGHLEVEL_API_TOKEN');
    const locationId = configService.get<string>('GOHIGHLEVEL_LOCATION_ID');
    
    if (!accessToken || !locationId) {
      throw new Error('GHL credentials not configured');
    }
    
    const credentials = { accessToken, locationId };

    console.log('📊 Fetching opportunities from GHL...');
    
    // Get all opportunities from GHL
    const opportunities = await ghlService.getOpportunities(
      credentials.accessToken,
      credentials.locationId
    );

    console.log(`📈 Found ${opportunities.length} opportunities in GHL`);

    // Get all users to map opportunities to users
    const users = await prisma.user.findMany({
      where: { 
        status: 'ACTIVE',
        ghlUserId: { not: null }
      }
    });

    console.log(`👥 Found ${users.length} active users with GHL IDs`);

    let processed = 0;
    let created = 0;
    let skipped = 0;
    let errors = 0;

    // Process each opportunity
    for (const opportunity of opportunities) {
      try {
        processed++;
        
        // Determine outcome from GHL data
        const outcome = determineOutcomeFromGHL(opportunity);
        
        if (!outcome || outcome === 'IN_PROGRESS') {
          skipped++;
          continue;
        }

        // Find the user who should own this opportunity
        // For now, we'll assign to the first admin user if no specific assignment logic exists
        const assignedUser = users.find(user => user.role === 'ADMIN') || users[0];
        
        if (!assignedUser) {
          console.warn(`⚠️  No user found for opportunity ${opportunity.id}`);
          skipped++;
          continue;
        }

        // Check if outcome already exists
        const existingOutcome = await prisma.opportunityOutcome.findUnique({
          where: { ghlOpportunityId: opportunity.id }
        });

        if (existingOutcome) {
          skipped++;
          continue;
        }

        // Create the outcome record
        await opportunityOutcomesService.recordOutcome({
          ghlOpportunityId: opportunity.id,
          userId: assignedUser.id,
          outcome: outcome,
          value: opportunity.monetaryValue || 0,
          notes: `Backfilled from GHL - Status: ${opportunity.status}`,
          stageAtOutcome: opportunity.pipelineStageId,
        });

        created++;
        
        if (processed % 50 === 0) {
          console.log(`📊 Progress: ${processed}/${opportunities.length} processed, ${created} created, ${skipped} skipped`);
        }

      } catch (error) {
        console.error(`❌ Error processing opportunity ${opportunity.id}:`, error.message);
        errors++;
      }
    }

    console.log('\n✅ Backfill completed!');
    console.log(`📊 Summary:`);
    console.log(`   - Total processed: ${processed}`);
    console.log(`   - Outcomes created: ${created}`);
    console.log(`   - Skipped: ${skipped}`);
    console.log(`   - Errors: ${errors}`);

  } catch (error) {
    console.error('❌ Backfill failed:', error);
    throw error;
  } finally {
    await prisma.$disconnect();
  }
}

/**
 * Determine outcome from GHL opportunity data
 */
function determineOutcomeFromGHL(opportunity: any): 'WON' | 'LOST' | 'ABANDONED' | 'IN_PROGRESS' | null {
  const status = opportunity.status?.toLowerCase();
  const stageId = opportunity.pipelineStageId;

  // Check for won status/tags
  if (status === 'won' || status === 'closed won' || status === 'sold') {
    return 'WON';
  }

  // Check for lost status/tags
  if (status === 'lost' || status === 'closed lost' || status === 'no sale') {
    return 'LOST';
  }

  // Check for abandoned status
  if (status === 'abandoned' || status === 'inactive') {
    return 'ABANDONED';
  }

  // Check tags for won/lost indicators
  const tags = opportunity.contact?.tags || opportunity.tags || [];
  if (Array.isArray(tags)) {
    for (const tag of tags) {
      const tagText = (typeof tag === 'string' ? tag : tag.name || tag.title || '').toLowerCase();
      if (tagText.includes('won') || tagText.includes('sold') || tagText.includes('closed won')) {
        return 'WON';
      }
      if (tagText.includes('lost') || tagText.includes('no sale') || tagText.includes('closed lost')) {
        return 'LOST';
      }
    }
  }

  // Default to in progress
  return 'IN_PROGRESS';
}

// Run the script if called directly
if (require.main === module) {
  backfillOpportunityOutcomes()
    .then(() => {
      console.log('✅ Script completed successfully');
      process.exit(0);
    })
    .catch((error) => {
      console.error('❌ Script failed:', error);
      process.exit(1);
    });
}

export { backfillOpportunityOutcomes };
