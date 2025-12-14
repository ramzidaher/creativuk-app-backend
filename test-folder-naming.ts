import { PrismaClient } from '@prisma/client';
import * as dotenv from 'dotenv';

// Load environment variables
dotenv.config();

const prisma = new PrismaClient();

/**
 * Helper function to check if a value is empty or represents "unknown"
 */
function isUnknownOrEmpty(value: string | null | undefined): boolean {
  if (!value) return true;
  const normalized = value.trim().toLowerCase();
  return normalized === '' || 
         normalized === 'unknown' || 
         normalized === 'unknown customer' ||
         normalized === 'unknownpos' ||
         normalized === 'unknown pos' ||
         normalized === 'unknownposition' ||
         normalized === 'unknown position';
}

/**
 * Create folder name for outcome-based organization (copied from service)
 */
function createOutcomeFolderName(customerName: string, postcode: string, opportunityId: string): string {
  // Clean and validate customer name
  let cleanName = '';
  if (customerName && !isUnknownOrEmpty(customerName)) {
    cleanName = customerName
      .replace(/[<>:"/\\|?*]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 30);
  }

  // Clean and validate postcode
  let cleanPostcode = '';
  if (postcode && !isUnknownOrEmpty(postcode)) {
    cleanPostcode = postcode
      .replace(/[<>:"/\\|?*]/g, '')
      .replace(/\s+/g, '')
      .trim()
      .substring(0, 10);
  }

  // Build folder name with consistent format
  const parts: string[] = [];
  
  if (cleanName) {
    parts.push(cleanName);
  }
  
  if (cleanPostcode) {
    parts.push(cleanPostcode);
  }
  
  // If we have at least one valid part, use it; otherwise just use opportunity ID
  if (parts.length > 0) {
    return `${parts.join(' ')} - ${opportunityId}`;
  } else {
    // Fallback: use opportunity ID only if both customer name and postcode are missing
    return `Opportunity ${opportunityId}`;
  }
}

/**
 * Create folder name without postcode (copied from service)
 */
function createFolderName(customerName: string, opportunityId: string): string {
  // Clean and validate customer name
  let cleanName = '';
  if (customerName && !isUnknownOrEmpty(customerName)) {
    cleanName = customerName
      .replace(/[<>:"/\\|?*]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 50);
  }

  // If we have a valid customer name, use it; otherwise use opportunity ID only
  if (cleanName) {
    return `${cleanName} - ${opportunityId}`;
  } else {
    // Fallback: use opportunity ID only if customer name is missing
    return `Opportunity ${opportunityId}`;
  }
}

/**
 * Test the folder naming logic
 */
async function testFolderNaming(opportunityId: string) {
  console.log('='.repeat(80));
  console.log(`Testing Folder Naming Logic for Opportunity: ${opportunityId}`);
  console.log('='.repeat(80));
  console.log();

  try {
    // Fetch opportunity progress from database
    const opportunityProgress = await prisma.opportunityProgress.findUnique({
      where: {
        ghlOpportunityId: opportunityId
      },
      select: {
        ghlOpportunityId: true,
        contactAddress: true,
        contactPostcode: true,
        userId: true,
        createdAt: true,
        updatedAt: true
      }
    });

    // Fetch calculator progress (most reliable source)
    const calculatorProgress = await prisma.calculatorProgress.findFirst({
      where: {
        opportunityId: opportunityId
      },
      orderBy: {
        updatedAt: 'desc'
      }
    });

    let calcCustomerName: string | null = null;
    let calcPostcode: string | null = null;
    if (calculatorProgress && calculatorProgress.data) {
      const data = calculatorProgress.data as any;
      if (data.customerDetails) {
        calcCustomerName = data.customerDetails.customerName || null;
        calcPostcode = data.customerDetails.postcode || null;
      }
    }

    if (!opportunityProgress && !calculatorProgress) {
      console.log('❌ Opportunity not found in database');
      console.log();
      console.log('Testing with sample data instead...');
      console.log();
      
      // Test with various scenarios
      const testCases = [
        { customerName: 'Unknown Customer', postcode: 'UnknownPos', description: 'Both unknown' },
        { customerName: 'Unknown Customer', postcode: 'SW1A1AA', description: 'Unknown customer, valid postcode' },
        { customerName: 'John Smith', postcode: 'UnknownPos', description: 'Valid customer, unknown postcode' },
        { customerName: 'John Smith', postcode: 'SW1A1AA', description: 'Both valid' },
        { customerName: '', postcode: '', description: 'Both empty' },
        { customerName: null, postcode: null, description: 'Both null' },
        { customerName: 'Test Customer', postcode: 'Unknown Pos', description: 'Valid customer, unknown pos with space' },
      ];

      console.log('Test Cases:');
      console.log('-'.repeat(80));
      
      for (const testCase of testCases) {
        const folderName = createOutcomeFolderName(
          testCase.customerName || '',
          testCase.postcode || '',
          opportunityId
        );
        console.log(`\n📋 ${testCase.description}:`);
        console.log(`   Customer Name: "${testCase.customerName || 'null'}"`);
        console.log(`   Postcode: "${testCase.postcode || 'null'}"`);
        console.log(`   Result: "${folderName}"`);
      }
      
      return;
    }

    console.log('✅ Opportunity found in database');
    console.log();
    console.log('Database Data:');
    console.log('-'.repeat(80));
    
    if (opportunityProgress) {
      console.log(`📊 OpportunityProgress:`);
      console.log(`   Contact Address: ${opportunityProgress.contactAddress || '(not set)'}`);
      console.log(`   Contact Postcode: ${opportunityProgress.contactPostcode || '(not set)'}`);
      console.log(`   User ID: ${opportunityProgress.userId}`);
      console.log(`   Created: ${opportunityProgress.createdAt}`);
      console.log(`   Updated: ${opportunityProgress.updatedAt}`);
    } else {
      console.log(`📊 OpportunityProgress: (not found)`);
    }
    
    if (calculatorProgress) {
      console.log(`\n📊 CalculatorProgress:`);
      console.log(`   Calculator Type: ${calculatorProgress.calculatorType}`);
      console.log(`   Customer Name: ${calcCustomerName || '(not set)'}`);
      console.log(`   Postcode: ${calcPostcode || '(not set)'}`);
      console.log(`   Updated: ${calculatorProgress.updatedAt}`);
    } else {
      console.log(`\n📊 CalculatorProgress: (not found)`);
    }
    console.log();

    // Test with actual data from database (prioritize CalculatorProgress)
    const customerName = calcCustomerName || opportunityProgress?.contactAddress || 'Unknown Customer';
    const postcode = calcPostcode || opportunityProgress?.contactPostcode || 'UnknownPos';

    console.log('Folder Naming Results:');
    console.log('-'.repeat(80));
    
    // Test outcome folder name (with postcode)
    const outcomeFolderName = createOutcomeFolderName(customerName, postcode, opportunityId);
    console.log(`\n📁 Outcome Folder Name (with postcode):`);
    console.log(`   Input - Customer Name: "${customerName}"`);
    console.log(`   Input - Postcode: "${postcode}"`);
    console.log(`   Output: "${outcomeFolderName}"`);
    console.log();

    // Test regular folder name (without postcode)
    const regularFolderName = createFolderName(customerName, opportunityId);
    console.log(`📁 Regular Folder Name (without postcode):`);
    console.log(`   Input - Customer Name: "${customerName}"`);
    console.log(`   Output: "${regularFolderName}"`);
    console.log();

    // Test with different outcomes
    console.log('Different Outcomes:');
    console.log('-'.repeat(80));
    console.log(`   Quotations Folder: "Customer Quotations/${outcomeFolderName}"`);
    console.log(`   Orders Folder: "Customer Orders/${outcomeFolderName}"`);
    console.log();

    // Additional test scenarios
    console.log('Additional Test Scenarios:');
    console.log('-'.repeat(80));
    
    const scenarios = [
      { name: customerName, postcode: postcode, desc: 'Current database values' },
      { name: 'Unknown Customer', postcode: postcode, desc: 'Unknown customer, current postcode' },
      { name: customerName, postcode: 'UnknownPos', desc: 'Current customer, unknown postcode' },
      { name: 'Unknown Customer', postcode: 'UnknownPos', desc: 'Both unknown' },
      { name: customerName || 'John Smith', postcode: 'SW1A1AA', desc: 'Valid customer and postcode' },
    ];

    for (const scenario of scenarios) {
      const result = createOutcomeFolderName(scenario.name, scenario.postcode, opportunityId);
      console.log(`\n   ${scenario.desc}:`);
      console.log(`   → "${result}"`);
    }

  } catch (error) {
    console.error('❌ Error:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
const opportunityId = process.argv[2] || 'ZR7YaU5mg42YI3sXaHBk';
testFolderNaming(opportunityId)
  .then(() => {
    console.log();
    console.log('='.repeat(80));
    console.log('Test completed!');
    console.log('='.repeat(80));
    process.exit(0);
  })
  .catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });

