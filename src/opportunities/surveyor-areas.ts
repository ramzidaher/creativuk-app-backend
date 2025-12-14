export interface SurveyorArea {
  name: string;
  location: string;
  areas: string[];
  maxTravelTime: string;
  ghlUserId?: string;
  ghlUserData?: any;
}

export const SURVEYOR_AREAS: SurveyorArea[] = [
  {
    name: 'Andrew Hughes',
    location: 'CB10 1PL',
    areas: ['PE', 'CB', 'SG', 'CM', 'CO', 'SS', 'IP', 'AL', 'NN'],
    maxTravelTime: '1hour 30 mins MAX'
  },
  {
    name: 'Ion Zacon',
    location: 'TN37 6BS',
    areas: ['RH', 'ME', 'TN', 'BN', 'CT', 'BR', 'DA', 'CR', 'KT'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Jordan Stewart',
    location: 'HD7 5PN',
    areas: ['BD', 'HG', 'LS', 'WF', 'HD', 'S', 'SK', 'M', 'OL', 'BB', 'PR', 'BL'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Onur Saliah',
    location: 'CM21 9ET',
    areas: ['MK', 'PE', 'LU', 'AL', 'SG', 'CM', 'CO', 'CB', 'IP', 'SS', 'RM', 'IG', 'EN'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'James Barnett',
    location: 'PL4 7HG',
    areas: ['TA', 'EX', 'TQ', 'PL', 'TR', 'CT'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Kenji Omachi',
    location: 'HA9 9LJ',
    areas: ['NN', 'OX', 'MK', 'SG', 'LU', 'HP', 'SL', 'AL', 'RG', 'GU', 'RH', 'CM', 'SS', 'LONDON', 'W', 'KT', 'TW'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Iuzu Alexandru',
    location: 'TN38 9HR',
    areas: ['TN', 'BN', 'ME', 'CT', 'RH'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Miles Kent',
    location: 'AL5 4NZ',
    areas: ['MK', 'HP', 'AL', 'CM', 'SG', 'WD', 'EN', 'IG', 'RM'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Hamzah Islam',
    location: 'CV2 4RD',
    areas: ['ST', 'TF', 'WV', 'WS', 'B', 'DY', 'WR', 'CV', 'NN', 'LE', 'OX'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Karl Gedney',
    location: 'London',
    areas: ['LONDON', 'W', 'E', 'N', 'SW', 'SE', 'NW', 'NE', 'KT', 'TW', 'CR', 'SM', 'BR', 'DA'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Kemberly Willocks',
    location: 'Manchester',
    areas: ['M', 'OL', 'BL', 'SK', 'BB', 'PR', 'WN', 'WA', 'L', 'FY', 'LA'],
    maxTravelTime: '9:30am to 6:30pm'
  },
  {
    name: 'Robert Koch',
    location: 'GU28 0EF',
    areas: ['GU', 'RG', 'SO', 'PO', 'SP', 'BH', 'SN'],
    maxTravelTime: '9:30am to 6:30pm',
    ghlUserId: 'er2hzTmwr4zgFoBpaaS1'
  },
  {
    name: 'Terrence Koch',
    location: 'CV21 2AB',
    areas: ['CV', 'LE', 'NN', 'MK', 'OX', 'GL', 'WR', 'B'],
    maxTravelTime: '9:30am to 6:30pm',
    ghlUserId: '1H8Dos8NFvnEV3RzL4Y5'
  }
];

export function getSurveyorByName(name: string): SurveyorArea | undefined {
  return SURVEYOR_AREAS.find(surveyor => 
    surveyor.name.toLowerCase().includes(name.toLowerCase()) ||
    name.toLowerCase().includes(surveyor.name.toLowerCase())
  );
}

export function isOpportunityAssignedToSurveyor(opportunity: any, surveyorName: string, userMap?: Map<string, any>): boolean {
  const surveyor = getSurveyorByName(surveyorName);
  if (!surveyor) {
    console.log(`Surveyor not found for name: ${surveyorName}`);
    return false;
  }

  console.log(`Checking opportunity: ${opportunity.name} for surveyor: ${surveyor.name} (${surveyor.areas.join(', ')})`);

  // Check if opportunity is assigned to this surveyor by user ID
  const assignedToId = opportunity.assignedTo;
  if (assignedToId && userMap) {
    const assignedUser = userMap.get(assignedToId);
    if (assignedUser) {
      const assignedUserName = assignedUser.name || assignedUser.firstName + ' ' + assignedUser.lastName;
      console.log(`Checking assignedTo user: ${assignedUserName} against surveyor: ${surveyor.name}`);
      
      // Check both normal and reversed name formats
      const surveyorNameLower = surveyor.name.toLowerCase();
      const assignedUserNameLower = assignedUserName.toLowerCase();
      
      // Check if names match (normal order)
      if (assignedUserNameLower.includes(surveyorNameLower)) {
        console.log(`✅ Match found by user assignment (normal): ${assignedUserName}`);
        return true;
      }
      
      // Check if names match (reversed order)
      const surveyorNameParts = surveyorNameLower.split(' ');
      if (surveyorNameParts.length === 2) {
        const reversedSurveyorName = `${surveyorNameParts[1]} ${surveyorNameParts[0]}`;
        if (assignedUserNameLower.includes(reversedSurveyorName)) {
          console.log(`✅ Match found by user assignment (reversed): ${assignedUserName}`);
          return true;
        }
      }
      
      // Check if any part of the surveyor name matches
      const surveyorNameWords = surveyorNameLower.split(' ');
      const assignedNameWords = assignedUserNameLower.split(' ');
      
      for (const surveyorWord of surveyorNameWords) {
        for (const assignedWord of assignedNameWords) {
          if (surveyorWord.length > 2 && assignedWord.length > 2 && surveyorWord === assignedWord) {
            console.log(`✅ Match found by user assignment (partial): ${assignedUserName} contains "${surveyorWord}"`);
            return true;
          }
        }
      }
    }
  }

  // Fallback: Check if opportunity is assigned to this surveyor by name
  const assignedTo = opportunity.assignedTo || opportunity.teamMember || opportunity.assignee || opportunity.assignedToName;
  if (assignedTo) {
    console.log(`Checking assignedTo: ${assignedTo} against surveyor: ${surveyor.name}`);
    
    const assignedToLower = assignedTo.toLowerCase();
    const surveyorNameLower = surveyor.name.toLowerCase();
    
    // Check if names match (normal order)
    if (assignedToLower.includes(surveyorNameLower)) {
      console.log(`✅ Match found by assignment (normal): ${assignedTo}`);
      return true;
    }
    
    // Check if names match (reversed order)
    const surveyorNameParts = surveyorNameLower.split(' ');
    if (surveyorNameParts.length === 2) {
      const reversedSurveyorName = `${surveyorNameParts[1]} ${surveyorNameParts[0]}`;
      if (assignedToLower.includes(reversedSurveyorName)) {
        console.log(`✅ Match found by assignment (reversed): ${assignedTo}`);
        return true;
      }
    }
    
    // Check if any part of the surveyor name matches
    const surveyorNameWords = surveyorNameLower.split(' ');
    const assignedToWords = assignedToLower.split(' ');
    
    for (const surveyorWord of surveyorNameWords) {
      for (const assignedWord of assignedToWords) {
        if (surveyorWord.length > 2 && assignedWord.length > 2 && surveyorWord === assignedWord) {
          console.log(`✅ Match found by assignment (partial): ${assignedTo} contains "${surveyorWord}"`);
          return true;
        }
      }
    }
  }

  console.log(`❌ No match found for opportunity - only checking by name assignment`);
  return false;
} 