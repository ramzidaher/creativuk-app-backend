export class OpportunityDto {
  id: string;
  title: string;
  stageName: string;
  monetaryValue: number;
  contactId: string;
  locationId: string;
  companyId: string;
  createdAt: string;
  updatedAt: string;
}

export class AiVsManualOpportunitiesDto {
  ai: {
    opportunities: OpportunityDto[];
    count: number;
    totalValue: number;
    stageName: string;
  };
  manual: {
    opportunities: OpportunityDto[];
    count: number;
    totalValue: number;
    stageName: string;
  };
  summary: {
    totalOpportunities: number;
    totalValue: number;
  };
}

export class OpportunityStageDto {
  id: string;
  name: string;
  probability: number;
  color: string;
  order: number;
} 