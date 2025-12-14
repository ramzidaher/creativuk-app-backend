import { IsString, IsOptional, IsDateString, IsEnum, IsBoolean, IsNumber, IsArray } from 'class-validator';

export enum SurveyStatus {
  DRAFT = 'DRAFT',
  IN_PROGRESS = 'IN_PROGRESS',
  COMPLETED = 'COMPLETED',
  SUBMITTED = 'SUBMITTED',
  APPROVED = 'APPROVED',
  REJECTED = 'REJECTED'
}

export enum HomeOwnerAvailability {
  YES_SKIP_NEXT = 'YES_SKIP_NEXT',
  NO_REBOOK_APPOINTMENT = 'NO_REBOOK_APPOINTMENT'
}

export class SurveyPage1Dto {
  @IsString()
  date: string;

  @IsString()
  renewableExecutiveFirstName: string;

  @IsString()
  renewableExecutiveLastName: string;

  @IsString()
  customerFirstName: string;

  @IsString()
  customerLastName: string;

  @IsOptional()
  @IsString()
  customer2FirstName?: string;

  @IsOptional()
  @IsString()
  customer2LastName?: string;

  @IsString()
  addressLine1: string;

  @IsOptional()
  @IsString()
  addressLine2?: string;

  @IsString()
  town: string;

  @IsString()
  county: string;

  @IsString()
  postcode: string;

  @IsEnum(HomeOwnerAvailability)
  homeOwnersAvailable: HomeOwnerAvailability;

  @IsOptional()
  @IsString()
  appointmentDateTime?: string;
}

export class SurveyPage2Dto {
  @IsArray()
  @IsString({ each: true })
  selectedReasons: string[];
}

export class SurveyPage3Dto {
  @IsOptional()
  @IsString()
  property?: string;

  @IsOptional()
  @IsString()
  propertyType?: string;

  @IsOptional()
  @IsString()
  bedrooms?: string;

  @IsOptional()
  @IsString()
  duration?: string;

  @IsOptional()
  @IsString()
  movingPlans?: string;

  @IsOptional()
  @IsString()
  occupants?: string;
}

export class SurveyPage4Dto {
  @IsOptional()
  @IsString()
  heatingType?: string;

  @IsOptional()
  @IsString()
  additionalFeatures?: string;

  @IsOptional()
  @IsString()
  prepaidMeter?: string;

  @IsOptional()
  @IsString()
  phaseMeter?: string;

  @IsOptional()
  @IsString()
  energyCompany?: string;

  @IsOptional()
  @IsString()
  monthlyElectricSpend?: string;

  @IsOptional()
  @IsString()
  electricPricePerUnit?: string;

  @IsOptional()
  @IsString()
  annualElectricUsage?: string;

  @IsOptional()
  @IsString()
  energyBillImage?: string;

  @IsOptional()
  @IsArray()
  energyBillFiles?: any[];
}

export class SurveyPage5Dto {
  @IsOptional()
  @IsString()
  epcRating?: string;

  @IsOptional()
  @IsString()
  epcCertificateImage?: string;

  @IsOptional()
  @IsArray()
  epcCertificateFiles?: any[];

  @IsOptional()
  @IsString()
  previousSolarFunding?: string;

  @IsOptional()
  @IsString()
  previousCompany?: string;
}

export class SurveyPage6Dto {
  @IsOptional()
  @IsString()
  additionalFeatures?: string;

  @IsOptional()
  @IsString()
  creditRating?: string;

  @IsOptional()
  @IsString()
  installationAvailability?: string;

  @IsOptional()
  @IsArray()
  frontDoorFiles?: any[];

  @IsOptional()
  @IsArray()
  frontPropertyFiles?: any[];

  @IsOptional()
  @IsArray()
  targetRoofsFiles?: any[];

  @IsOptional()
  @IsArray()
  propertySidesFiles?: any[];
}

export class SurveyPage7Dto {
  @IsOptional()
  @IsString()
  frontDoorImage?: string;

  @IsOptional()
  @IsString()
  frontPropertyImage?: string;

  @IsOptional()
  @IsString()
  targetRoofsImage?: string;

  @IsOptional()
  @IsString()
  propertySidesImage?: string;

  @IsOptional()
  @IsString()
  roofAngleImage?: string;

  @IsOptional()
  @IsString()
  otherRoofImages?: string;

  @IsOptional()
  @IsString()
  roofTileType?: string;

  @IsOptional()
  @IsString()
  roofType?: string;

  @IsOptional()
  @IsString()
  roofTileCloseupImage?: string;

  @IsOptional()
  @IsString()
  otherBuildingsImage?: string;

  @IsOptional()
  @IsString()
  electricMeterImage?: string;

  @IsOptional()
  @IsString()
  garageImage?: string;

  @IsOptional()
  @IsString()
  fuseBoardImage?: string;

  @IsOptional()
  @IsString()
  batteryInverterLocationImage?: string;

  @IsOptional()
  @IsString()
  hasSolarBattery?: string;

  @IsOptional()
  @IsArray()
  roofAngleFiles?: any[];

  @IsOptional()
  @IsArray()
  otherRoofPicturesFiles?: any[];

  @IsOptional()
  @IsArray()
  roofTileCloseupFiles?: any[];

  @IsOptional()
  @IsArray()
  otherBuildingsFiles?: any[];

  @IsOptional()
  @IsArray()
  electricMeterFiles?: any[];

  @IsOptional()
  @IsArray()
  garageFiles?: any[];

  @IsOptional()
  @IsArray()
  fuseBoardFiles?: any[];

  @IsOptional()
  @IsArray()
  batteryInverterLocationFiles?: any[];
}

export class SurveyPage8Dto {
  @IsOptional()
  @IsString()
  evLocation?: string;

  @IsOptional()
  @IsString()
  evChargerRequired?: string;

  @IsOptional()
  @IsString()
  optimisersRequired?: string;

  @IsOptional()
  @IsString()
  optimiserDetails?: string;

  @IsOptional()
  @IsString()
  shadingIssues?: string;

  @IsOptional()
  @IsString()
  scaffoldingRequired?: string;

  @IsOptional()
  @IsString()
  scaffoldingThroughHouse?: string;

  @IsOptional()
  @IsString()
  scaffoldingImages?: string;

  @IsOptional()
  @IsString()
  furtherInformation?: string;

  @IsOptional()
  @IsArray()
  evLocationFiles?: any[];

  @IsOptional()
  @IsArray()
  evChargerFiles?: any[];

  @IsOptional()
  @IsArray()
  shadingIssuesFiles?: any[];

  @IsOptional()
  @IsArray()
  scaffoldingFiles?: any[];

  @IsOptional()
  @IsArray()
  customerSignatureFiles?: any[];

  @IsOptional()
  @IsArray()
  renewableExecutiveSignatureFiles?: any[];
}

export class CompleteSurveyDto {
  @IsString()
  ghlOpportunityId: string;

  @IsString()
  ghlUserId: string;

  page1: SurveyPage1Dto;
  page2: SurveyPage2Dto;
  page3: SurveyPage3Dto;
  page4: SurveyPage4Dto;
  page5: SurveyPage5Dto;
  page6: SurveyPage6Dto;
  page7: SurveyPage7Dto;
  page8: SurveyPage8Dto;

  @IsEnum(SurveyStatus)
  status: SurveyStatus;

  @IsOptional()
  @IsNumber()
  eligibilityScore?: number;

  @IsOptional()
  @IsString()
  rejectionReason?: string;
}

export class SurveyResponseDto {
  id: string;
  ghlOpportunityId: string;
  ghlUserId: string;
  status: SurveyStatus;
  eligibilityScore?: number;
  rejectionReason?: string;
  page1?: SurveyPage1Dto;
  page2?: SurveyPage2Dto;
  page3?: SurveyPage3Dto;
  page4?: SurveyPage4Dto;
  page5?: SurveyPage5Dto;
  page6?: SurveyPage6Dto;
  page7?: SurveyPage7Dto;
  page8?: SurveyPage8Dto;
  createdAt: Date;
  updatedAt: Date;
  submittedAt?: Date;
  approvedAt?: Date;
  rejectedAt?: Date;
}

export class UpdateSurveyDto {
  @IsOptional() page1?: SurveyPage1Dto;
  @IsOptional() page2?: SurveyPage2Dto;
  @IsOptional() page3?: SurveyPage3Dto;
  @IsOptional() page4?: SurveyPage4Dto;
  @IsOptional() page5?: SurveyPage5Dto;
  @IsOptional() page6?: SurveyPage6Dto;
  @IsOptional() page7?: SurveyPage7Dto;
  @IsOptional() page8?: SurveyPage8Dto;
  @IsOptional() @IsEnum(SurveyStatus) status?: SurveyStatus;
  @IsOptional() @IsNumber() eligibilityScore?: number;
  @IsOptional() @IsString() rejectionReason?: string;
} 