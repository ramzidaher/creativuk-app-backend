import { Test, TestingModule } from '@nestjs/testing';
import { ConfigService } from '@nestjs/config';
import { EmailService } from './email.service';
import { SurveyResponseDto, SurveyStatus } from '../opportunities/dto/survey.dto';

describe('EmailService', () => {
  let service: EmailService;
  let configService: ConfigService;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [
        EmailService,
        {
          provide: ConfigService,
          useValue: {
            get: jest.fn((key: string) => {
              switch (key) {
                case 'SMTP_HOST':
                  return 'smtp.gmail.com';
                case 'SMTP_PORT':
                  return 587;
                case 'SMTP_USER':
                  return 'test@example.com';
                case 'SMTP_PASS':
                  return 'test-password';
                case 'SMTP_FROM':
                  return 'noreply@creativsolar.com';
                default:
                  return undefined;
              }
            }),
          },
        },
      ],
    }).compile();

    service = module.get<EmailService>(EmailService);
    configService = module.get<ConfigService>(ConfigService);
  });

  it('should be defined', () => {
    expect(service).toBeDefined();
  });

  it('should generate email content correctly', () => {
    const mockSurvey: SurveyResponseDto = {
      id: 'test-id',
      ghlOpportunityId: 'test-opportunity',
      ghlUserId: 'test-user',
      status: SurveyStatus.SUBMITTED,
      eligibilityScore: 85,
      page1: {
        renewableExecutiveFirstName: 'John',
        renewableExecutiveLastName: 'Doe',
        customerFirstName: 'Jane',
        customerLastName: 'Smith',
        addressLine1: '123 Test Street',
        town: 'Test Town',
        county: 'Test County',
        postcode: 'TE1 1ST',
        homeOwnersAvailable: 'YES' as any,
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    };

    const content = (service as any).generateSurveyEmailContent(mockSurvey);
    
    expect(content).toContain('John Doe');
    expect(content).toContain('Jane Smith');
    expect(content).toContain('123 Test Street');
    expect(content).toContain('85/100');
    expect(content).toContain('SUBMITTED');
  });

  it('should generate attachments correctly', () => {
    const mockSurvey: SurveyResponseDto = {
      id: 'test-id',
      ghlOpportunityId: 'test-opportunity',
      ghlUserId: 'test-user',
      status: SurveyStatus.SUBMITTED,
      page4: {
        energyBillImage: 'test-energy-bill.jpg',
      },
      page5: {
        epcCertificateImage: 'test-epc.jpg',
      },
      page7: {
        frontDoorImage: 'test-front-door.jpg',
        frontPropertyImage: 'test-front-property.jpg',
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    };

    const attachments = (service as any).generateAttachments(mockSurvey);
    
    expect(attachments).toHaveLength(4);
    expect(attachments[0].filename).toBe('energy-bill.jpg');
    expect(attachments[1].filename).toBe('epc-certificate.jpg');
    expect(attachments[2].filename).toBe('front-door.jpg');
    expect(attachments[3].filename).toBe('front-property.jpg');
  });
});
