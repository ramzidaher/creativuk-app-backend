import { Injectable, Logger, NotFoundException, BadRequestException } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { UserService } from '../user/user.service';
import { EmailService } from '../email/email.service';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { SurveyImageService } from '../survey/survey-image.service';
import { SurveyReportService } from '../survey/survey-report.service';
import { CloudinaryService } from '../cloudinary/cloudinary.service';
import { 
  SurveyResponseDto, 
  CompleteSurveyDto, 
  UpdateSurveyDto,
  SurveyStatus,
  SurveyPage1Dto,
  SurveyPage2Dto,
  SurveyPage3Dto,
  SurveyPage4Dto,
  SurveyPage5Dto,
  SurveyPage6Dto,
  SurveyPage7Dto
} from './dto/survey.dto';

@Injectable()
export class SurveyService {
  private readonly logger = new Logger(SurveyService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly userService: UserService,
    private readonly emailService: EmailService,
    private readonly goHighLevelService: GoHighLevelService,
    private readonly surveyImageService: SurveyImageService,
    private readonly surveyReportService: SurveyReportService,
    private readonly cloudinaryService: CloudinaryService,
  ) {}

  async createSurvey(ghlUserId: string, ghlOpportunityId: string): Promise<SurveyResponseDto> {
    this.logger.log(`Creating survey for ghlUserId: ${ghlUserId}, ghlOpportunityId: ${ghlOpportunityId}`);
    
    let user;
    
    // Handle the case where ghlUserId is 'default' (testing mode)
    if (ghlUserId === 'default') {
      // Get the first available user for testing
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException('No active users found for testing');
      }
    } else {
      // First try to find by ghlUserId
      user = await this.userService.findByGhlUserId(ghlUserId);
      this.logger.log(`User lookup by ghlUserId ${ghlUserId}: ${user ? 'found' : 'not found'}`);
      
      // If not found, try to find by internal ID (fallback)
      if (!user) {
        user = await this.userService.findById(ghlUserId);
        this.logger.log(`User lookup by internal ID ${ghlUserId}: ${user ? 'found' : 'not found'}`);
      }
      
      // If still not found, get any active user for testing
      if (!user) {
        this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
        user = await this.prisma.user.findFirst({
          where: { status: 'ACTIVE' }
        });
        if (!user) {
          throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
        }
      }
    }

    this.logger.log(`Using user: ${user.id}, ghlUserId: ${user.ghlUserId}, name: ${user.name}`);

    // Check if survey already exists
    const existingSurvey = await this.prisma.survey.findUnique({
      where: { ghlOpportunityId }
    });

    if (existingSurvey) {
      this.logger.log(`Survey already exists for opportunity ${ghlOpportunityId}`);
      return this.mapToDto(existingSurvey);
    }

    // Create new survey - use the user's ghlUserId to avoid foreign key constraint violations
    const survey = await this.prisma.survey.create({
      data: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId, // Use the user's ghlUserId to avoid foreign key constraint violations
        status: SurveyStatus.DRAFT,
        createdBy: user.id // Add the createdBy field
      }
    });

    this.logger.log(`Created survey for opportunity ${ghlOpportunityId} with ghlUserId: ${user.ghlUserId}`);
    return this.mapToDto(survey);
  }

  async getSurvey(ghlUserId: string, ghlOpportunityId: string): Promise<SurveyResponseDto> {
    let user;
    
    // Handle the case where ghlUserId is 'default' (testing mode)
    if (ghlUserId === 'default') {
      // Get the first available user for testing
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException('No active users found for testing');
      }
    } else {
      // First try to find by ghlUserId
      user = await this.userService.findByGhlUserId(ghlUserId);
      
      // If not found, try to find by internal ID
      if (!user) {
        user = await this.userService.findById(ghlUserId);
      }
      
      // If still not found, get any active user for testing
      if (!user) {
        this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
        user = await this.prisma.user.findFirst({
          where: { status: 'ACTIVE' }
        });
        if (!user) {
          throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
        }
      }
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId || ghlUserId, // Use user's ghlUserId or fallback
        isDeleted: false
      }
    });

    if (!survey) {
      // Instead of throwing an error, create a new survey
      this.logger.log(`No survey found for opportunity ${ghlOpportunityId}, creating new one`);
      return this.createSurvey(ghlUserId, ghlOpportunityId);
    }

    return this.mapToDto(survey);
  }

  async getUserSurveys(ghlUserId: string): Promise<SurveyResponseDto[]> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const surveys = await this.prisma.survey.findMany({
      where: {
        ghlUserId: user.ghlUserId,
        isDeleted: false
      },
      orderBy: { updatedAt: 'desc' }
    });

    return surveys.map(survey => this.mapToDto(survey));
  }

  async updateSurvey(ghlUserId: string, ghlOpportunityId: string, updateDto: UpdateSurveyDto): Promise<SurveyResponseDto> {
    this.logger.log(`Updating survey for ghlUserId: ${ghlUserId}, ghlOpportunityId: ${ghlOpportunityId}`);
    
    let user = await this.userService.findByGhlUserId(ghlUserId);
    
    // If not found by ghlUserId, try by internal ID
    if (!user) {
      user = await this.userService.findById(ghlUserId);
    }
    
    // If still not found, get any active user for testing
    if (!user) {
      this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
      }
    }

    this.logger.log(`Using user: ${user.id}, ghlUserId: ${user.ghlUserId}, name: ${user.name}`);

    // First try to find survey with matching ghlUserId
    let survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: ghlUserId, // Use the ghlUserId from the request
        isDeleted: false
      }
    });

    // If not found, try to find survey with null ghlUserId (for existing surveys)
    if (!survey) {
      this.logger.log(`Survey not found with ghlUserId: ${ghlUserId}, trying with null ghlUserId`);
      survey = await this.prisma.survey.findFirst({
        where: {
          ghlOpportunityId,
          ghlUserId: null,
          isDeleted: false
        }
      });
      
      // If found with null ghlUserId, update it to have the correct ghlUserId
      if (survey) {
        this.logger.log(`Found survey with null ghlUserId, updating to: ${ghlUserId}`);
        const validGhlUserId = await this.validateAndGetGhlUserId(ghlUserId, user);
        survey = await this.prisma.survey.update({
          where: { id: survey.id },
          data: { ghlUserId: validGhlUserId }
        });
      }
    }

    if (!survey) {
      this.logger.log(`No survey found for opportunity ${ghlOpportunityId}, creating new one`);
      return this.createSurvey(ghlUserId, ghlOpportunityId);
    }

    // Prepare update data with proper JSON handling
    const updateData: any = {
      updatedAt: new Date(),
    };

    // Only update fields that are provided
    if (updateDto.page1 !== undefined) {
      updateData.page1 = updateDto.page1;
    }
    if (updateDto.page2 !== undefined) {
      updateData.page2 = updateDto.page2;
    }
    if (updateDto.page3 !== undefined) {
      updateData.page3 = updateDto.page3;
    }
    if (updateDto.page4 !== undefined) {
      updateData.page4 = updateDto.page4;
    }
    if (updateDto.page5 !== undefined) {
      updateData.page5 = updateDto.page5;
    }
    if (updateDto.page6 !== undefined) {
      updateData.page6 = updateDto.page6;
    }
    if (updateDto.page7 !== undefined) {
      updateData.page7 = updateDto.page7;
    }
    if (updateDto.page8 !== undefined) {
      updateData.page8 = updateDto.page8;
    }
    if (updateDto.status !== undefined) {
      updateData.status = updateDto.status;
    }
    if (updateDto.eligibilityScore !== undefined) {
      updateData.eligibilityScore = updateDto.eligibilityScore;
    }
    if (updateDto.rejectionReason !== undefined) {
      updateData.rejectionReason = updateDto.rejectionReason;
    }

    // Handle timestamp updates
    if (updateDto.status === SurveyStatus.SUBMITTED) {
      updateData.submittedAt = new Date();
    }
    if (updateDto.status === SurveyStatus.APPROVED) {
      updateData.approvedAt = new Date();
    }
    if (updateDto.status === SurveyStatus.REJECTED) {
      updateData.rejectedAt = new Date();
    }

    // Update survey data
    const updatedSurvey = await this.prisma.survey.update({
      where: { id: survey.id },
      data: updateData
    });

    this.logger.log(`Updated survey ${survey.id} for opportunity ${ghlOpportunityId}`);
    this.logger.log(`Page 8 data in updateData:`, updateData.page8);
    return this.mapToDto(updatedSurvey);
  }

  async finalizeSurveySubmission(ghlUserId: string, ghlOpportunityId: string): Promise<SurveyResponseDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    // Validate survey completion - only require pages 1-5 for submission
    if (!survey.page1 || !survey.page2 || !survey.page3 || !survey.page4 || !survey.page5) {
      throw new BadRequestException('All survey pages must be completed before submission');
    }

    // Get image URLs from database (Cloudinary URLs)
    const surveyImages = await this.surveyImageService.getSurveyImages(survey.id);
    const imageUrls: { [key: string]: string[] } = {};
    
    for (const image of surveyImages) {
      if (!imageUrls[image.fieldName]) {
        imageUrls[image.fieldName] = [];
      }
      imageUrls[image.fieldName].push(image.filePath); // filePath contains Cloudinary URL
    }

    this.logger.log(`Found ${Object.keys(imageUrls).length} image fields with ${Object.values(imageUrls).flat().length} total images`);

    // Create HTML survey report with embedded images
    try {
      const reportFilePath = await this.surveyReportService.generateHtmlReport(ghlOpportunityId, {
        page1: survey.page1,
        page2: survey.page2,
        page3: survey.page3,
        page4: survey.page4,
        page5: survey.page5,
        page6: survey.page6,
        page7: survey.page7,
        page8: survey.page8
      }, survey.id);
      this.logger.log(`Created HTML survey report: ${reportFilePath}`);
    } catch (reportError) {
      this.logger.error(`Failed to create HTML survey report: ${reportError.message}`);
      // Don't fail the survey submission if report creation fails
    }

    // Also create PDF report
    try {
      const pdfReportFilePath = await this.surveyReportService.generatePdfReport(ghlOpportunityId, {
        page1: survey.page1,
        page2: survey.page2,
        page3: survey.page3,
        page4: survey.page4,
        page5: survey.page5,
        page6: survey.page6,
        page7: survey.page7,
        page8: survey.page8
      }, survey.id);
      this.logger.log(`Created PDF survey report: ${pdfReportFilePath}`);
    } catch (pdfReportError) {
      this.logger.error(`Failed to create PDF survey report: ${pdfReportError.message}`);
      // Don't fail the survey submission if PDF report creation fails
    }

    // Prepare update data
    const updateData: any = {
      status: SurveyStatus.SUBMITTED,
      submittedAt: new Date(),
      updatedAt: new Date()
    };

    // Update survey status
    const updatedSurvey = await this.prisma.survey.update({
      where: { id: survey.id },
      data: updateData
    });

    this.logger.log(`Survey ${survey.id} submitted successfully with ${Object.keys(imageUrls).length} image fields`);

    // Send email notification with Cloudinary image URLs
    try {
      const surveyDto = this.mapToDto(updatedSurvey);
      const recipientEmail = 'paldevtechnologies@gmail.com'; // All survey emails sent to this address
      
      // Get opportunity details if available
      let opportunityDetails: any = null;
      try {
        // Fetch opportunity details from GHL to get the lead name
        const user = await this.userService.findByGhlUserId(ghlUserId);
        if (user && user.ghlAccessToken) {
          const ghlOpportunity = await this.goHighLevelService.getOpportunityById(user.ghlAccessToken, ghlOpportunityId);
          if (ghlOpportunity) {
            opportunityDetails = {
              name: ghlOpportunity.name || `Opportunity ${ghlOpportunityId}`,
              ghlOpportunityId,
              createdAt: ghlOpportunity.createdAt || new Date().toISOString(),
              contactName: ghlOpportunity.contact?.firstName && ghlOpportunity.contact?.lastName 
                ? `${ghlOpportunity.contact.firstName} ${ghlOpportunity.contact.lastName}`
                : ghlOpportunity.contact?.firstName || ghlOpportunity.contact?.lastName || 'N/A',
              stageName: ghlOpportunity.stage?.name || 'N/A'
            };
            this.logger.log(`Fetched opportunity details: ${opportunityDetails.name} (${opportunityDetails.contactName})`);
          }
        }
      } catch (error) {
        this.logger.warn('Could not fetch opportunity details for email:', error.message);
        // Fallback to basic details
        opportunityDetails = {
          name: `Opportunity ${ghlOpportunityId}`,
          ghlOpportunityId,
          createdAt: new Date().toISOString(),
          contactName: 'N/A',
          stageName: 'N/A'
        };
      }

      // Pass Cloudinary image URLs to email service
      await this.emailService.sendSurveyResponseEmail(surveyDto, recipientEmail, opportunityDetails, imageUrls);
      this.logger.log(`Survey email sent successfully for opportunity ${ghlOpportunityId}`);
    } catch (emailError) {
      this.logger.error(`Failed to send survey email: ${emailError.message}`);
      // Don't fail the survey submission if email fails
    }

    this.logger.log(`Finalized survey submission ${survey.id} successfully`);
    return this.mapToDto(updatedSurvey);
  }

  async submitSurveyWithImages(ghlUserId: string, ghlOpportunityId: string, surveyData: CompleteSurveyDto, images: any): Promise<SurveyResponseDto> {
    let user;
    
    // Handle the case where ghlUserId is 'default' (testing mode)
    if (ghlUserId === 'default') {
      // Get the first available user for testing
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException('No active users found for testing');
      }
    } else {
      // First try to find by ghlUserId
      user = await this.userService.findByGhlUserId(ghlUserId);
      
      // If not found, try to find by internal ID (in case frontend sent internal ID instead of ghlUserId)
      if (!user) {
        user = await this.userService.findById(ghlUserId);
      }
      
      // If still not found, get any active user for testing
      if (!user) {
        this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
        user = await this.prisma.user.findFirst({
          where: { status: 'ACTIVE' }
        });
        if (!user) {
          throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
        }
      }
    }

    // First, update the survey with the provided data
    // console.log(`🔧 Updating survey for user ${user.ghlUserId || user.id} and opportunity ${ghlOpportunityId}`);
    const survey = await this.updateSurvey(ghlUserId, ghlOpportunityId, surveyData);
    // console.log(`🔧 Survey updated successfully:`, survey.id);

    // Process images and upload them to Cloudinary
    const cloudinaryUrls: { [key: string]: string[] } = {};
    
    if (images && Object.keys(images).length > 0) {
      // console.log(`🔧 Processing ${Object.keys(images).length} image fields for Cloudinary upload...`);
      
      for (const fieldName of Object.keys(images)) {
        const fieldImages = images[fieldName];
        if (fieldImages && fieldImages.length > 0) {
          // console.log(`🔧 Processing ${fieldImages.length} images for field: ${fieldName}`);
          
          for (let index = 0; index < fieldImages.length; index++) {
            const imageData = fieldImages[index];
            
            try {
              // Create unique public ID for Cloudinary
              const publicId = `surveys/${ghlOpportunityId}/${fieldName}-${index}-${Date.now()}`;
              
              // Upload base64 image to Cloudinary
              // console.log(`🔧 Uploading image to Cloudinary: ${fieldName}-${index}`);
              // console.log(`🔧 Base64 data length: ${imageData.base64Data?.length || 0}`);
              // console.log(`🔧 Public ID: ${publicId}`);
              
              const uploadResult = await this.cloudinaryService.uploadBase64Image(
                imageData.base64Data,
                `surveys/${ghlOpportunityId}`,
                publicId,
                {
                  resource_type: 'image',
                  quality: 'auto',
                  fetch_format: 'auto'
                }
              );
              
              // console.log(`✅ Cloudinary upload successful: ${uploadResult.secure_url}`);
              
              // Store the Cloudinary URL
              if (!cloudinaryUrls[fieldName]) {
                cloudinaryUrls[fieldName] = [];
              }
              cloudinaryUrls[fieldName].push(uploadResult.secure_url);
              
              // Save image record to database with Cloudinary URL
              await this.prisma.surveyImage.create({
                data: {
                  surveyId: survey.id,
                  fieldName: fieldName,
                  fileName: imageData.name || `${fieldName}-${index}`,
                  originalName: imageData.name,
                  mimeType: imageData.mimeType,
                  fileSize: imageData.size || 0,
                  filePath: uploadResult.secure_url, // Store Cloudinary URL instead of local path
                  base64Data: null // No need to store base64 data when using Cloudinary
                }
              });
              
              // console.log(`✅ Uploaded ${fieldName}-${index} to Cloudinary: ${uploadResult.secure_url}`);
            } catch (error) {
              console.error(`❌ Error uploading ${fieldName}-${index} to Cloudinary:`, error);
              console.error(`❌ Error details:`, {
                fieldName,
                index,
                base64Length: imageData.base64Data?.length || 0,
                mimeType: imageData.mimeType,
                error: error.message
              });
              // Continue with other images even if one fails
            }
          }
        }
      }
      
      this.logger.log(`Uploaded ${Object.keys(cloudinaryUrls).length} file fields to Cloudinary`);
    } else {
      // console.log(`🔧 No images to process`);
    }

    // Finalize the survey submission (generate reports and send email)
    try {
      // console.log(`🔧 Finalizing survey submission...`);
      const finalizedSurvey = await this.finalizeSurveySubmission(ghlUserId, ghlOpportunityId);
      // console.log(`✅ Survey submission finalized successfully`);
      return finalizedSurvey;
    } catch (error) {
      console.error(`❌ Error finalizing survey submission:`, error);
      // Return the survey even if finalization fails
      return survey;
    }
  }

  async submitSurveyWithFiles(ghlUserId: string, ghlOpportunityId: string, surveyData: CompleteSurveyDto, files: Express.Multer.File[]): Promise<SurveyResponseDto> {
    let user;
    
    // Handle the case where ghlUserId is 'default' (testing mode)
    if (ghlUserId === 'default') {
      // Get the first available user for testing
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException('No active users found for testing');
      }
    } else {
      // First try to find by ghlUserId
      user = await this.userService.findByGhlUserId(ghlUserId);
      
      // If not found, try to find by internal ID (in case frontend sent internal ID instead of ghlUserId)
      if (!user) {
        user = await this.userService.findById(ghlUserId);
      }
      
      // If still not found, get any active user for testing
      if (!user) {
        this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
        user = await this.prisma.user.findFirst({
          where: { status: 'ACTIVE' }
        });
        if (!user) {
          throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
        }
      }
    }

    // First, update the survey with the provided data
    // console.log(`🔧 Updating survey for user ${user.ghlUserId || user.id} and opportunity ${ghlOpportunityId}`);
    const survey = await this.updateSurvey(ghlUserId, ghlOpportunityId, surveyData);
    // console.log(`🔧 Survey updated successfully:`, survey.id);

    // Process uploaded files and upload to Cloudinary
    const cloudinaryUrls: { [key: string]: string[] } = {};
    
    if (files && files.length > 0) {
      // console.log(`📁 Processing ${files.length} files for Cloudinary upload...`);
      
      for (const file of files) {
        try {
          // Extract field name from the fieldname (e.g., "frontProperty-0" -> "frontProperty")
          const fieldName = file.fieldname.split('-')[0];
          const index = file.fieldname.split('-')[1] || '0';
          
          // Determine resource type based on file mime type
          // Upload PDFs as 'image' resource type to avoid Cloudinary free account restrictions
          // Cloudinary supports PDFs as images and will preserve the file format
          const isPdf = file.mimetype === 'application/pdf';
          const resourceType = 'image'; // Always use 'image' for both images and PDFs
          
          // Get file extension from original filename
          const fileExtension = file.originalname?.split('.').pop()?.toLowerCase() || '';
          
          // Create unique public ID for Cloudinary (without folder path - folder is passed separately)
          // Preserve extension for PDFs so Cloudinary knows it's a PDF
          let publicId = `${fieldName}-${index}-${Date.now()}`;
          if (isPdf && fileExtension === 'pdf') {
            publicId = `${publicId}.pdf`;
          }
          
          // Prepare upload options
          const uploadOptions: any = {
            resource_type: resourceType,
            format: isPdf ? 'pdf' : 'auto', // Explicitly set format for PDFs
          };
          
          // Only add image-specific options for non-PDF images
          if (!isPdf) {
            uploadOptions.quality = 'auto';
            uploadOptions.fetch_format = 'auto';
          }
          
          // Upload to Cloudinary
          const uploadResult = await this.cloudinaryService.uploadFile(
            file.buffer,
            `surveys/${ghlOpportunityId}`,
            publicId,
            uploadOptions
          );
          
          // For PDFs, generate a signed URL to bypass "untrusted customer" restrictions
          // The public_id includes the folder path, so we need to use the full public_id from the result
          let fileUrl = uploadResult.secure_url;
          if (isPdf) {
            try {
              // Generate signed URL for PDF to ensure it's accessible
              // Use the public_id from the upload result (which includes folder path)
              const folderPath = `surveys/${ghlOpportunityId}`;
              const fullPublicId = uploadResult.public_id || `${folderPath}/${publicId}`.replace(/\/+/g, '/');
              fileUrl = this.cloudinaryService.generateSignedUrl(fullPublicId, 'image', { format: 'pdf' });
              this.logger.log(`📄 Generated signed URL for PDF: ${fileUrl}`);
            } catch (error) {
              this.logger.warn(`⚠️ Failed to generate signed URL for PDF, using regular URL: ${error.message}`);
              // Fallback to regular URL if signed URL generation fails
              fileUrl = uploadResult.secure_url;
            }
          }
          
          // Store the Cloudinary URL
          if (!cloudinaryUrls[fieldName]) {
            cloudinaryUrls[fieldName] = [];
          }
          cloudinaryUrls[fieldName].push(fileUrl);
          
          // Save to database with Cloudinary URL (signed URL for PDFs)
          await this.prisma.surveyImage.create({
            data: {
              surveyId: survey.id,
              fieldName: fieldName,
              fileName: file.filename,
              originalName: file.originalname,
              mimeType: file.mimetype,
              fileSize: file.size,
              filePath: fileUrl, // Store signed URL for PDFs, regular URL for images
              base64Data: null // No need for base64 data when using Cloudinary
            }
          });
          
          // console.log(`✅ Uploaded ${fieldName}-${index} to Cloudinary: ${uploadResult.secure_url}`);
        } catch (error) {
          console.error(`❌ Error uploading ${file.fieldname} to Cloudinary:`, error);
          // Continue with other files even if one fails
        }
      }
      
      this.logger.log(`Uploaded ${Object.keys(cloudinaryUrls).length} file fields to Cloudinary`);
    }

    // Finalize the survey submission (generate reports and send email)
    try {
      // console.log(`🔧 Finalizing survey submission...`);
      const finalizedSurvey = await this.finalizeSurveySubmission(ghlUserId, ghlOpportunityId);
      // console.log(`✅ Survey submission finalized successfully`);
      return finalizedSurvey;
    } catch (error) {
      console.error(`❌ Error finalizing survey submission:`, error);
      // Return the survey even if finalization fails
      return survey;
    }
    return survey;
  }


  async approveSurvey(ghlUserId: string, ghlOpportunityId: string): Promise<SurveyResponseDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    if (survey.status !== SurveyStatus.SUBMITTED) {
      throw new BadRequestException('Survey must be submitted before approval');
    }

    const updatedSurvey = await this.prisma.survey.update({
      where: { id: survey.id },
      data: {
        status: SurveyStatus.APPROVED,
        approvedAt: new Date(),
        updatedAt: new Date()
      }
    });

    this.logger.log(`Approved survey ${survey.id} for opportunity ${ghlOpportunityId}`);
    return this.mapToDto(updatedSurvey);
  }

  async rejectSurvey(ghlUserId: string, ghlOpportunityId: string, rejectionReason: string): Promise<SurveyResponseDto> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    const updatedSurvey = await this.prisma.survey.update({
      where: { id: survey.id },
      data: {
        status: SurveyStatus.REJECTED,
        rejectionReason,
        rejectedAt: new Date(),
        updatedAt: new Date()
      }
    });

    this.logger.log(`Rejected survey ${survey.id} for opportunity ${ghlOpportunityId}`);
    return this.mapToDto(updatedSurvey);
  }

  async deleteSurvey(ghlUserId: string, ghlOpportunityId: string): Promise<void> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    // Soft delete
    await this.prisma.survey.update({
      where: { id: survey.id },
      data: {
        isDeleted: true,
        deletedAt: new Date()
      }
    });

    this.logger.log(`Deleted survey ${survey.id} for opportunity ${ghlOpportunityId}`);
  }

  async resetSurvey(ghlUserId: string, ghlOpportunityId: string): Promise<SurveyResponseDto> {
    let user;
    
    // Handle the case where ghlUserId is 'default' (testing mode)
    if (ghlUserId === 'default') {
      // Get the first available user for testing
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException('No active users found for testing');
      }
    } else {
      // First try to find by ghlUserId
      user = await this.userService.findByGhlUserId(ghlUserId);
      
      // If not found, try to find by internal ID
      if (!user) {
        user = await this.userService.findById(ghlUserId);
      }
      
      // If still not found, get any active user for testing
      if (!user) {
        this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
        user = await this.prisma.user.findFirst({
          where: { status: 'ACTIVE' }
        });
        if (!user) {
          throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
        }
      }
    }

    // First try to find survey with matching ghlUserId
    let survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: ghlUserId, // Use the ghlUserId from the request
        isDeleted: false
      }
    });

    // If not found, try to find survey with null ghlUserId (for existing surveys)
    if (!survey) {
      this.logger.log(`Survey not found with ghlUserId: ${ghlUserId}, trying with null ghlUserId`);
      survey = await this.prisma.survey.findFirst({
        where: {
          ghlOpportunityId,
          ghlUserId: null,
          isDeleted: false
        }
      });
      
      // If found with null ghlUserId, update it to have the correct ghlUserId
      if (survey) {
        this.logger.log(`Found survey with null ghlUserId, updating to: ${ghlUserId}`);
        const validGhlUserId = await this.validateAndGetGhlUserId(ghlUserId, user);
        survey = await this.prisma.survey.update({
          where: { id: survey.id },
          data: { ghlUserId: validGhlUserId }
        });
      }
    }

    if (!survey) {
      this.logger.log(`No survey found for opportunity ${ghlOpportunityId}, creating new one`);
      return this.createSurvey(ghlUserId, ghlOpportunityId);
    }

    // Reset all survey data to null/empty
    const resetData: any = {
      page1: null,
      page2: null,
      page3: null,
      page4: null,
      page5: null,
      page6: null,
      page7: null,
      page8: null,
      status: SurveyStatus.DRAFT,
      eligibilityScore: null,
      rejectionReason: null,
      submittedAt: null,
      approvedAt: null,
      rejectedAt: null,
      updatedAt: new Date()
    };

    // Update survey with reset data
    const resetSurvey = await this.prisma.survey.update({
      where: { id: survey.id },
      data: resetData
    });

    // Delete all associated images
    try {
      await this.surveyImageService.deleteAllSurveyImages(survey.id);
      this.logger.log(`Deleted all images for survey ${survey.id}`);
    } catch (error) {
      this.logger.warn(`Failed to delete images for survey ${survey.id}: ${error.message}`);
      // Don't fail the reset if image deletion fails
    }

    this.logger.log(`Reset survey ${survey.id} for opportunity ${ghlOpportunityId}`);
    return this.mapToDto(resetSurvey);
  }

  async getSurveyImages(ghlUserId: string, ghlOpportunityId: string): Promise<any[]> {
    // Try to find user by ghlUserId first, then by internal ID if ghlUserId is actually an internal ID
    let user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // If not found by ghlUserId, try to find by internal ID
      user = await this.userService.findById(ghlUserId);
    }
    
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    return await this.surveyImageService.getSurveyImages(survey.id);
  }

  async getSurveyImagesByField(ghlUserId: string, ghlOpportunityId: string, fieldName: string): Promise<any[]> {
    // Try to find user by ghlUserId first, then by internal ID if ghlUserId is actually an internal ID
    let user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // If not found by ghlUserId, try to find by internal ID
      user = await this.userService.findById(ghlUserId);
    }
    
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    return await this.surveyImageService.getSurveyImagesByField(survey.id, fieldName);
  }

  async getSurveyImageAsBase64(ghlUserId: string, ghlOpportunityId: string, imageId: string): Promise<string | null> {
    // Try to find user by ghlUserId first, then by internal ID if ghlUserId is actually an internal ID
    let user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      // If not found by ghlUserId, try to find by internal ID
      user = await this.userService.findById(ghlUserId);
    }
    
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    return await this.surveyImageService.getImageAsBase64(imageId);
  }

  async sendSurveyEmail(ghlUserId: string, ghlOpportunityId: string, recipientEmail: string): Promise<boolean> {
    const user = await this.userService.findByGhlUserId(ghlUserId);
    if (!user) {
      throw new NotFoundException('User not found');
    }

    const survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: user.ghlUserId,
        isDeleted: false
      }
    });

    if (!survey) {
      throw new NotFoundException('Survey not found');
    }

    try {
      const surveyDto = this.mapToDto(survey);
      
      // Get opportunity details if available
      let opportunityDetails: any = null;
      try {
        // You can add logic here to fetch opportunity details from GHL or your database
        opportunityDetails = {
          name: `Opportunity ${ghlOpportunityId}`,
          ghlOpportunityId,
          createdAt: new Date().toISOString()
        };
      } catch (error) {
        this.logger.warn('Could not fetch opportunity details for email');
      }

      const success = await this.emailService.sendSurveyResponseEmail(surveyDto, recipientEmail, opportunityDetails);
      
      if (success) {
        this.logger.log(`Survey email sent successfully to ${recipientEmail} for opportunity ${ghlOpportunityId}`);
      } else {
        this.logger.error(`Failed to send survey email to ${recipientEmail} for opportunity ${ghlOpportunityId}`);
      }
      
      return success;
    } catch (error) {
      this.logger.error(`Error sending survey email: ${error.message}`);
      return false;
    }
  }



  private async validateAndGetGhlUserId(ghlUserId: string, user: any): Promise<string> {
    // Validate that the ghlUserId exists in the User table
    const userExists = await this.prisma.user.findUnique({
      where: { ghlUserId: ghlUserId }
    });
    
    if (!userExists) {
      this.logger.warn(`User with ghlUserId ${ghlUserId} not found, using user's ghlUserId instead`);
      return user.ghlUserId;
    }
    
    return ghlUserId;
  }

  private mapToDto(survey: any): SurveyResponseDto {
    return {
      id: survey.id,
      ghlOpportunityId: survey.ghlOpportunityId,
      ghlUserId: survey.ghlUserId,
      status: survey.status,
      rejectionReason: survey.rejectionReason,
      page1: survey.page1,
      page2: survey.page2,
      page3: survey.page3,
      page4: survey.page4,
      page5: survey.page5,
      page6: survey.page6,
      page7: survey.page7,
      page8: survey.page8,
      createdAt: survey.createdAt,
      updatedAt: survey.updatedAt,
      submittedAt: survey.submittedAt,
      approvedAt: survey.approvedAt,
      rejectedAt: survey.rejectedAt
    };
  }


  async getSurveyReportPath(opportunityId: string, format: 'html' | 'pdf' = 'html'): Promise<string | null> {
    try {
      return await this.surveyReportService.getReportPath(opportunityId, format);
    } catch (error) {
      this.logger.error(`Failed to get survey report path: ${error.message}`);
      return null;
    }
  }

  async saveSurveyPage(ghlUserId: string, ghlOpportunityId: string, pageData: any, images?: any): Promise<SurveyResponseDto> {
    this.logger.log(`🔧 [SAVE_PAGE] Starting saveSurveyPage for ghlUserId: ${ghlUserId}, ghlOpportunityId: ${ghlOpportunityId}`);
    this.logger.log(`🔧 [SAVE_PAGE] Page data keys: ${Object.keys(pageData || {}).join(', ')}`);
    this.logger.log(`🔧 [SAVE_PAGE] Images provided: ${images ? 'Yes' : 'No'}`);
    this.logger.log(`🔧 [SAVE_PAGE] Raw pageData:`, JSON.stringify(pageData, null, 2));
    
    let user;
    
    // Handle the case where ghlUserId is 'default' (testing mode)
    if (ghlUserId === 'default') {
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException('No active users found for testing');
      }
    } else {
      // First try to find by ghlUserId
      user = await this.userService.findByGhlUserId(ghlUserId);
      
      // If not found, try to find by internal ID
      if (!user) {
        user = await this.userService.findById(ghlUserId);
      }
      
      // If still not found, get any active user for testing
      if (!user) {
        this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
        user = await this.prisma.user.findFirst({
          where: { status: 'ACTIVE' }
        });
        if (!user) {
          throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
        }
      }
    }

    this.logger.log(`🔧 [SAVE_PAGE] Using user: ${user.id}, ghlUserId: ${user.ghlUserId}, name: ${user.name}`);

    // First try to find survey with matching ghlUserId
    this.logger.log(`🔧 [SAVE_PAGE] Looking for survey with ghlOpportunityId: ${ghlOpportunityId}, ghlUserId: ${ghlUserId}`);
    let survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: ghlUserId,
        isDeleted: false
      }
    });
    
    this.logger.log(`🔧 [SAVE_PAGE] Survey found with ghlUserId: ${survey ? 'Yes' : 'No'}`);

    // If not found, try to find survey with null ghlUserId (for existing surveys)
    if (!survey) {
      this.logger.log(`🔧 [SAVE_PAGE] Survey not found with ghlUserId: ${ghlUserId}, trying with null ghlUserId`);
      survey = await this.prisma.survey.findFirst({
        where: {
          ghlOpportunityId,
          ghlUserId: null,
          isDeleted: false
        }
      });
      
      this.logger.log(`🔧 [SAVE_PAGE] Survey found with null ghlUserId: ${survey ? 'Yes' : 'No'}`);
      
      // If found with null ghlUserId, update it to have the correct ghlUserId
      if (survey) {
        this.logger.log(`🔧 [SAVE_PAGE] Found survey with null ghlUserId, updating to: ${ghlUserId}`);
        const validGhlUserId = await this.validateAndGetGhlUserId(ghlUserId, user);
        survey = await this.prisma.survey.update({
          where: { id: survey.id },
          data: { ghlUserId: validGhlUserId }
        });
      }
    }

    if (!survey) {
      this.logger.log(`🔧 [SAVE_PAGE] No survey found for opportunity ${ghlOpportunityId}, creating new one`);
      return this.createSurvey(ghlUserId, ghlOpportunityId);
    }

    this.logger.log(`🔧 [SAVE_PAGE] Updating survey ${survey.id} with page data`);
    this.logger.log(`🔧 [SAVE_PAGE] Page data to save:`, JSON.stringify(pageData, null, 2));

    // Determine which page this data belongs to based on the data structure
    let targetPage: string | null = null;
    let actualPageData: any = null;
    
    // Check if pageData contains page-specific fields (properly wrapped)
    if (pageData.page1) {
      targetPage = 'page1';
      actualPageData = pageData.page1;
    } else if (pageData.page2) {
      targetPage = 'page2';
      actualPageData = pageData.page2;
    } else if (pageData.page3) {
      targetPage = 'page3';
      actualPageData = pageData.page3;
    } else if (pageData.page4) {
      targetPage = 'page4';
      actualPageData = pageData.page4;
    } else if (pageData.page5) {
      targetPage = 'page5';
      actualPageData = pageData.page5;
    } else if (pageData.page6) {
      targetPage = 'page6';
      actualPageData = pageData.page6;
    } else if (pageData.page7) {
      targetPage = 'page7';
      actualPageData = pageData.page7;
    } else if (pageData.page8) {
      targetPage = 'page8';
      actualPageData = pageData.page8;
    } else {
      // If no page field is specified, try to determine from the data content
      // This is a fallback for when the frontend sends data without page wrapper
      if (pageData.selectedReasons || pageData.reasons) {
        targetPage = 'page2'; // Assuming selectedReasons belongs to page2
        actualPageData = pageData;
      } else {
        // Default to page1 if we can't determine
        targetPage = 'page1';
        actualPageData = pageData;
      }
    }

    this.logger.log(`🔧 [SAVE_PAGE] Determined target page: ${targetPage}`);
    this.logger.log(`🔧 [SAVE_PAGE] Actual page data:`, JSON.stringify(actualPageData, null, 2));

    // Prepare update data
    const updateData: any = {
      updatedAt: new Date()
    };

    if (targetPage && actualPageData) {
      updateData[targetPage] = actualPageData;
    }

    this.logger.log(`🔧 [SAVE_PAGE] Final update data:`, JSON.stringify(updateData, null, 2));

    // Update survey with page data
    const updatedSurvey = await this.prisma.survey.update({
      where: { id: survey.id },
      data: updateData
    });

    this.logger.log(`✅ [SAVE_PAGE] Successfully saved survey page for opportunity ${ghlOpportunityId}, survey ID: ${updatedSurvey.id}`);
    return this.mapToDto(updatedSurvey);
  }

  async uploadImagesAndGetUrls(ghlUserId: string, ghlOpportunityId: string, fieldName: string, images: any[]): Promise<{ urls: string[] }> {
    this.logger.log(`📷 [UPLOAD_IMAGES] Starting uploadImagesAndGetUrls for ghlUserId: ${ghlUserId}, ghlOpportunityId: ${ghlOpportunityId}`);
    this.logger.log(`📷 [UPLOAD_IMAGES] Field name: ${fieldName}, Number of images: ${images.length}`);
    this.logger.log(`📷 [UPLOAD_IMAGES] Image details:`, images.map((img, index) => ({
      index,
      name: img.name,
      mimeType: img.mimeType,
      size: img.size,
      hasBase64Data: !!img.base64Data
    })));
    
    let user;
    
    // Handle the case where ghlUserId is 'default' (testing mode)
    if (ghlUserId === 'default') {
      user = await this.prisma.user.findFirst({
        where: { status: 'ACTIVE' }
      });
      if (!user) {
        throw new NotFoundException('No active users found for testing');
      }
    } else {
      // First try to find by ghlUserId
      user = await this.userService.findByGhlUserId(ghlUserId);
      
      // If not found, try to find by internal ID
      if (!user) {
        user = await this.userService.findById(ghlUserId);
      }
      
      // If still not found, get any active user for testing
      if (!user) {
        this.logger.warn(`User not found with ghlUserId: ${ghlUserId}, using fallback user for testing`);
        user = await this.prisma.user.findFirst({
          where: { status: 'ACTIVE' }
        });
        if (!user) {
          throw new NotFoundException(`User not found with ghlUserId: ${ghlUserId} and no fallback user available`);
        }
      }
    }

    // Find or create survey
    this.logger.log(`📷 [UPLOAD_IMAGES] Looking for survey with ghlOpportunityId: ${ghlOpportunityId}, ghlUserId: ${ghlUserId}`);
    let survey = await this.prisma.survey.findFirst({
      where: {
        ghlOpportunityId,
        ghlUserId: ghlUserId,
        isDeleted: false
      }
    });

    this.logger.log(`📷 [UPLOAD_IMAGES] Survey found: ${survey ? 'Yes' : 'No'}`);

    if (!survey) {
      this.logger.log(`📷 [UPLOAD_IMAGES] No survey found for opportunity ${ghlOpportunityId}, creating new one`);
      const newSurvey = await this.createSurvey(ghlUserId, ghlOpportunityId);
      survey = await this.prisma.survey.findUnique({
        where: { id: newSurvey.id }
      });
      this.logger.log(`📷 [UPLOAD_IMAGES] Created new survey with ID: ${survey?.id}`);
    }

    if (!survey) {
      this.logger.error(`📷 [UPLOAD_IMAGES] Failed to create or find survey for opportunity ${ghlOpportunityId}`);
      throw new NotFoundException('Failed to create or find survey');
    }

    this.logger.log(`📷 [UPLOAD_IMAGES] Using survey ID: ${survey.id}`);
    const uploadedUrls: string[] = [];

    // Upload each image/file to Cloudinary
    for (let index = 0; index < images.length; index++) {
      const imageData = images[index];
      
      this.logger.log(`📷 [UPLOAD_IMAGES] Processing file ${index + 1}/${images.length} for field ${fieldName}`);
      
      try {
        // Determine resource type based on file mime type
        // Upload PDFs as 'image' resource type to avoid Cloudinary free account restrictions
        // Cloudinary supports PDFs as images and will preserve the file format
        const isPdf = imageData.mimeType === 'application/pdf';
        const resourceType = 'image'; // Always use 'image' for both images and PDFs
        
        // Get file extension from filename
        const fileName = imageData.name || `${fieldName}-${index}`;
        const fileExtension = fileName.split('.').pop()?.toLowerCase() || '';
        
        // Create unique public ID for Cloudinary (without folder path - folder is passed separately)
        // Preserve extension for PDFs so Cloudinary knows it's a PDF
        let publicId = `${fieldName}-${index}-${Date.now()}`;
        if (isPdf && fileExtension === 'pdf') {
          publicId = `${publicId}.pdf`;
        }
        
        this.logger.log(`📷 [UPLOAD_IMAGES] Public ID: ${publicId}`);
        
        // Prepare upload options
        const uploadOptions: any = {
          resource_type: resourceType,
        };
        
        // Set format only for PDFs (Cloudinary doesn't accept 'auto' for format parameter)
        if (isPdf) {
          uploadOptions.format = 'pdf';
        } else {
          // For images, use fetch_format and quality for automatic optimization
          uploadOptions.quality = 'auto';
          uploadOptions.fetch_format = 'auto';
        }
        
        // Upload base64 file to Cloudinary
        this.logger.log(`📷 [UPLOAD_IMAGES] Uploading to Cloudinary...`);
        const uploadResult = await this.cloudinaryService.uploadBase64Image(
          imageData.base64Data,
          `surveys/${ghlOpportunityId}`,
          publicId,
          uploadOptions
        );
        
        this.logger.log(`📷 [UPLOAD_IMAGES] Cloudinary upload successful: ${uploadResult.secure_url}`);
        
        // For PDFs, generate a signed URL to bypass "untrusted customer" restrictions
        let fileUrl = uploadResult.secure_url;
        if (isPdf) {
          try {
            // Generate signed URL for PDF to ensure it's accessible
            // Use the public_id from the upload result (which includes folder path)
            const fullPublicId = uploadResult.public_id || `surveys/${ghlOpportunityId}/${publicId}`;
            fileUrl = this.cloudinaryService.generateSignedUrl(fullPublicId, 'image', { format: 'pdf' });
            this.logger.log(`📄 [UPLOAD_IMAGES] Generated signed URL for PDF: ${fileUrl}`);
          } catch (error) {
            this.logger.warn(`⚠️ [UPLOAD_IMAGES] Failed to generate signed URL for PDF, using regular URL: ${error.message}`);
            // Fallback to regular URL if signed URL generation fails
            fileUrl = uploadResult.secure_url;
          }
        }
        
        // Store the Cloudinary URL (signed URL for PDFs)
        uploadedUrls.push(fileUrl);
        
        // Save image record to database with Cloudinary URL (signed URL for PDFs)
        this.logger.log(`📷 [UPLOAD_IMAGES] Saving image record to database...`);
        await this.prisma.surveyImage.create({
          data: {
            surveyId: survey.id,
            fieldName: fieldName,
            fileName: imageData.name || `${fieldName}-${index}`,
            originalName: imageData.name,
            mimeType: imageData.mimeType,
            fileSize: imageData.size || 0,
            filePath: fileUrl, // Store signed URL for PDFs, regular URL for images
            base64Data: null
          }
        });
        
        this.logger.log(`✅ [UPLOAD_IMAGES] Successfully uploaded ${fieldName}-${index} to Cloudinary: ${uploadResult.secure_url}`);
      } catch (error) {
        this.logger.error(`❌ [UPLOAD_IMAGES] Error uploading ${fieldName}-${index} to Cloudinary:`, error);
        this.logger.error(`❌ [UPLOAD_IMAGES] Error details:`, {
          fieldName,
          index,
          imageName: imageData.name,
          mimeType: imageData.mimeType,
          hasBase64Data: !!imageData.base64Data,
          base64Length: imageData.base64Data?.length || 0,
          error: error.message
        });
        // Continue with other images even if one fails
      }
    }

    this.logger.log(`✅ [UPLOAD_IMAGES] Successfully uploaded ${uploadedUrls.length}/${images.length} images for field ${fieldName}`);
    this.logger.log(`📷 [UPLOAD_IMAGES] Returning URLs:`, uploadedUrls);
    return { urls: uploadedUrls };
  }
} 