import { Controller, Post, Body, HttpException, HttpStatus } from '@nestjs/common';
import { EmailService } from './email.service';

export interface SendWelcomeEmailDto {
  customerName: string;
  customerEmail: string;
  ghlOpportunityId?: string;
  ghlUserId?: string;
}

@Controller('email')
export class EmailController {
  constructor(private readonly emailService: EmailService) {}

  @Post('send-welcome')
  async sendWelcomeEmail(@Body() sendWelcomeEmailDto: SendWelcomeEmailDto) {
    try {
      const { customerName, customerEmail, ghlOpportunityId, ghlUserId } = sendWelcomeEmailDto;

      if (!customerEmail || !customerName) {
        throw new HttpException('Customer name and email are required', HttpStatus.BAD_REQUEST);
      }

      const result = await this.emailService.sendWelcomeEmailOutlook(customerName, customerEmail);

      if (result.success) {
        // If opportunity ID and user ID are provided, complete the welcome email step
        if (ghlOpportunityId && ghlUserId) {
          try {
            // Make internal API call to complete the step using fetch
            const stepCompletionPayload = {
              stepNumber: 12, // WELCOME_EMAIL step
              data: {
                emailSent: true,
                customerEmail: customerEmail,
                customerName: customerName,
                sentAt: new Date().toISOString()
              }
            };

            // Get the base URL for internal API calls
            const baseUrl = process.env.API_BASE_URL || 'http://localhost:3000';
            const completeStepUrl = `${baseUrl}/opportunity-workflow/progress/${ghlOpportunityId}/complete-step`;

            // Make the API call to complete the step using fetch
            const stepResponse = await fetch(completeStepUrl, {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
              },
              body: JSON.stringify(stepCompletionPayload)
            });

            if (stepResponse.ok) {
              const stepData = await stepResponse.json();
              console.log('Step completion response:', stepData);
            } else {
              console.error('Step completion failed with status:', stepResponse.status);
            }
          } catch (stepError) {
            console.error('Failed to complete welcome email step:', stepError);
            // Don't fail the email sending if step completion fails
          }
        }

        return {
          success: true,
          message: result.message,
          data: {
            to: customerEmail,
            from: 'support@creativuk.co.uk',
            subject: 'Hi, and welcome to your new energy future.',
            sentAt: new Date().toISOString(),
            stepCompleted: !!(ghlOpportunityId && ghlUserId)
          }
        };
      } else {
        throw new HttpException(
          {
            message: result.message,
            error: result.error,
          },
          HttpStatus.INTERNAL_SERVER_ERROR,
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }
      
      throw new HttpException(
        {
          message: 'Failed to send welcome email',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  @Post('send-custom')
  async sendCustomEmail(@Body() emailData: { to: string; subject: string; body: string; attachments?: string[]; fromEmail?: string }) {
    try {
      const result = await this.emailService.sendOutlookEmail(emailData);

      if (result.success) {
        return {
          success: true,
          message: result.message,
          data: {
            to: emailData.to,
            from: emailData.fromEmail || 'support@creativuk.co.uk',
            subject: emailData.subject,
            sentAt: new Date().toISOString()
          }
        };
      } else {
        throw new HttpException(
          {
            message: result.message,
            error: result.error,
          },
          HttpStatus.INTERNAL_SERVER_ERROR,
        );
      }
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }
      
      throw new HttpException(
        {
          message: 'Failed to send email',
          error: error.message,
        },
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }
}
