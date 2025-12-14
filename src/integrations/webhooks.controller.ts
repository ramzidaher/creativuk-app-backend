import { Controller, Post, Body, Headers, HttpCode, HttpStatus, Logger } from '@nestjs/common';
import { DocuSealService } from './docuseal.service';

/**
 * Webhook controller for external service integrations
 * Handles webhooks from DocuSeal and other services
 * 
 * Routes:
 * - POST /hooks - Main webhook endpoint (reverse proxy strips /api prefix)
 * Note: DocuSeal is configured with https://creativuk-app.paldev.tech/api/hooks
 * but reverse proxy forwards /hooks to this controller
 */
@Controller('hooks')
export class WebhooksController {
  private readonly logger = new Logger(WebhooksController.name);

  constructor(
    private readonly docuSealService: DocuSealService
  ) {}

  /**
   * Main webhook endpoint for DocuSeal
   * POST /api/hooks
   * 
   * Receives real-time notifications from DocuSeal about document events
   * No authentication required - uses webhook secret for verification
   * 
   * DocuSeal webhook payload structure:
   * {
   *   "event_type": "form.completed" | "submission.completed" | etc.,
   *   "timestamp": "2025-12-01T22:47:17Z",
   *   "data": {
   *     "id": 616343,  // submitter ID
   *     "submission": { "id": 473508 },  // submission ID
   *     "template": { "external_id": "template-opportunityId" },
   *     "status": "completed",
   *     ...
   *   }
   * }
   */
  @Post()
  @HttpCode(HttpStatus.OK)
  async handleWebhook(
    @Body() payload: any,
    @Headers() headers: Record<string, string>
  ) {
    try {
      const eventType = payload.event_type || payload.event || 'unknown';
      this.logger.log(`📥 Received webhook event: ${eventType}`);
      this.logger.debug(`Webhook payload: ${JSON.stringify(payload, null, 2)}`);

      // Verify webhook secret if configured
      const secretVerified = await this.docuSealService.verifyWebhookSecret(headers, payload);
      if (!secretVerified) {
        this.logger.warn('⚠️ Webhook secret verification failed - request may not be from DocuSeal');
        // Continue processing but log warning
        // In production, you might want to reject here: throw new HttpException('Unauthorized', HttpStatus.UNAUTHORIZED);
      }

      // Skip template events - they don't have submissions
      if (eventType === 'template.created' || eventType === 'template.updated') {
        this.logger.log(`ℹ️ Skipping template event: ${eventType} (no submission to process)`);
        return {
          success: true,
          message: `Template event ${eventType} received and skipped`,
        };
      }

      // Extract submission ID from DocuSeal payload structure
      // DocuSeal sends: { "data": { "id": 473788, ... } } for submissions
      const submissionId = payload.data?.id?.toString() || 
                          payload.data?.submission?.id?.toString() ||
                          payload.submission?.id?.toString() ||
                          payload.submission_id?.toString();

      if (!submissionId) {
        this.logger.warn('⚠️ Webhook payload missing submission ID');
        this.logger.debug(`Available payload keys: ${Object.keys(payload).join(', ')}`);
        return {
          success: true,
          message: 'Webhook received but no submission ID found',
        };
      }

      this.logger.log(`Processing webhook event: ${eventType} for submission: ${submissionId}`);

      // Process the webhook event with the actual DocuSeal payload structure
      await this.docuSealService.handleWebhookEvent(eventType, payload, submissionId);

      return {
        success: true,
        message: 'Webhook processed successfully',
      };
    } catch (error) {
      this.logger.error(`❌ Error processing webhook: ${error.message}`);
      this.logger.error(`Stack trace: ${error.stack}`);
      
      // Return 200 to prevent DocuSeal from retrying
      // Log the error for manual investigation
      return {
        success: false,
        message: 'Webhook processing error (logged for investigation)',
        error: error.message,
      };
    }
  }
}

