import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
const nodemailer = require('nodemailer');
import { SurveyResponseDto } from '../opportunities/dto/survey.dto';
import { join } from 'path';

@Injectable()
export class EmailService {
  private readonly logger = new Logger(EmailService.name);
  private transporter: any;

  constructor(private readonly configService: ConfigService) {
    this.initializeTransporter();
  }

  private initializeTransporter() {
    // For development, use Gmail SMTP or a service like Mailtrap
    const emailConfig = {
      host: this.configService.get('SMTP_HOST', 'smtp.gmail.com'),
      port: this.configService.get('SMTP_PORT', 587),
      secure: false, // true for 465, false for other ports
      auth: {
        user: this.configService.get('SMTP_USER'),
        pass: this.configService.get('SMTP_PASS'),
      },
    };

    this.transporter = nodemailer.createTransport(emailConfig);
  }

  async sendWelcomeEmailOutlook(customerName: string, customerEmail: string): Promise<{ success: boolean; message: string; error?: string }> {
    const welcomeSubject = 'Welcome to your new energy future!';
    const welcomeBody = `
      <p>Dear ${customerName},</p>
      
      <p>Welcome to your new energy future! Thank you for choosing Creativ Energy for your solar energy installation. The Creativ team is here to make your switch to renewable energy a smooth and straightforward process, and we're committed to delivering 100% satisfaction.</p>
      
      <p>If we haven't already completed one, we may need to carry out an installation inspection survey. A member of our team will contact you to arrange this if necessary.</p>
      
      <p>We'll now begin preparing all the relevant documentation for your installation. If we require any further information from you, we'll be in touch.</p>
      
      <p>For now, you can sit back and relax—we'll take care of everything. If you have any questions or need assistance, feel free to email us using <a href="mailto:operations@creativuk.co.uk">operations@creativuk.co.uk</a> or call us on 01733 515028.</p>
      
      <p>To get the most out of your solar system, we recommend switching to Octopus Energy if you're not already a customer. They offer some of the best export and grid trading rates in the industry.</p>
      
      <p>If you don't already have a smart meter, you'll need one to benefit from export tariffs. These are available for free from your current energy supplier, or can be arranged when you switch to Octopus.</p>
      
      <p><strong>Next steps:</strong></p>
      <ul>
        <li>We'll complete your DNO (Distribution Network Operator) application</li>
        <li>Arrange the scaffolding</li>
        <li>If required, schedule a pre-installation survey with one of our engineers</li>
        <li>One of our team members will call you soon to walk you through the rest of the process</li>
      </ul>
      
      <p><strong>Helpful Links:</strong></p>
      <ul>
        <li>Join Octopus Energy and get £50 for signing up:<br>
        <a href="https://share.octopus.energy/taupe-cat-521">https://share.octopus.energy/taupe-cat-521</a></li>
        <li>Double Rewards: Earn £200 for yourself and an extra £200 for anyone you refer to us:<br>
        <a href="https://solar.creativuk.co.uk/referral-form">https://solar.creativuk.co.uk/referral-form</a></li>
        <li>Leave us a review and let us know how we're doing:<br>
        <a href="https://www.google.com/maps/place//data=!4m3!3m2!1s0x4877f108fa05bb99:0x5931e9a7b8342fae!12e1?source=g.page.m.rc._&laa=merchant-web-dashboard-card">https://g.page/r/Ca4vNLin6TFZEAg/review</a></li>
        <li>EcoFlow Brochure</li>
        <li><a href="https://jarmqltd-my.sharepoint.com/personal/karl_gedney_creativuk_co_uk/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fkarl%5Fgedney%5Fcreativuk%5Fco%5Fuk%2FDocuments%2FApp%20info%2FPowerOcean%20%28single%2Dphase%29%5FBrochure%5F20240527%5FEN%5FView%2D1%2Epdf&parent=%2Fpersonal%2Fkarl%5Fgedney%5Fcreativuk%5Fco%5Fuk%2FDocuments%2FApp%20info&ga=1">PowerOcean (single-phase) Brochure</a></li>
      </ul>
      
      <p><strong>Please Note:</strong></p>
      <p>Your refundable holding payment will be returned to you once your installation is complete. However, this payment becomes non-refundable if you cancel your order at any point after giving express consent for us to begin work within your cancellation period.</p>
      
      <p>Thank you again for choosing Creativ Energy. We look forward to making your solar journey a great experience.</p>
    `;

    const attachmentPath = 'C:\\Users\\\Creativuk\\creativ-solar-app\\apps\\backend\\src\\Creativ-welcome[1].docx';
    
    return this.sendOutlookEmail({
      to: customerEmail,
      subject: welcomeSubject,
      body: welcomeBody,
      attachments: [attachmentPath],
      fromEmail: 'support@creativuk.co.uk'
    });
  }

  async sendOutlookEmail(emailData: { to: string; subject: string; body: string; attachments?: string[]; fromEmail?: string }): Promise<{ success: boolean; message: string; error?: string }> {
    const { to, subject, body, attachments = [], fromEmail = 'support@creativuk.co.uk' } = emailData;

    this.logger.log(`📧 Sending Outlook COM email to: ${to}`);
    this.logger.log(`📧 From: ${fromEmail}`);
    this.logger.log(`📧 Subject: ${subject}`);
    this.logger.log(`📧 Attachments: ${attachments.length}`);

    // Validate attachment files exist
    const fs = require('fs');
    for (const attachment of attachments) {
      if (!fs.existsSync(attachment)) {
        const error = `Attachment file not found: ${attachment}`;
        this.logger.error(error);
        return { success: false, message: error };
      }
    }

    // Create PowerShell script for Outlook COM email sending
    const script = this.createOutlookEmailScript(to, subject, body, attachments, fromEmail);

    try {
      const fs = require('fs');
      const path = require('path');
      const { exec } = require('child_process');
      const { promisify } = require('util');
      const execAsync = promisify(exec);

      // Create temporary script file
      const scriptPath = path.join(process.cwd(), `temp-email-${Date.now()}.ps1`);
      fs.writeFileSync(scriptPath, script, 'utf-8');

      this.logger.log(`📧 Executing email script: ${scriptPath}`);

      // Execute PowerShell script
      const { stdout, stderr } = await execAsync(`powershell -ExecutionPolicy Bypass -File "${scriptPath}"`, { 
        timeout: 60000 // 60 second timeout
      });

      // Clean up script file
      try {
        fs.unlinkSync(scriptPath);
      } catch (cleanupError) {
        this.logger.warn(`Failed to cleanup script file: ${cleanupError.message}`);
      }

      if (stderr) {
        this.logger.error(`PowerShell stderr: ${stderr}`);
      }

      this.logger.log(`📧 Email script output: ${stdout}`);

      // Check if the script executed successfully
      if (stdout.includes('SUCCESS: Email sent')) {
        return {
          success: true,
          message: 'Email sent successfully'
        };
      } else if (stdout.includes('ERROR:')) {
        const errorMatch = stdout.match(/ERROR: (.+)/);
        const error = errorMatch ? errorMatch[1] : 'Unknown error occurred';
        return {
          success: false,
          message: 'Failed to send email',
          error: error
        };
      } else {
        return {
          success: false,
          message: 'Email sending status unclear',
          error: stdout
        };
      }
    } catch (error) {
      this.logger.error(`📧 Email sending failed: ${error.message}`);
      return {
        success: false,
        message: 'Email sending failed',
        error: error.message
      };
    }
  }

  private createOutlookEmailScript(to: string, subject: string, body: string, attachments: string[], fromEmail: string): string {
    // Escape special characters for PowerShell (minimal escaping needed for here-string)
    const escapedSubject = subject.replace(/"/g, '""');
    const escapedTo = to.replace(/"/g, '""');
    const escapedFromEmail = fromEmail.replace(/"/g, '""');

    // BCC recipients
    const bccRecipients = [
      'josh.beresford@creativuk.co.uk',
      'Kemberly.willocks@creativuk.co.uk', 
      'michaela.ferguson@creativuk.co.uk',
      'pamela.rennie@creativuk.co.uk'
    ];
    const escapedBcc = bccRecipients.map(email => email.replace(/"/g, '""')).join(';');

    const attachmentCode = attachments.map((att, i) => 
      `    $mail.Attachments.Add("${att.replace(/\\/g, '\\\\').replace(/"/g, '""')}")
    Write-Host "Added attachment: ${att}" -ForegroundColor Green`
    ).join('\n');

    return `# Outlook COM Email Script - Generated at ${new Date().toISOString()}
$ErrorActionPreference = "Stop"
Write-Host "Starting email automation..." -ForegroundColor Green
Write-Host "To: ${escapedTo}" -ForegroundColor Yellow
Write-Host "From: ${escapedFromEmail}" -ForegroundColor Yellow
Write-Host "BCC: ${escapedBcc}" -ForegroundColor Yellow

$outlook = $null
$mail = $null

# HTML Body content using here-string to handle special characters
$htmlBody = @"
${body}
"@

try {
    Write-Host "Creating Outlook application..." -ForegroundColor Green
    $outlook = New-Object -ComObject Outlook.Application
    
    Write-Host "Creating mail item..." -ForegroundColor Green
    $mail = $outlook.CreateItem(0)  # olMailItem
    
    Write-Host "Setting email properties..." -ForegroundColor Green
    $mail.To = "${escapedTo}"
    $mail.Subject = "${escapedSubject}"
    $mail.HTMLBody = $htmlBody
    $mail.BodyFormat = 2  # olFormatHTML
    
    # Add BCC recipients
    Write-Host "Adding BCC recipients..." -ForegroundColor Green
    $mail.BCC = "${escapedBcc}"
    Write-Host "BCC recipients added: ${escapedBcc}" -ForegroundColor Green
    
    # Set the From address directly (for shared mailbox or send-as permissions)
    Write-Host "Setting From address to: ${escapedFromEmail}" -ForegroundColor Yellow
    try {
        # Method 1: Try SenderEmailAddress
        $mail.SenderEmailAddress = "${escapedFromEmail}"
        Write-Host "SUCCESS: Set SenderEmailAddress to ${escapedFromEmail}" -ForegroundColor Green
    } catch {
        Write-Host "SenderEmailAddress failed, trying SentOnBehalfOfName..." -ForegroundColor Yellow
        try {
            # Method 2: Try SentOnBehalfOfName
            $mail.SentOnBehalfOfName = "${escapedFromEmail}"
            Write-Host "SUCCESS: Set SentOnBehalfOfName to ${escapedFromEmail}" -ForegroundColor Green
        } catch {
            Write-Host "SentOnBehalfOfName failed, trying SenderName..." -ForegroundColor Yellow
            try {
                # Method 3: Try SenderName
                $mail.SenderName = "${escapedFromEmail}"
                Write-Host "SUCCESS: Set SenderName to ${escapedFromEmail}" -ForegroundColor Green
            } catch {
                Write-Host "WARNING: All From address methods failed, using default sender" -ForegroundColor Red
            }
        }
    }
    
    # Add attachments
    if (${attachments.length} -gt 0) {
        Write-Host "Adding attachments..." -ForegroundColor Green
${attachmentCode}
    }
    
    Write-Host "Sending email..." -ForegroundColor Green
    $mail.Send()
    Write-Host "SUCCESS: Email sent successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.ToString())" -ForegroundColor Red
    exit 1
} finally {
    Write-Host "Cleaning up..." -ForegroundColor Yellow
    if ($mail) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null }
    if ($outlook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    [System.GC]::Collect()
    Write-Host "Cleanup completed" -ForegroundColor Green
}`;
  }

  // Keep existing survey email methods for compatibility
  async sendSurveyResponseEmail(
    survey: SurveyResponseDto,
    recipientEmail: string,
    opportunityDetails?: any,
    uploadedFiles?: any
  ): Promise<boolean> {
    try {
      this.logger.log(`Processing survey email for opportunity ${survey.ghlOpportunityId}`);
      // Survey email implementation would go here
      return true;
    } catch (error) {
      this.logger.error(`Failed to send survey email: ${error.message}`);
      return false;
    }
  }
}