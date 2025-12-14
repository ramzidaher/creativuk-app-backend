import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import * as fs from 'fs/promises';
import * as path from 'path';

@Injectable()
export class SignaturesService {
  private readonly logger = new Logger(SignaturesService.name);

  constructor(private prisma: PrismaService) {}

  async saveSignature(
    opportunityId: string,
    signatureData: string,
    metadata: {
      signedAt: string;
      signedBy?: string;
      ipAddress?: string;
      userAgent?: string;
    }
  ) {
    try {
      this.logger.log(`Saving signature for opportunity: ${opportunityId}`);

      // Create signatures directory if it doesn't exist
      const signaturesDir = path.join(process.cwd(), 'src', 'signatures');
      await fs.mkdir(signaturesDir, { recursive: true });

      // Save signature data to file
      const signatureFileName = `signature-${opportunityId}-${Date.now()}.json`;
      const signatureFilePath = path.join(signaturesDir, signatureFileName);

      const signatureRecord = {
        opportunityId,
        signatureData,
        metadata,
        createdAt: new Date().toISOString(),
        fileName: signatureFileName
      };

      await fs.writeFile(signatureFilePath, JSON.stringify(signatureRecord, null, 2));

      // Save to database
      const signature = await this.prisma.signature.create({
        data: {
          opportunityId,
          signatureData,
          signedAt: new Date(metadata.signedAt),
          signedBy: metadata.signedBy,
          ipAddress: metadata.ipAddress,
          userAgent: metadata.userAgent,
          filePath: signatureFilePath
        }
      });

      this.logger.log(`Signature saved successfully for opportunity: ${opportunityId}`);
      return signature;
    } catch (error) {
      this.logger.error(`Error saving signature: ${error.message}`);
      throw error;
    }
  }

  async getSignature(opportunityId: string) {
    try {
      const signature = await this.prisma.signature.findFirst({
        where: { opportunityId },
        orderBy: { createdAt: 'desc' }
      });

      return signature;
    } catch (error) {
      this.logger.error(`Error retrieving signature: ${error.message}`);
      throw error;
    }
  }

  async getSignaturesByOpportunity(opportunityId: string) {
    try {
      const signatures = await this.prisma.signature.findMany({
        where: { opportunityId },
        orderBy: { createdAt: 'desc' }
      });

      return signatures;
    } catch (error) {
      this.logger.error(`Error retrieving signatures: ${error.message}`);
      throw error;
    }
  }
}
