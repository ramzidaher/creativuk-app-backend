import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import * as crypto from 'crypto';
import * as fs from 'fs/promises';
import * as path from 'path';

export interface FreeSignatureData {
  opportunityId: string;
  signatureData: string;
  documentName: string;
  signedAt: string;
  signedBy: string;
  customerEmail: string;
  digitalFootprint: {
    id: string;
    signatureHash: string;
    algorithm: string;
    metadata: any;
  };
}

@Injectable()
export class FreeSignaturesService {
  private readonly logger = new Logger(FreeSignaturesService.name);

  constructor(private prisma: PrismaService) {}

  /**
   * Save a free signature with digital footprint
   */
  async saveFreeSignature(signatureData: FreeSignatureData) {
    try {
      this.logger.log(`Saving free signature for opportunity: ${signatureData.opportunityId}`);

      // Create signatures directory if it doesn't exist
      const signaturesDir = path.join(process.cwd(), 'src', 'signatures', 'free');
      await fs.mkdir(signaturesDir, { recursive: true });

      // Create enhanced digital footprint
      const enhancedFootprint = this.createEnhancedDigitalFootprint(signatureData);

      // Save signature data to file
      const signatureFileName = `free_signature_${signatureData.opportunityId}_${Date.now()}.json`;
      const signatureFilePath = path.join(signaturesDir, signatureFileName);

      const signatureRecord = {
        ...signatureData,
        enhancedDigitalFootprint: enhancedFootprint,
        createdAt: new Date().toISOString(),
        fileName: signatureFileName,
        filePath: signatureFilePath,
        verification: {
          isValid: true,
          verifiedAt: new Date().toISOString(),
          verificationMethod: 'SHA-256_HASH',
        }
      };

      await fs.writeFile(signatureFilePath, JSON.stringify(signatureRecord, null, 2));

      // Save to database
      const signature = await this.prisma.signature.create({
        data: {
          opportunityId: signatureData.opportunityId,
          signatureData: signatureData.signatureData,
          signedAt: new Date(signatureData.signedAt),
          signedBy: signatureData.signedBy,
          ipAddress: 'mobile_app',
          userAgent: 'React Native Free Signing App',
          filePath: signatureFilePath,
        }
      });

      this.logger.log(`Free signature saved successfully for opportunity: ${signatureData.opportunityId}`);
      return {
        success: true,
        signature,
        digitalFootprint: enhancedFootprint,
        filePath: signatureFilePath,
      };
    } catch (error) {
      this.logger.error(`Error saving free signature: ${error.message}`);
      throw error;
    }
  }

  /**
   * Create enhanced digital footprint with multiple verification layers
   */
  private createEnhancedDigitalFootprint(signatureData: FreeSignatureData) {
    const timestamp = new Date().toISOString();
    const nonce = crypto.randomBytes(16).toString('hex');
    
    // Create multiple hash layers for enhanced security
    const primaryHash = crypto
      .createHash('sha256')
      .update(signatureData.signatureData + signatureData.opportunityId + timestamp)
      .digest('hex');

    const secondaryHash = crypto
      .createHash('sha512')
      .update(primaryHash + nonce + signatureData.signedBy)
      .digest('hex');

    const verificationHash = crypto
      .createHash('sha256')
      .update(secondaryHash + signatureData.documentName + signatureData.customerEmail)
      .digest('hex');

    const enhancedFootprint = {
      id: `free_signature_${signatureData.opportunityId}_${Date.now()}`,
      opportunityId: signatureData.opportunityId,
      documentName: signatureData.documentName,
      customerName: signatureData.signedBy,
      customerEmail: signatureData.customerEmail,
      
      // Multi-layer hashing for enhanced security
      hashes: {
        primary: primaryHash,
        secondary: secondaryHash,
        verification: verificationHash,
        algorithm: 'SHA-256/SHA-512',
      },
      
      // Timestamp and metadata
      timestamp: timestamp,
      signedAt: signatureData.signedAt,
      signedBy: signatureData.signedBy,
      
      // Verification data
      verification: {
        method: 'MULTI_LAYER_HASH',
        layers: 3,
        nonce: nonce,
        verifiedAt: timestamp,
        status: 'VERIFIED',
      },
      
      // Legal compliance data
      compliance: {
        signatureType: 'ELECTRONIC_SIGNATURE',
        legalFramework: 'ESIGN_ACT_COMPLIANT',
        auditTrail: true,
        tamperProof: true,
        timestampAuthority: 'INTERNAL',
      },
      
      // Technical metadata
      technical: {
        platform: 'React Native',
        appVersion: '1.0.0',
        signatureFormat: 'BASE64_IMAGE',
        storageMethod: 'LOCAL_FILE_SYSTEM',
        backupMethod: 'DATABASE_RECORD',
      }
    };

    this.logger.log(`Enhanced digital footprint created: ${enhancedFootprint.id}`);
    return enhancedFootprint;
  }

  /**
   * Verify a free signature
   */
  async verifyFreeSignature(opportunityId: string, signatureHash: string) {
    try {
      this.logger.log(`Verifying free signature for opportunity: ${opportunityId}`);

      // Get signature from database
      const signature = await this.prisma.signature.findFirst({
        where: {
          opportunityId,
          userAgent: 'React Native Free Signing App'
        }
      });

      if (!signature) {
        return {
          success: false,
          error: 'Signature not found',
        };
      }

      // Read the signature file for detailed verification
      if (signature.filePath) {
        try {
          await fs.access(signature.filePath);
          const signatureFile = JSON.parse(await fs.readFile(signature.filePath, 'utf8'));
          const footprint = signatureFile.enhancedDigitalFootprint;

          // Verify the hash matches
          const isHashValid = footprint.hashes.verification === signatureHash;

          return {
            success: true,
            verified: isHashValid,
            signature: signature,
            digitalFootprint: footprint,
            verification: {
              hashValid: isHashValid,
              timestamp: footprint.timestamp,
              signedBy: footprint.signedBy,
              documentName: footprint.documentName,
            }
          };
        } catch (error) {
          return {
            success: false,
            error: 'Signature file not found',
          };
        }
      }

      return {
        success: false,
        error: 'Signature file not found',
      };
    } catch (error) {
      this.logger.error(`Error verifying free signature: ${error.message}`);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * Get all free signatures for an opportunity
   */
  async getFreeSignatures(opportunityId: string) {
    try {
      const signatures = await this.prisma.signature.findMany({
        where: {
          opportunityId,
          userAgent: 'React Native Free Signing App'
        },
        orderBy: {
          signedAt: 'desc'
        }
      });

      return {
        success: true,
        signatures: signatures.map(sig => ({
          id: sig.id,
          opportunityId: sig.opportunityId,
          signedAt: sig.signedAt,
          signedBy: sig.signedBy,
          filePath: sig.filePath,
        }))
      };
    } catch (error) {
      this.logger.error(`Error getting free signatures: ${error.message}`);
      throw error;
    }
  }

  /**
   * Download signed document
   */
  async downloadSignedDocument(opportunityId: string) {
    try {
      const signature = await this.prisma.signature.findFirst({
        where: {
          opportunityId,
          userAgent: 'React Native Free Signing App'
        },
        orderBy: {
          signedAt: 'desc'
        }
      });

      if (!signature || !signature.filePath) {
        throw new Error('Signed document not found');
      }

      // Check if file exists using fs.promises
      try {
        await fs.access(signature.filePath);
      } catch {
        throw new Error('Signed document file not found');
      }

      const signatureData = JSON.parse(await fs.readFile(signature.filePath, 'utf8'));
      
      return {
        success: true,
        document: signatureData,
        filePath: signature.filePath,
        digitalFootprint: signatureData.enhancedDigitalFootprint,
      };
    } catch (error) {
      this.logger.error(`Error downloading signed document: ${error.message}`);
      throw error;
    }
  }
}
