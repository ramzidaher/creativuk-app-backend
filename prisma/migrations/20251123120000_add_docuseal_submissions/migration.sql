-- CreateEnum
CREATE TYPE "DocuSealDocumentType" AS ENUM ('CONTRACT', 'DISCLAIMER', 'BOOKING_CONFIRMATION');

-- CreateTable
CREATE TABLE "DocuSealSubmission" (
    "id" TEXT NOT NULL,
    "opportunityId" TEXT NOT NULL,
    "documentType" "DocuSealDocumentType" NOT NULL,
    "templateId" TEXT NOT NULL,
    "submissionId" TEXT NOT NULL,
    "signingUrl" TEXT,
    "status" TEXT NOT NULL DEFAULT 'pending',
    "signedDocumentUrl" TEXT,
    "customerName" TEXT,
    "customerEmail" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    "completedAt" TIMESTAMP(3),

    CONSTRAINT "DocuSealSubmission_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "DocuSealSubmission_opportunityId_documentType_key" ON "DocuSealSubmission"("opportunityId", "documentType");

-- CreateIndex
CREATE INDEX "DocuSealSubmission_opportunityId_idx" ON "DocuSealSubmission"("opportunityId");

-- CreateIndex
CREATE INDEX "DocuSealSubmission_submissionId_idx" ON "DocuSealSubmission"("submissionId");

-- CreateIndex
CREATE INDEX "DocuSealSubmission_documentType_idx" ON "DocuSealSubmission"("documentType");

-- CreateIndex
CREATE INDEX "DocuSealSubmission_status_idx" ON "DocuSealSubmission"("status");



