-- CreateTable
CREATE TABLE "public"."Signature" (
    "id" TEXT NOT NULL,
    "opportunityId" TEXT NOT NULL,
    "signatureData" TEXT NOT NULL,
    "signedAt" TIMESTAMP(3) NOT NULL,
    "signedBy" TEXT,
    "ipAddress" TEXT,
    "userAgent" TEXT,
    "filePath" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Signature_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "Signature_opportunityId_idx" ON "public"."Signature"("opportunityId");
