-- AlterEnum
-- This migration adds more than one value to an enum.
-- With PostgreSQL versions 11 and earlier, this is not possible
-- in a single migration. This can be worked around by creating
-- multiple migrations, each migration adding only one value to
-- the enum.


ALTER TYPE "public"."StepType" ADD VALUE 'PAYMENT';
ALTER TYPE "public"."StepType" ADD VALUE 'PRICING';
ALTER TYPE "public"."StepType" ADD VALUE 'FINISH_APPOINTMENT';

-- CreateTable
CREATE TABLE "public"."SurveyImage" (
    "id" TEXT NOT NULL,
    "surveyId" TEXT NOT NULL,
    "fieldName" TEXT NOT NULL,
    "fileName" TEXT NOT NULL,
    "originalName" TEXT NOT NULL,
    "mimeType" TEXT NOT NULL,
    "fileSize" INTEGER NOT NULL,
    "filePath" TEXT NOT NULL,
    "base64Data" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "SurveyImage_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "public"."AutoSave" (
    "id" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "opportunityId" TEXT NOT NULL,
    "data" JSONB NOT NULL DEFAULT '{}',
    "lastSavedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    "deletedAt" TIMESTAMP(3),
    "isDeleted" BOOLEAN NOT NULL DEFAULT false,

    CONSTRAINT "AutoSave_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "SurveyImage_surveyId_idx" ON "public"."SurveyImage"("surveyId");

-- CreateIndex
CREATE INDEX "SurveyImage_fieldName_idx" ON "public"."SurveyImage"("fieldName");

-- CreateIndex
CREATE INDEX "AutoSave_userId_idx" ON "public"."AutoSave"("userId");

-- CreateIndex
CREATE INDEX "AutoSave_opportunityId_idx" ON "public"."AutoSave"("opportunityId");

-- CreateIndex
CREATE INDEX "AutoSave_lastSavedAt_idx" ON "public"."AutoSave"("lastSavedAt");

-- CreateIndex
CREATE UNIQUE INDEX "AutoSave_userId_opportunityId_isDeleted_key" ON "public"."AutoSave"("userId", "opportunityId", "isDeleted");

-- AddForeignKey
ALTER TABLE "public"."SurveyImage" ADD CONSTRAINT "SurveyImage_surveyId_fkey" FOREIGN KEY ("surveyId") REFERENCES "public"."Survey"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "public"."AutoSave" ADD CONSTRAINT "AutoSave_userId_fkey" FOREIGN KEY ("userId") REFERENCES "public"."User"("id") ON DELETE CASCADE ON UPDATE CASCADE;
