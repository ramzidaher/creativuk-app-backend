-- CreateEnum - Create UserRole enum
CREATE TYPE "UserRole" AS ENUM ('ADMIN', 'SURVEYOR');

-- CreateEnum - Create UserStatus enum
CREATE TYPE "UserStatus" AS ENUM ('ACTIVE', 'INACTIVE', 'SUSPENDED');

-- CreateEnum - Create SurveyStatus enum
CREATE TYPE "SurveyStatus" AS ENUM ('DRAFT', 'IN_PROGRESS', 'COMPLETED', 'SUBMITTED', 'APPROVED', 'REJECTED');

-- AlterTable - Add missing authentication fields to User
ALTER TABLE "User" ADD COLUMN "username" TEXT NOT NULL,
ADD COLUMN "password" TEXT NOT NULL,
ADD COLUMN "role" "UserRole" NOT NULL DEFAULT 'SURVEYOR',
ADD COLUMN "status" "UserStatus" NOT NULL DEFAULT 'ACTIVE',
ADD COLUMN "isEmailVerified" BOOLEAN NOT NULL DEFAULT false,
ADD COLUMN "emailVerificationToken" TEXT,
ADD COLUMN "passwordResetToken" TEXT,
ADD COLUMN "passwordResetExpires" TIMESTAMP(3),
ADD COLUMN "lastLoginAt" TIMESTAMP(3);

-- AlterTable - Make email unique and not null
ALTER TABLE "User" ALTER COLUMN "email" SET NOT NULL;

-- CreateTable - Create Survey table
CREATE TABLE "Survey" (
    "id" TEXT NOT NULL,
    "ghlOpportunityId" TEXT NOT NULL,
    "ghlUserId" TEXT,
    "page1" JSONB,
    "page2" JSONB,
    "page3" JSONB,
    "page4" JSONB,
    "page5" JSONB,
    "page6" JSONB,
    "page7" JSONB,
    "page8" JSONB,
    "status" "SurveyStatus" NOT NULL DEFAULT 'DRAFT',
    "eligibilityScore" INTEGER,
    "rejectionReason" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    "submittedAt" TIMESTAMP(3),
    "approvedAt" TIMESTAMP(3),
    "rejectedAt" TIMESTAMP(3),
    "deletedAt" TIMESTAMP(3),
    "isDeleted" BOOLEAN NOT NULL DEFAULT false,
    "createdBy" TEXT NOT NULL,
    "updatedBy" TEXT,

    CONSTRAINT "Survey_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "User_username_key" ON "User"("username");
CREATE UNIQUE INDEX "User_email_key" ON "User"("email");
CREATE UNIQUE INDEX "Survey_ghlOpportunityId_key" ON "Survey"("ghlOpportunityId");

-- AddForeignKey
ALTER TABLE "Survey" ADD CONSTRAINT "Survey_ghlUserId_fkey" FOREIGN KEY ("ghlUserId") REFERENCES "User"("ghlUserId") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Survey" ADD CONSTRAINT "Survey_createdBy_fkey" FOREIGN KEY ("createdBy") REFERENCES "User"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Survey" ADD CONSTRAINT "Survey_updatedBy_fkey" FOREIGN KEY ("updatedBy") REFERENCES "User"("id") ON DELETE SET NULL ON UPDATE CASCADE; 