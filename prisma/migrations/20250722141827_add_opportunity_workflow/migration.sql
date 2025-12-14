-- CreateEnum
CREATE TYPE "OpportunityStatus" AS ENUM ('IN_PROGRESS', 'COMPLETED', 'PAUSED', 'CANCELLED');

-- CreateEnum
CREATE TYPE "StepType" AS ENUM ('INITIAL_CONTACT', 'SURVEY_SCHEDULING', 'SITE_SURVEY', 'PROPOSAL_GENERATION', 'CONTRACT_SIGNING', 'INSTALLATION_SCHEDULING', 'FOLLOW_UP');

-- CreateEnum
CREATE TYPE "StepStatus" AS ENUM ('PENDING', 'IN_PROGRESS', 'COMPLETED', 'SKIPPED');

-- CreateTable
CREATE TABLE "OpportunityProgress" (
    "id" TEXT NOT NULL,
    "ghlOpportunityId" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "currentStep" INTEGER NOT NULL DEFAULT 1,
    "totalSteps" INTEGER NOT NULL DEFAULT 5,
    "status" "OpportunityStatus" NOT NULL DEFAULT 'IN_PROGRESS',
    "startedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "lastActivityAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "completedAt" TIMESTAMP(3),
    "stepData" JSONB,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "OpportunityProgress_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "OpportunityStep" (
    "id" TEXT NOT NULL,
    "opportunityProgressId" TEXT NOT NULL,
    "stepNumber" INTEGER NOT NULL,
    "stepType" "StepType" NOT NULL,
    "status" "StepStatus" NOT NULL DEFAULT 'PENDING',
    "data" JSONB,
    "startedAt" TIMESTAMP(3),
    "completedAt" TIMESTAMP(3),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "OpportunityStep_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "OpportunityProgress_ghlOpportunityId_key" ON "OpportunityProgress"("ghlOpportunityId");

-- CreateIndex
CREATE UNIQUE INDEX "OpportunityStep_opportunityProgressId_stepNumber_key" ON "OpportunityStep"("opportunityProgressId", "stepNumber");

-- AddForeignKey
ALTER TABLE "OpportunityProgress" ADD CONSTRAINT "OpportunityProgress_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "OpportunityStep" ADD CONSTRAINT "OpportunityStep_opportunityProgressId_fkey" FOREIGN KEY ("opportunityProgressId") REFERENCES "OpportunityProgress"("id") ON DELETE CASCADE ON UPDATE CASCADE;
