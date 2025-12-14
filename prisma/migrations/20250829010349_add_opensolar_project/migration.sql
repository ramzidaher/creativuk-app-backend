-- CreateTable
CREATE TABLE "public"."OpenSolarProject" (
    "id" TEXT NOT NULL,
    "opportunityId" TEXT NOT NULL,
    "opensolarProjectId" INTEGER NOT NULL,
    "projectName" TEXT NOT NULL,
    "address" TEXT,
    "systems" JSONB NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "OpenSolarProject_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "OpenSolarProject_opportunityId_key" ON "public"."OpenSolarProject"("opportunityId");
