/*
  Warnings:

  - You are about to drop the column `role` on the `User` table. All the data in the column will be lost.
  - A unique constraint covering the columns `[ghlUserId]` on the table `User` will be added. If there are existing duplicate values, this will fail.
  - Added the required column `ghlAccessToken` to the `User` table without a default value. This is not possible if the table is not empty.
  - Added the required column `ghlRefreshToken` to the `User` table without a default value. This is not possible if the table is not empty.
  - Added the required column `ghlUserId` to the `User` table without a default value. This is not possible if the table is not empty.
  - Added the required column `updatedAt` to the `User` table without a default value. This is not possible if the table is not empty.

*/
-- DropIndex
DROP INDEX "User_email_key";

-- AlterTable
ALTER TABLE "User" DROP COLUMN "role",
ADD COLUMN     "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
ADD COLUMN     "ghlAccessToken" TEXT NOT NULL,
ADD COLUMN     "ghlRefreshToken" TEXT NOT NULL,
ADD COLUMN     "ghlUserId" TEXT NOT NULL,
ADD COLUMN     "tokenExpiresAt" TIMESTAMP(3),
ADD COLUMN     "updatedAt" TIMESTAMP(3) NOT NULL,
ALTER COLUMN "email" DROP NOT NULL,
ALTER COLUMN "name" DROP NOT NULL;

-- CreateIndex
CREATE UNIQUE INDEX "User_ghlUserId_key" ON "User"("ghlUserId");
