-- AlterEnum
-- This migration adds missing StepType enum values that exist in the Prisma schema
-- but are missing from the database. Adding enum values is safe and does not cause data loss.
-- This uses DO blocks to safely add values only if they don't already exist.

DO $$ 
BEGIN
    -- Add EXPRESS_CONSENT if it doesn't exist
    IF NOT EXISTS (
        SELECT 1 FROM pg_enum 
        WHERE enumlabel = 'EXPRESS_CONSENT' 
        AND enumtypid = (SELECT oid FROM pg_type WHERE typname = 'StepType')
    ) THEN
        ALTER TYPE "public"."StepType" ADD VALUE 'EXPRESS_CONSENT';
    END IF;
END $$;

DO $$ 
BEGIN
    -- Add WELCOME_EMAIL if it doesn't exist
    IF NOT EXISTS (
        SELECT 1 FROM pg_enum 
        WHERE enumlabel = 'WELCOME_EMAIL' 
        AND enumtypid = (SELECT oid FROM pg_type WHERE typname = 'StepType')
    ) THEN
        ALTER TYPE "public"."StepType" ADD VALUE 'WELCOME_EMAIL';
    END IF;
END $$;

DO $$ 
BEGIN
    -- Add INSTALLATION_BOOKING if it doesn't exist
    IF NOT EXISTS (
        SELECT 1 FROM pg_enum 
        WHERE enumlabel = 'INSTALLATION_BOOKING' 
        AND enumtypid = (SELECT oid FROM pg_type WHERE typname = 'StepType')
    ) THEN
        ALTER TYPE "public"."StepType" ADD VALUE 'INSTALLATION_BOOKING';
    END IF;
END $$;

DO $$ 
BEGIN
    -- Add SOLAR_PROJECTION if it doesn't exist
    IF NOT EXISTS (
        SELECT 1 FROM pg_enum 
        WHERE enumlabel = 'SOLAR_PROJECTION' 
        AND enumtypid = (SELECT oid FROM pg_type WHERE typname = 'StepType')
    ) THEN
        ALTER TYPE "public"."StepType" ADD VALUE 'SOLAR_PROJECTION';
    END IF;
END $$;

DO $$ 
BEGIN
    -- Add EMAIL_CONFIRMATION if it doesn't exist
    IF NOT EXISTS (
        SELECT 1 FROM pg_enum 
        WHERE enumlabel = 'EMAIL_CONFIRMATION' 
        AND enumtypid = (SELECT oid FROM pg_type WHERE typname = 'StepType')
    ) THEN
        ALTER TYPE "public"."StepType" ADD VALUE 'EMAIL_CONFIRMATION';
    END IF;
END $$;

