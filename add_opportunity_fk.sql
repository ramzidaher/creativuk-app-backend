-- Add foreign key constraint with NOT VALID to avoid validating existing data
-- This allows existing OpportunityOutcome records without matching Opportunity records
-- The constraint will only be enforced for new inserts/updates, not existing data
ALTER TABLE "OpportunityOutcome" 
ADD CONSTRAINT "OpportunityOutcome_opportunityGhlOpportunityId_fkey" 
FOREIGN KEY ("ghlOpportunityId") 
REFERENCES "Opportunity"("ghlOpportunityId") 
ON DELETE CASCADE 
ON UPDATE CASCADE 
NOT VALID;

-- Note: We're NOT validating the constraint yet to preserve existing data
-- Once opportunities are synced, you can validate it with:
-- ALTER TABLE "OpportunityOutcome" VALIDATE CONSTRAINT "OpportunityOutcome_opportunityGhlOpportunityId_fkey";

