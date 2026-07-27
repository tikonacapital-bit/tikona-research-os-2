-- ============================================================
-- Fix RLS policies for master_company table
-- RLS was enabled with zero policies, silently blocking all
-- SELECT access from the browser client (anon/authenticated) —
-- useCompanySearch always returned an empty result with no
-- error, so the company search box showed no suggestions.
-- ============================================================

-- 1. Enable Row Level Security (already enabled, but safe to re-assert)
ALTER TABLE master_company ENABLE ROW LEVEL SECURITY;

-- 2. Drop existing policies to prevent conflicts
DROP POLICY IF EXISTS "master_company_select" ON master_company;
DROP POLICY IF EXISTS "master_company_insert" ON master_company;
DROP POLICY IF EXISTS "master_company_update" ON master_company;
DROP POLICY IF EXISTS "master_company_delete" ON master_company;

-- 3. Create clean policies allowing authenticated users full access
CREATE POLICY "master_company_select" ON master_company
  FOR SELECT TO authenticated USING (true);

CREATE POLICY "master_company_insert" ON master_company
  FOR INSERT TO authenticated WITH CHECK (true);

CREATE POLICY "master_company_update" ON master_company
  FOR UPDATE TO authenticated USING (true);

CREATE POLICY "master_company_delete" ON master_company
  FOR DELETE TO authenticated USING (true);
