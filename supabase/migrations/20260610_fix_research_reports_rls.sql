-- ============================================================
-- Fix RLS policies for research_reports table
-- ============================================================

-- 1. Enable Row Level Security
ALTER TABLE research_reports ENABLE ROW LEVEL SECURITY;

-- 2. Drop existing policies to prevent conflicts
DROP POLICY IF EXISTS "research_reports_select" ON research_reports;
DROP POLICY IF EXISTS "research_reports_insert" ON research_reports;
DROP POLICY IF EXISTS "research_reports_update" ON research_reports;
DROP POLICY IF EXISTS "research_reports_delete" ON research_reports;
DROP POLICY IF EXISTS "Allow select for authenticated" ON research_reports;
DROP POLICY IF EXISTS "Allow insert for authenticated" ON research_reports;
DROP POLICY IF EXISTS "Allow update for authenticated" ON research_reports;

-- 3. Create clean policies allowing authenticated users full access
CREATE POLICY "research_reports_select" ON research_reports 
  FOR SELECT TO authenticated USING (true);

CREATE POLICY "research_reports_insert" ON research_reports 
  FOR INSERT TO authenticated WITH CHECK (true);

CREATE POLICY "research_reports_update" ON research_reports 
  FOR UPDATE TO authenticated USING (true);

CREATE POLICY "research_reports_delete" ON research_reports 
  FOR DELETE TO authenticated USING (true);

-- 4. Enable public read access if needed for public sharing/viewing
-- (Optional: uncomment if you need unauthenticated public users to see reports)
-- CREATE POLICY "research_reports_public_read" ON research_reports
--   FOR SELECT TO anon USING (true);
