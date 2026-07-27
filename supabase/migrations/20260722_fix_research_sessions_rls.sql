-- ============================================================
-- Fix RLS policies for research_sessions table
-- RLS was enabled with zero policies, silently blocking all
-- SELECT access from the browser client (anon/authenticated) —
-- listPipelineSessions() always returned an empty result with
-- no error, so "Recent Pipelines" appeared empty despite rows
-- existing in the table.
-- ============================================================

-- 1. Enable Row Level Security (already enabled, but safe to re-assert)
ALTER TABLE research_sessions ENABLE ROW LEVEL SECURITY;

-- 2. Drop existing policies to prevent conflicts
DROP POLICY IF EXISTS "research_sessions_select" ON research_sessions;
DROP POLICY IF EXISTS "research_sessions_insert" ON research_sessions;
DROP POLICY IF EXISTS "research_sessions_update" ON research_sessions;
DROP POLICY IF EXISTS "research_sessions_delete" ON research_sessions;

-- 3. Create clean policies allowing authenticated users full access
CREATE POLICY "research_sessions_select" ON research_sessions
  FOR SELECT TO authenticated USING (true);

CREATE POLICY "research_sessions_insert" ON research_sessions
  FOR INSERT TO authenticated WITH CHECK (true);

CREATE POLICY "research_sessions_update" ON research_sessions
  FOR UPDATE TO authenticated USING (true);

CREATE POLICY "research_sessions_delete" ON research_sessions
  FOR DELETE TO authenticated USING (true);
