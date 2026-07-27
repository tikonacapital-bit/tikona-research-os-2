-- ============================================================
-- Fix RLS policies for the remaining tables the frontend actually
-- queries: session_documents, sector_playbooks, recommendations,
-- audit_logs. All four had RLS enabled with zero policies after
-- a blanket "enable RLS" pass, silently blocking the app.
--
-- The other 14 tables flagged by the same audit (ai_wallets,
-- analysis_runs, document_chunks, documents, evidence_registry,
-- financial_models, llm_calls, report_charts, sebi_compliance_vault,
-- stock_analysis, stocks, user_roles, watchlist, webhook_logs) are
-- not referenced anywhere in this codebase — leave them locked down.
-- ============================================================

-- session_documents — full CRUD from the client (vault document selection)
ALTER TABLE session_documents ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "session_documents_select" ON session_documents;
DROP POLICY IF EXISTS "session_documents_insert" ON session_documents;
DROP POLICY IF EXISTS "session_documents_update" ON session_documents;
DROP POLICY IF EXISTS "session_documents_delete" ON session_documents;
CREATE POLICY "session_documents_select" ON session_documents FOR SELECT TO authenticated USING (true);
CREATE POLICY "session_documents_insert" ON session_documents FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "session_documents_update" ON session_documents FOR UPDATE TO authenticated USING (true);
CREATE POLICY "session_documents_delete" ON session_documents FOR DELETE TO authenticated USING (true);

-- sector_playbooks — full CRUD from the client (Stage 0 cache + Sector Thesis editor)
ALTER TABLE sector_playbooks ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "sector_playbooks_select" ON sector_playbooks;
DROP POLICY IF EXISTS "sector_playbooks_insert" ON sector_playbooks;
DROP POLICY IF EXISTS "sector_playbooks_update" ON sector_playbooks;
DROP POLICY IF EXISTS "sector_playbooks_delete" ON sector_playbooks;
CREATE POLICY "sector_playbooks_select" ON sector_playbooks FOR SELECT TO authenticated USING (true);
CREATE POLICY "sector_playbooks_insert" ON sector_playbooks FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "sector_playbooks_update" ON sector_playbooks FOR UPDATE TO authenticated USING (true);
CREATE POLICY "sector_playbooks_delete" ON sector_playbooks FOR DELETE TO authenticated USING (true);

-- recommendations — full CRUD from the client (Telegram recommendation log)
ALTER TABLE recommendations ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "recommendations_select" ON recommendations;
DROP POLICY IF EXISTS "recommendations_insert" ON recommendations;
DROP POLICY IF EXISTS "recommendations_update" ON recommendations;
DROP POLICY IF EXISTS "recommendations_delete" ON recommendations;
CREATE POLICY "recommendations_select" ON recommendations FOR SELECT TO authenticated USING (true);
CREATE POLICY "recommendations_insert" ON recommendations FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "recommendations_update" ON recommendations FOR UPDATE TO authenticated USING (true);
CREATE POLICY "recommendations_delete" ON recommendations FOR DELETE TO authenticated USING (true);

-- audit_logs — INSERT only. The app never reads, updates, or deletes
-- audit log rows from the client; keep it that way so the log stays
-- tamper-evident even from an authenticated session.
ALTER TABLE audit_logs ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "audit_logs_insert" ON audit_logs;
CREATE POLICY "audit_logs_insert" ON audit_logs FOR INSERT TO authenticated WITH CHECK (true);
