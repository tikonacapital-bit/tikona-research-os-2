-- =====================================================
-- Pipeline Migration: 3-Stage Research Pipeline
-- =====================================================
-- Run this migration against your Supabase database.
-- It adds pipeline_status and stage output columns to
-- research_sessions, creates sector knowledge tables,
-- and the research_sections table.
-- =====================================================

-- 1. Add pipeline columns to research_sessions
ALTER TABLE research_sessions
  ADD COLUMN IF NOT EXISTS pipeline_status text DEFAULT 'documents_ready',
  ADD COLUMN IF NOT EXISTS sector_framework jsonb DEFAULT NULL,
  ADD COLUMN IF NOT EXISTS thesis_condensed text DEFAULT NULL,
  ADD COLUMN IF NOT EXISTS thesis_output text DEFAULT NULL,
  ADD COLUMN IF NOT EXISTS report_content text DEFAULT NULL,
  ADD COLUMN IF NOT EXISTS selected_model text DEFAULT NULL,
  ADD COLUMN IF NOT EXISTS total_tokens_used integer DEFAULT 0,
  ADD COLUMN IF NOT EXISTS generation_time_seconds integer DEFAULT 0;

-- 2. Sectors table
CREATE TABLE IF NOT EXISTS sectors (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sector_name text UNIQUE NOT NULL,
  description text,
  parent_sector text,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

-- 3. Sector Knowledge Base (SKB)
CREATE TABLE IF NOT EXISTS sector_knowledge (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sector_name text NOT NULL REFERENCES sectors(sector_name) ON DELETE CASCADE,
  category text NOT NULL, -- e.g., 'overview', 'key_metrics', 'value_chain', 'competitive_dynamics', 'regulatory', 'growth_drivers', 'risks', 'valuation'
  title text NOT NULL,
  content text NOT NULL,
  source text,
  sort_order integer DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_sector_knowledge_sector ON sector_knowledge(sector_name);
CREATE INDEX IF NOT EXISTS idx_sector_knowledge_category ON sector_knowledge(category);

-- 4. Sector Knowledge Embeddings (for vector search)
CREATE TABLE IF NOT EXISTS sector_knowledge_embeddings (
  id bigserial PRIMARY KEY,
  knowledge_id uuid REFERENCES sector_knowledge(id) ON DELETE CASCADE,
  sector_name text NOT NULL,
  content text NOT NULL,
  embedding vector(1536), -- OpenAI embedding dimension
  created_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_ske_sector ON sector_knowledge_embeddings(sector_name);

-- 5. Sector-Company Mapping
CREATE TABLE IF NOT EXISTS sector_company_mapping (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sector_name text NOT NULL REFERENCES sectors(sector_name) ON DELETE CASCADE,
  company_id integer NOT NULL,
  nse_symbol text NOT NULL,
  company_name text NOT NULL,
  created_at timestamptz DEFAULT now(),
  UNIQUE(sector_name, company_id)
);

CREATE INDEX IF NOT EXISTS idx_scm_sector ON sector_company_mapping(sector_name);
CREATE INDEX IF NOT EXISTS idx_scm_company ON sector_company_mapping(company_id);

-- 6. Research Sections (output from each pipeline stage)
CREATE TABLE IF NOT EXISTS research_sections (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  session_id uuid NOT NULL REFERENCES research_sessions(session_id) ON DELETE CASCADE,
  section_key text NOT NULL,
  section_title text NOT NULL,
  stage text NOT NULL CHECK (stage IN ('stage0', 'stage1', 'stage2')),
  content text NOT NULL DEFAULT '',
  heading text,
  sort_order integer DEFAULT 0,
  tokens_used integer DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_rs_session ON research_sections(session_id);
CREATE INDEX IF NOT EXISTS idx_rs_stage ON research_sections(stage);

-- 7. SKB Suggested Updates (pipeline suggests updates to sector knowledge)
CREATE TABLE IF NOT EXISTS skb_suggested_updates (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  session_id uuid REFERENCES research_sessions(session_id) ON DELETE SET NULL,
  sector_name text NOT NULL,
  category text NOT NULL,
  title text NOT NULL,
  suggested_content text NOT NULL,
  status text DEFAULT 'pending' CHECK (status IN ('pending', 'approved', 'rejected')),
  reviewed_by text,
  reviewed_at timestamptz,
  created_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_ssu_sector ON skb_suggested_updates(sector_name);
CREATE INDEX IF NOT EXISTS idx_ssu_status ON skb_suggested_updates(status);

-- 8. Enable RLS on new tables
ALTER TABLE sectors ENABLE ROW LEVEL SECURITY;
ALTER TABLE sector_knowledge ENABLE ROW LEVEL SECURITY;
ALTER TABLE sector_knowledge_embeddings ENABLE ROW LEVEL SECURITY;
ALTER TABLE sector_company_mapping ENABLE ROW LEVEL SECURITY;
ALTER TABLE research_sections ENABLE ROW LEVEL SECURITY;
ALTER TABLE skb_suggested_updates ENABLE ROW LEVEL SECURITY;

-- 9. RLS Policies (allow authenticated users full access for now)
-- Sectors
CREATE POLICY IF NOT EXISTS "sectors_select" ON sectors FOR SELECT TO authenticated USING (true);
CREATE POLICY IF NOT EXISTS "sectors_insert" ON sectors FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY IF NOT EXISTS "sectors_update" ON sectors FOR UPDATE TO authenticated USING (true);

-- Sector Knowledge
CREATE POLICY IF NOT EXISTS "sk_select" ON sector_knowledge FOR SELECT TO authenticated USING (true);
CREATE POLICY IF NOT EXISTS "sk_insert" ON sector_knowledge FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY IF NOT EXISTS "sk_update" ON sector_knowledge FOR UPDATE TO authenticated USING (true);

-- Sector Knowledge Embeddings
CREATE POLICY IF NOT EXISTS "ske_select" ON sector_knowledge_embeddings FOR SELECT TO authenticated USING (true);

-- Sector Company Mapping
CREATE POLICY IF NOT EXISTS "scm_select" ON sector_company_mapping FOR SELECT TO authenticated USING (true);
CREATE POLICY IF NOT EXISTS "scm_insert" ON sector_company_mapping FOR INSERT TO authenticated WITH CHECK (true);

-- Research Sections
CREATE POLICY IF NOT EXISTS "rs_select" ON research_sections FOR SELECT TO authenticated USING (true);
CREATE POLICY IF NOT EXISTS "rs_insert" ON research_sections FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY IF NOT EXISTS "rs_update" ON research_sections FOR UPDATE TO authenticated USING (true);
CREATE POLICY IF NOT EXISTS "rs_delete" ON research_sections FOR DELETE TO authenticated USING (true);

-- SKB Suggested Updates
CREATE POLICY IF NOT EXISTS "ssu_select" ON skb_suggested_updates FOR SELECT TO authenticated USING (true);
CREATE POLICY IF NOT EXISTS "ssu_insert" ON skb_suggested_updates FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY IF NOT EXISTS "ssu_update" ON skb_suggested_updates FOR UPDATE TO authenticated USING (true);

-- 10. Seed some common sectors
INSERT INTO sectors (sector_name, description) VALUES
  ('Automobile', 'Automobile manufacturers, auto ancillaries, EV companies'),
  ('Banking', 'Private banks, PSU banks, small finance banks, payment banks'),
  ('Capital Goods', 'Industrial machinery, electrical equipment, defense'),
  ('Cement', 'Cement manufacturers and building materials'),
  ('Chemicals', 'Specialty chemicals, agrochemicals, petrochemicals'),
  ('Consumer Durables', 'Electronics, appliances, lifestyle products'),
  ('FMCG', 'Fast moving consumer goods, food & beverages'),
  ('Healthcare', 'Pharmaceuticals, hospitals, diagnostics, medical devices'),
  ('Information Technology', 'IT services, product companies, SaaS'),
  ('Infrastructure', 'Construction, EPC, real estate development'),
  ('Insurance', 'Life insurance, general insurance, health insurance'),
  ('Media', 'Broadcasting, digital media, entertainment, publishing'),
  ('Metals & Mining', 'Steel, aluminum, copper, mining companies'),
  ('NBFC', 'Non-banking financial companies, housing finance, microfinance'),
  ('Oil & Gas', 'Exploration, refining, distribution, gas utilities'),
  ('Power', 'Power generation, transmission, distribution, renewables'),
  ('Retail', 'Organized retail, e-commerce, fashion, grocery'),
  ('Telecom', 'Telecom service providers, tower companies, fiber'),
  ('Textiles', 'Apparel, home textiles, yarn, fabric manufacturers')
ON CONFLICT (sector_name) DO NOTHING;
