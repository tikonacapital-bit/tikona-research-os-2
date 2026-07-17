// Database types for Tikona Research OS
// Maps directly to Supabase table schemas

// ========================
// Report Section Keys (13-section structure)
// ========================

// Text sections (AI-generated content)
export type TextSectionKey =
  | 'company_background'
  | 'business_model'
  | 'management_analysis'
  | 'industry_overview'
  | 'industry_tailwinds'
  | 'demand_drivers'
  | 'industry_risks';

// Chart sections (images from Supabase Storage)
export type ChartSectionKey =
  | 'revenue_mix'
  | 'profit_loss'
  | 'balance_sheet'
  | 'cash_flow'
  | 'ratio_analysis'
  | 'summary';

// Combined type for all report sections
export type ReportSectionKey = TextSectionKey | ChartSectionKey;

// Slide order configuration for PPT generation
export interface SlideConfig {
  key: ReportSectionKey;
  title: string;
  type: 'text' | 'chart';
  order: number;
}

// Complete slide order for PPT
export const SLIDE_ORDER: SlideConfig[] = [
  { key: 'company_background', title: 'Company Background', type: 'text', order: 1 },
  { key: 'business_model', title: 'Business Model', type: 'text', order: 2 },
  { key: 'revenue_mix', title: 'Revenue Mix', type: 'chart', order: 3 },
  { key: 'management_analysis', title: 'Management Analysis', type: 'text', order: 4 },
  { key: 'industry_overview', title: 'Industry Overview', type: 'text', order: 5 },
  { key: 'industry_tailwinds', title: 'Key Industry Tailwinds', type: 'text', order: 6 },
  { key: 'demand_drivers', title: 'Demand Drivers', type: 'text', order: 7 },
  { key: 'industry_risks', title: 'Industry Risks', type: 'text', order: 8 },
  { key: 'profit_loss', title: 'Profit & Loss Statement', type: 'chart', order: 9 },
  { key: 'balance_sheet', title: 'Balance Sheet', type: 'chart', order: 10 },
  { key: 'cash_flow', title: 'Cash Flow Statement', type: 'chart', order: 11 },
  { key: 'ratio_analysis', title: 'Ratio Analysis', type: 'chart', order: 12 },
  { key: 'summary', title: 'Summary', type: 'chart', order: 13 },
];

// Matches the equity_universe table (financial data)
export interface EquityUniverse {
  company_id: number;
  company_name: string | null;
  isin_code: string;
  nse_code: string | null;
  bse_code: string | null;
  google_code: string | null;
  broad_sector: string | null;
  sector: string | null;
  broad_industry: string | null;
  industry: string | null;
  nse_bom_code: string | null;

  // Market Data
  market_cap: number | null;
  current_price: number | null;
  high_52_week: number | null;
  low_52_week: number | null;
  volume: number | null;

  // Latest Quarter Results
  quarterly_results_date: string | null;
  sales_latest_qtr: number | null;
  op_profit_latest_qtr: number | null;
  pat_latest_qtr: number | null;
  ebitda_margin_latest_qtr: number | null;
  pat_margin_latest_qtr: number | null;

  // Preceding Quarter
  sales_preceding_qtr: number | null;
  op_profit_preceding_qtr: number | null;
  pat_preceding_qtr: number | null;

  // QoQ Growth
  revenue_growth_qoq: number | null;
  ebitda_growth_qoq: number | null;
  ebitda_margin_growth_qoq_bps: number | null;
  pat_growth_qoq: number | null;
  pat_margin_growth_qoq_bps: number | null;

  // YoY Growth
  sales_growth_yoy_qtr: number | null;
  profit_growth_yoy_qtr: number | null;

  // Annual Data
  last_annual_result_date: string | null;
  sales_ttm_screener: number | null;
  op_profit_ttm: number | null;
  pat_ttm_screener: number | null;
  ebitda_margin_ttm: number | null;
  opm_last_year: number | null;
  pat_margin_ttm: number | null;

  // Per Share Data
  num_equity_shares: number | null;
  eps_ttm_actual: number | null;

  // Balance Sheet
  debt: number | null;
  cash_equivalents: number | null;
  net_debt: number | null;
  net_worth: number | null;
  book_value: number | null;
  enterprise_value: number | null;

  // Shareholding
  promoter_holding_pct: number | null;
  unpledged_promoter_holding_pct: number | null;

  // Ratios
  working_capital_to_sales_ratio: number | null;
  asset_turnover_ratio: number | null;
  roe: number | null;
  roce: number | null;
  roic: number | null;

  // Fixed Assets
  net_block: number | null;
  cwip: number | null;
  cwip_to_net_block_ratio: number | null;

  // Historical Revenue
  revenue_fy2023: number | null;
  revenue_fy2024: number | null;
  revenue_fy2025: number | null;
  revenue_ttm: number | null;
  revenue_fy2026e: number | null;
  revenue_fy2027e: number | null;
  revenue_fy2028e: number | null;
  revenue_cagr_hist_2yr: number | null;
  revenue_cagr_fwd_2yr: number | null;

  // Historical EBITDA
  ebitda_fy2023: number | null;
  ebitda_fy2024: number | null;
  ebitda_fy2025: number | null;
  ebitda_ttm: number | null;
  ebitda_fy2026e: number | null;
  ebitda_fy2027e: number | null;
  ebitda_fy2028e: number | null;
  ebitda_cagr_hist_2yr: number | null;
  ebitda_cagr_fwd_2yr: number | null;

  // Historical PAT
  pat_fy2023: number | null;
  pat_fy2024: number | null;
  pat_fy2025: number | null;
  pat_ttm: number | null;
  pat_fy2026e: number | null;
  pat_fy2027e: number | null;
  pat_fy2028e: number | null;
  pat_cagr_hist_2yr: number | null;
  pat_cagr_fwd_2yr: number | null;

  // Margins
  ebitda_margin_fy2023: number | null;
  ebitda_margin_fy2024: number | null;
  ebitda_margin_fy2025: number | null;
  ebitda_margin_ttm_calc: number | null;
  ebitda_margin_fy2026e: number | null;
  ebitda_margin_fy2027e: number | null;
  ebitda_margin_fy2028e: number | null;
  pat_margin_fy2023: number | null;
  pat_margin_fy2024: number | null;
  pat_margin_fy2025: number | null;
  pat_margin_ttm_calc: number | null;
  pat_margin_fy2026e: number | null;
  pat_margin_fy2027e: number | null;
  pat_margin_fy2028e: number | null;

  // EPS
  eps_fy2023: number | null;
  eps_fy2024: number | null;
  eps_fy2025: number | null;
  eps_ttm: number | null;
  eps_fy2026e: number | null;
  eps_fy2027e: number | null;
  eps_fy2028e: number | null;
  eps_cagr_hist_2yr: number | null;
  eps_cagr_fwd_2yr: number | null;

  // Valuation Multiples
  pe_ttm: number | null;
  pe_fy2026e: number | null;
  pe_fy2027e: number | null;
  pe_fy2028e: number | null;
  ev_ebitda_ttm: number | null;
  ev_ebitda_fy2026e: number | null;
  ev_ebitda_fy2027e: number | null;
  ev_ebitda_fy2028e: number | null;
  ps_ttm: number | null;
  ps_fy2026e: number | null;
  ps_fy2027e: number | null;
  ps_fy2028e: number | null;

  // Historical Valuation
  pe_avg_3yr: number | null;
  pe_avg_5yr: number | null;
  pe_high_hist: number | null;
  pe_low_hist: number | null;

  // Target Prices
  sotp_value: number | null;
  target_price_high: number | null;
  target_price_low: number | null;
  potential_upside_high: number | null;
  potential_upside_low: number | null;
  consensus_target_price: number | null;
  consensus_upside_pct: number | null;

  // Returns
  return_down_from_52w_high: number | null;
  return_up_from_52w_low: number | null;
  return_1m: number | null;
  return_3m: number | null;
  return_6m: number | null;
  return_12m: number | null;

  // Additional fields
  dividend_yield: number | null;
  face_value: number | null;

  // Timestamps
  created_at: string;
  updated_at: string;
}

// EquityUniverse with joined master_company data for display
export interface EquityUniverseWithCompany extends EquityUniverse {
  master_company: {
    company_name: string;
    nse_symbol: string | null;
    bse_code: string | null;
  } | null;
}

// Matches the master_company table
export interface MasterCompany {
  company_id: number;
  company_name: string;
  bse_code: string | null;
  nse_symbol: string | null;
  isin: string | null;
  date_of_listing: string | null; // date stored as ISO string
  paid_up_value: number | null;
  face_value: number | null;
  created_at: string;
  accord_code: string | null;
  google_code: string | null;
  bloomberg_ticker: string | null;
  yahoo_code: string | null;
  modified_at: string | null;
}

export interface AuditLog {
  id?: number;
  user_email: string;
  action: string;
  details: Record<string, unknown>;
  created_at?: string;
}

// Form types for creating new records
export interface CreateMasterCompanyInput {
  company_name: string;
  nse_symbol?: string;
  bse_code?: string;
  isin?: string;
  face_value?: number;
  paid_up_value?: number;
  date_of_listing?: string;
  accord_code?: string;
  google_code?: string;
  bloomberg_ticker?: string;
  yahoo_code?: string;
}

export interface CreateAuditLogInput {
  user_email: string;
  action: string;
  details: Record<string, unknown>;
}

// Research session (matches research_sessions table)
export interface ResearchSession {
  session_id: string;
  company_name: string;
  company_nse_code: string | null;
  sector: string | null;
  current_state: string;
  status: 'document_review' | 'drafting' | 'completed'; // alias for current_state used in UI
  created_by: string;
  pipeline_status: string | null;
  selected_model: string | null;
  total_tokens_used: number;
  generation_time_seconds: number;
  created_at: string;
  updated_at: string;
  [key: string]: unknown; // allow extra columns from DB
}

// Research session with joined master_company data
export interface ResearchSessionWithCompany extends ResearchSession {
  master_company: {
    company_name: string;
    nse_symbol: string | null;
    bse_code: string | null;
    isin: string | null;
  } | null;
}

// Session document (matches session_documents table)
export interface SessionDocument {
  id: string; // UUID
  session_id: string; // UUID FK
  drive_file_id: string;
  file_name: string;
  mime_type: string | null;
  file_size: number | null;
  view_url: string | null;
  download_url: string | null;
  document_type: string | null;
  category: string | null;
  created_at: string;
}

// Input for creating a research session
export interface CreateResearchSessionInput {
  user_email: string;
  nse_symbol: string;
  company_name: string;
  sector: string | null;
  status?: string;
}

// Input for creating session documents
export interface CreateSessionDocumentInput {
  session_id: string;
  drive_file_id: string;
  file_name: string;
  mime_type?: string;
  file_size?: number;
  view_url?: string;
  download_url?: string;
  document_type?: string;
  category?: string;
}

// Research report (matches research_reports table)
export interface ResearchReport {
  report_id: string;
  session_id: string;
  user_email: string;
  company_name: string;
  nse_symbol: string;

  // Text sections (AI-generated - 7 sections)
  company_background: string | null;
  business_model: string | null;
  management_analysis: string | null;
  industry_overview: string | null;
  industry_tailwinds: string | null;
  demand_drivers: string | null;
  industry_risks: string | null;

  // Section headings (AI-generated dynamic headings)
  company_background_h: string | null;
  business_model_h: string | null;
  management_analysis_h: string | null;
  industry_overview_h: string | null;
  industry_tailwinds_h: string | null;
  demand_drivers_h: string | null;
  industry_risks_h: string | null;

  // Metadata
  drive_file_id: string | null;
  drive_file_url: string | null;
  status: 'generating' | 'draft' | 'completed' | 'error';
  tokens_used: number | null;
  generation_time_seconds: number | null;

  // Rating & Recommendation
  recommendation: 'BUY' | 'SELL' | 'HOLD' | null;
  target_price: number | null;
  recommendation_rationale: string | null;

  // PPTX Report (reportgen library output — uploaded to research-reports-pptx bucket)
  pptx_file_path: string | null;
  pptx_file_url: string | null;
  pptx_pdf_file_path: string | null;  // companion PDF (LibreOffice/PowerPoint conversion)
  pptx_pdf_file_url: string | null;
  pptx_generated_at: string | null;
  pptx_status: 'generating' | 'ready' | 'error' | null;

  // @deprecated — legacy HTML/PDF/PPT flow (columns retained for back-compat with /admin/pipelinea and /admin/generate-research).
  // The active flow at /admin/pipeline uses pptx_* fields above.
  ppt_file_id: string | null;
  ppt_file_url: string | null;
  pdf_file_id: string | null;
  pdf_file_url: string | null;
  html_file_path: string | null;
  html_file_url: string | null;
  html_last_edited_at: string | null;

  // Media / Podcast
  podcast_script: string | null;
  audio_file_url: string | null;
  video_script: string | null;
  video_file_url: string | null;

  // Publishing
  is_published: boolean;
  published_at: string | null;
  plan: string | null;

  // Dynamic custom section columns (cs_* prefixed, created via RPC)
  // e.g. cs_valuation_analysis, cs_swot_analysis, etc.
  [key: `cs_${string}`]: string | null;

  // Additional data
  chart_custom: string | null;
  summary_table: string | null;

  created_at: string;
  updated_at: string;
}

// Report chart (matches report_charts table)
export interface ReportChart {
  id: string;
  report_id: string;
  chart_key: ChartSectionKey;
  storage_path: string;
  public_url: string | null;
  caption: string | null;
  display_order: number;
  created_at: string;
  updated_at: string;
}

// Input for creating/updating a report chart
export interface CreateReportChartInput {
  report_id: string;
  chart_key: ChartSectionKey;
  storage_path: string;
  public_url?: string;
  caption?: string;
  display_order?: number;
}

// Prompt template (matches prompt_templates table)
export interface PromptTemplate {
  id: string; // UUID
  section_key: string; // Custom section name
  title: string;
  heading_prompt: string; // Prompt to generate a dynamic heading for this section
  prompt_text: string;
  search_keywords: string[]; // JSON array
  sort_order: number; // Display order (lower = first)
  is_default: boolean;
  user_email: string | null; // null for default templates
  created_at: string;
  updated_at: string;
}

// Input for creating a prompt template
export interface CreatePromptTemplateInput {
  section_key: string;
  title: string;
  prompt_text: string;
  search_keywords: string[];
  user_email: string;
  is_default?: boolean;
}

// API response types
export interface PaginatedResponse<T> {
  data: T[];
  count: number;
  page: number;
  pageSize: number;
}
