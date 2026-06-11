// Recommendation types — Tikona Capital
// Tracks equity research recommendations sent to subscribers via Telegram

export type RecommendationRating =
  | 'STRONG BUY'
  | 'BUY'
  | 'ACCUMULATE'
  | 'HOLD'
  | 'UNDERPERFORM'
  | 'SELL';

export type RecommendationStatus = 'active' | 'closed';
export type ValidityType = '1_year' | 'custom';

export const PLAN_OPTIONS = [
  { id: 'midcap_wealth',  label: 'Mid Cap Wealth Builders' },
  { id: 'smallcap_alpha', label: 'Smallcap Alpha Picks' },
  { id: 'sme_emerging',   label: 'SME Emerging Business' },
] as const;

export type PlanId = (typeof PLAN_OPTIONS)[number]['id'];

export const RATING_CONFIG: Record<
  RecommendationRating,
  { color: string; bg: string; border: string; dot: string }
> = {
  'STRONG BUY': { color: 'text-emerald-700', bg: 'bg-emerald-50', border: 'border-emerald-200', dot: 'bg-emerald-500' },
  'BUY':        { color: 'text-green-700',   bg: 'bg-green-50',   border: 'border-green-200',   dot: 'bg-green-500' },
  'ACCUMULATE': { color: 'text-teal-700',    bg: 'bg-teal-50',    border: 'border-teal-200',    dot: 'bg-teal-500' },
  'HOLD':       { color: 'text-amber-700',   bg: 'bg-amber-50',   border: 'border-amber-200',   dot: 'bg-amber-500' },
  'UNDERPERFORM':{ color: 'text-orange-700', bg: 'bg-orange-50',  border: 'border-orange-200',  dot: 'bg-orange-500' },
  'SELL':       { color: 'text-red-700',     bg: 'bg-red-50',     border: 'border-red-200',     dot: 'bg-red-500' },
};

export interface Recommendation {
  id: string;
  company_name: string;
  nse_symbol: string;
  rating: RecommendationRating;
  cmp: number | null;
  target_price: number;
  upside_pct: number | null;
  validity_type: ValidityType;
  validity_date: string | null;
  plans: PlanId[];
  trade_notes: string | null;
  report_file_url: string | null;
  session_id: string | null;
  status: RecommendationStatus;
  is_successful: boolean | null;
  telegram_sent: boolean;
  created_by: string | null;
  created_at: string;
  updated_at: string;
}

export interface CreateRecommendationPayload {
  company_name: string;
  nse_symbol: string;
  rating: RecommendationRating;
  cmp: number | null;
  target_price: number;
  validity_type: ValidityType;
  validity_date: string | null;
  plans: PlanId[];
  trade_notes: string | null;
  report_file_url: string | null;
  session_id: string | null;
  send_telegram: boolean;
  send_push?: boolean;
  created_by: string | null;
  pdf_file_id?: string | null;
}
