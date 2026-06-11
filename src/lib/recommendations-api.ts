import { supabase } from './supabase';
import type { Recommendation, CreateRecommendationPayload } from '@/types/recommendations';

const N8N_BASE_URL = 'https://n8n.tikonacapital.com';

// ========================
// Create
// ========================

export async function createRecommendation(
  payload: CreateRecommendationPayload
): Promise<Recommendation> {
  const { send_telegram, pdf_file_id, send_push, ...dbPayload } = payload;

  const upside_pct =
    dbPayload.cmp && dbPayload.target_price
      ? +((((dbPayload.target_price - dbPayload.cmp) / dbPayload.cmp) * 100).toFixed(2))
      : null;

  const { data, error } = await supabase
    .from('recommendations')
    .insert({
      ...dbPayload,
      upside_pct,
      status: 'active',
      telegram_sent: false,
      is_successful: null,
    })
    .select()
    .single();

  if (error) throw new Error(error.message);
  const rec = data as Recommendation;

  if (send_telegram) {
    try {
      await sendToTelegram(rec, pdf_file_id ?? null);
      await supabase
        .from('recommendations')
        .update({ telegram_sent: true })
        .eq('id', rec.id);
      rec.telegram_sent = true;
    } catch (e) {
      console.error('[Recommendations] Telegram send failed:', e);
    }
  }

  if (send_push) {
    try {
      const { data: functionData, error: functionError } = await supabase.functions.invoke('send-push-notification', {
        body: {
          company_name: rec.company_name,
          nse_symbol: rec.nse_symbol,
          rating: rec.rating,
          cmp: rec.cmp,
          target_price: rec.target_price,
          upside_pct: rec.upside_pct,
          validity_type: rec.validity_type,
          validity_date: rec.validity_date,
          session_id: rec.session_id,
          plans: rec.plans,
        },
      });
      if (functionError) throw functionError;
      console.log('[Recommendations] Push notification sent:', functionData);
    } catch (e) {
      console.error('[Recommendations] Push notification send failed:', e);
    }
  }

  return rec;
}

// ========================
// List
// ========================

export async function listRecommendations(filters?: {
  status?: string;
  plan?: string;
  fromDate?: string;
  toDate?: string;
  createdBy?: string;
}): Promise<Recommendation[]> {
  let query = supabase
    .from('recommendations')
    .select('*')
    .order('created_at', { ascending: false });

  if (filters?.status) query = query.eq('status', filters.status);
  if (filters?.plan) query = query.contains('plans', [filters.plan]);
  if (filters?.fromDate) query = query.gte('created_at', filters.fromDate);
  if (filters?.toDate) query = query.lte('created_at', filters.toDate + 'T23:59:59');
  if (filters?.createdBy) query = query.eq('created_by', filters.createdBy);

  const { data, error } = await query;
  if (error) throw new Error(error.message);
  return (data ?? []) as Recommendation[];
}

// ========================
// Close
// ========================

export async function closeRecommendation(
  id: string,
  is_successful: boolean
): Promise<void> {
  const { error } = await supabase
    .from('recommendations')
    .update({ status: 'closed', is_successful, updated_at: new Date().toISOString() })
    .eq('id', id);
  if (error) throw new Error(error.message);
}

// ========================
// Delete
// ========================

export async function deleteRecommendation(id: string): Promise<void> {
  const { error } = await supabase.from('recommendations').delete().eq('id', id);
  if (error) throw new Error(error.message);
}

// ========================
// Check if recommendation exists for session
// ========================

export async function hasRecommendationForSession(sessionId: string): Promise<boolean> {
  const { count, error } = await supabase
    .from('recommendations')
    .select('id', { count: 'exact', head: true })
    .eq('session_id', sessionId);
  if (error) return false;
  return (count ?? 0) > 0;
}

// ========================
// Resend Telegram
// ========================

export async function resendTelegram(rec: Recommendation): Promise<void> {
  await sendToTelegram(rec);
  await supabase
    .from('recommendations')
    .update({ telegram_sent: true, updated_at: new Date().toISOString() })
    .eq('id', rec.id);
}

// ========================
// Internal: send to n8n → Telegram
// ========================

async function sendToTelegram(rec: Recommendation, pdfFileId?: string | null): Promise<void> {
  const res = await fetch(`${N8N_BASE_URL}/webhook/send-recommendation`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      company_name: rec.company_name,
      nse_symbol: rec.nse_symbol,
      rating: rec.rating,
      cmp: rec.cmp,
      target_price: rec.target_price,
      upside_pct: rec.upside_pct,
      validity_type: rec.validity_type,
      validity_date: rec.validity_date,
      plans: rec.plans,
      trade_notes: rec.trade_notes,
      report_file_url: rec.report_file_url,
      session_id: rec.session_id,
      pdf_file_id: pdfFileId || null,
    }),
  });
  if (!res.ok) {
    throw new Error(`Telegram webhook failed: ${res.status} ${res.statusText}`);
  }
}
