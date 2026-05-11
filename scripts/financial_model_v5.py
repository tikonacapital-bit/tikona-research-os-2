"""
═══════════════════════════════════════════════════════════════════
 TIKONA CAPITAL — Financial Model Generator v5.1
═══════════════════════════════════════════════════════════════════
 Model:    claude-sonnet-4-20250514 (Claude Sonnet 4) + web_search
 Changes:  Year-agnostic prompt, formula-based Excel, cost tracking
 Output:   Screener Excel + 8 appended model sheets (all formulas)

 Refactored from Colab notebook into a callable function so
 financial_model_server.py (FastAPI) can import and invoke it.

 Required environment variables:
   ANTHROPIC_API_KEY     - Anthropic API key
   SCREENER_USERNAME     - Screener.in account email
   SCREENER_PASSWORD     - Screener.in password

 Public API:
   generate_financial_model(nse_code, company_name, sector, output_dir)
       -> {file_path, json_path, model_json, cost_summary, elapsed_seconds}
═══════════════════════════════════════════════════════════════════
"""

from __future__ import annotations

import os
import re
import json
import time
import shutil
import logging
import traceback
from typing import Optional

import anthropic
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup
import requests as req_lib


logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s")
logger = logging.getLogger(__name__)

MODEL_NAME = "claude-sonnet-4-20250514"
WEB_SEARCH_TOOL = [{"type": "web_search_20250305", "name": "web_search"}]


# ══════════════════════════════════════════════════════════════════
# COST TRACKER
# ══════════════════════════════════════════════════════════════════
class CostTracker:
    INPUT_COST_PER_M = 3.00
    OUTPUT_COST_PER_M = 15.00
    USD_TO_INR = 84.0

    def __init__(self):
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self.total_cache_read = 0
        self.total_cache_create = 0
        self.calls = 0

    def track(self, response):
        usage = response.usage
        self.total_input_tokens += usage.input_tokens
        self.total_output_tokens += usage.output_tokens
        if hasattr(usage, "cache_read_input_tokens"):
            self.total_cache_read += usage.cache_read_input_tokens or 0
        if hasattr(usage, "cache_creation_input_tokens"):
            self.total_cache_create += usage.cache_creation_input_tokens or 0
        self.calls += 1

    @property
    def total_tokens(self):
        return self.total_input_tokens + self.total_output_tokens

    @property
    def cost_usd(self):
        return (
            (self.total_input_tokens / 1_000_000) * self.INPUT_COST_PER_M
            + (self.total_output_tokens / 1_000_000) * self.OUTPUT_COST_PER_M
        )

    @property
    def cost_inr(self):
        return self.cost_usd * self.USD_TO_INR

    def to_dict(self):
        return {
            "model": MODEL_NAME,
            "calls": self.calls,
            "input_tokens": self.total_input_tokens,
            "output_tokens": self.total_output_tokens,
            "total_tokens": self.total_tokens,
            "cache_read": self.total_cache_read,
            "cache_create": self.total_cache_create,
            "cost_usd": round(self.cost_usd, 4),
            "cost_inr": round(self.cost_inr, 2),
        }

    def summary(self):
        return (
            f"\n{'='*60}\n"
            f"  💰 API COST SUMMARY\n"
            f"{'='*60}\n"
            f"  Model:           {MODEL_NAME}\n"
            f"  API Calls:       {self.calls}\n"
            f"  Input Tokens:    {self.total_input_tokens:,}\n"
            f"  Output Tokens:   {self.total_output_tokens:,}\n"
            f"  Total Tokens:    {self.total_tokens:,}\n"
            f"  Cache Read:      {self.total_cache_read:,}\n"
            f"  Cache Create:    {self.total_cache_create:,}\n"
            f"  ─────────────────────────────\n"
            f"  Cost (USD):      ${self.cost_usd:.4f}\n"
            f"  Cost (INR):      ₹{self.cost_inr:.2f}\n"
            f"{'='*60}"
        )


# ══════════════════════════════════════════════════════════════════
# STYLES
# ══════════════════════════════════════════════════════════════════
COLORS = {
    "navy": "1F4690", "med_blue": "3A5BA0", "orange": "FFA500", "peach": "FFE5B4",
    "white": "FFFFFF", "light_grey": "F2F2F2", "black": "000000", "red": "D32F2F",
    "green": "2E7D32", "dark_grey": "D9D9D9",
}

navy_fill = PatternFill("solid", fgColor=COLORS["navy"])
blue_fill = PatternFill("solid", fgColor=COLORS["med_blue"])
orange_fill = PatternFill("solid", fgColor=COLORS["orange"])
peach_fill = PatternFill("solid", fgColor=COLORS["peach"])
grey_fill = PatternFill("solid", fgColor=COLORS["light_grey"])

hdr_font = Font(name="Arial", bold=True, color=COLORS["white"], size=11)
sub_font = Font(name="Arial", bold=True, color=COLORS["white"], size=10)
title_font = Font(name="Arial", bold=True, color=COLORS["navy"], size=14)
sec_font = Font(name="Arial", bold=True, color=COLORS["navy"], size=11)
data_font = Font(name="Arial", color=COLORS["black"], size=10)
input_font = Font(name="Arial", color=COLORS["navy"], size=10)
red_font = Font(name="Arial", color=COLORS["red"], size=10)

thin_bdr = Border(
    left=Side(style="thin", color=COLORS["dark_grey"]),
    right=Side(style="thin", color=COLORS["dark_grey"]),
    top=Side(style="thin", color=COLORS["dark_grey"]),
    bottom=Side(style="thin", color=COLORS["dark_grey"]),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
INR = "#,##,##0"
INR2 = "#,##,##0.00"
PCT = "0.00%"
RATIO = "0.00"


# ══════════════════════════════════════════════════════════════════
# SCREENER DOWNLOADER
# ══════════════════════════════════════════════════════════════════
def download_screener_excel(symbol: str, output_dir: str, screener_username: str, screener_password: str) -> str:
    import cloudscraper

    logger.info("📡 Logging into Screener.in...")
    session = cloudscraper.create_scraper(browser="chrome")
    login_url = "https://www.screener.in/login/"
    resp = session.get(login_url, timeout=15)
    soup = BeautifulSoup(resp.text, "html.parser")
    csrf = soup.find("input", {"name": "csrfmiddlewaretoken"})
    csrf_token = csrf["value"] if csrf else session.cookies.get("csrftoken")

    post_resp = session.post(
        login_url,
        data={
            "username": screener_username,
            "password": screener_password,
            "csrfmiddlewaretoken": csrf_token,
            "next": "/",
        },
        headers={"Referer": login_url},
        timeout=15,
    )
    if "Logout" not in post_resp.text and "account" not in post_resp.url:
        raise Exception("❌ Screener login failed — check SCREENER_USERNAME / SCREENER_PASSWORD env vars")
    logger.info("✅ Screener login OK")

    company_url = None
    page_resp = None
    for suffix in [f"{symbol}/consolidated/", f"{symbol}/"]:
        url = f"https://www.screener.in/company/{suffix}"
        resp = session.get(url, timeout=15)
        if resp.status_code == 200:
            company_url = resp.url  # use final URL after any redirects
            page_resp = resp
            break
    else:
        raise Exception(f"❌ {symbol} not found on Screener")

    logger.info(f"📄 Company page resolved to: {company_url}")

    soup = BeautifulSoup(page_resp.text, "html.parser")

    # Dump all forms and export-related strings from raw HTML for diagnosis
    for f in soup.find_all("form"):
        logger.info(f"  [FORM] action={f.get('action')} method={f.get('method')}")
    for tag in soup.find_all(True):
        for attr in ("href", "action", "formaction", "data-url"):
            val = tag.get(attr, "")
            if val and "export" in val.lower():
                logger.info(f"  [EXPORT-ATTR] <{tag.name} {attr}={val}>")
    import re as _re
    for m in _re.findall(r'["\']([^"\']{5,120}export[^"\']{0,60})["\']', page_resp.text, _re.I)[:8]:
        logger.info(f"  [EXPORT-RAW] {m}")

    # Extract CSRF token from page HTML first (more reliable than cookie)
    page_csrf_input = soup.find("input", {"name": "csrfmiddlewaretoken"})
    csrf2_token = page_csrf_input["value"] if page_csrf_input else session.cookies.get("csrftoken", "")

    export_url = None

    # Try button with aria-label
    btn = soup.find("button", attrs={"aria-label": "Export to Excel"})
    # Try button with export text (case-insensitive)
    if not btn:
        btn = soup.find("button", string=lambda t: t and "export" in str(t).lower())
    # Try anchor with export in href
    if not btn:
        btn = soup.find("a", href=lambda h: h and "export" in str(h).lower())
    # Try form whose action contains "export"
    if not btn:
        form = soup.find("form", action=lambda a: a and "export" in str(a).lower())
        if form:
            action = form.get("action", "")
            export_url = action if action.startswith("http") else "https://www.screener.in" + action
            ci = form.find("input", {"name": "csrfmiddlewaretoken"})
            if ci:
                csrf2_token = ci["value"]

    if btn and not export_url:
        raw = btn.get("formaction", "") or btn.get("href", "")
        export_url = raw if raw.startswith("http") else "https://www.screener.in" + raw
        form = btn.find_parent("form")
        ci = (form or soup).find("input", {"name": "csrfmiddlewaretoken"})
        if ci:
            csrf2_token = ci["value"]

    # Screener now uses a no-action POST form — submits back to the company page URL
    if not export_url:
        for f in soup.find_all("form"):
            if f.get("action") is None and (f.get("method") or "").lower() == "post":
                export_url = company_url
                ci = f.find("input", {"name": "csrfmiddlewaretoken"})
                if ci:
                    csrf2_token = ci["value"]
                logger.info(f"  Found no-action POST form → submitting to {export_url}")
                break

    # Last-resort fallback
    if not export_url:
        export_url = company_url
        logger.warning(f"⚠️  No export form found; POSTing to company URL: {export_url}")

    logger.info(f"📥 POSTing export: {export_url}")
    file_resp = session.post(
        export_url,
        data={"csrfmiddlewaretoken": csrf2_token},
        headers={
            "Referer": company_url,
            "Origin": "https://www.screener.in",
            "X-CSRFToken": csrf2_token,
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
        },
        timeout=30,
    )
    content_type = file_resp.headers.get("Content-Type", "").lower()
    if "html" in content_type:
        snippet = file_resp.text[:300].replace("\n", " ")
        raise Exception(
            f"❌ Got HTML instead of Excel from Screener export "
            f"(status={file_resp.status_code}, url={file_resp.url}, preview: {snippet})"
        )

    out_path = os.path.join(output_dir, f"{symbol}_screener.xlsx")
    with open(out_path, "wb") as f:
        f.write(file_resp.content)
    logger.info(f"✅ Downloaded → {out_path} ({len(file_resp.content)//1024} KB)")
    return out_path


# ══════════════════════════════════════════════════════════════════
# DATA EXTRACTOR
# ══════════════════════════════════════════════════════════════════
def extract_screener_data(filepath: str) -> dict:
    wb = load_workbook(filepath, data_only=True)
    if "Data Sheet" not in wb.sheetnames:
        raise Exception("❌ No 'Data Sheet' in Screener Excel")
    ws = wb["Data Sheet"]

    company_name = ws.cell(row=1, column=2).value or ""
    face_value = ws.cell(row=7, column=2).value or 1
    cmp = ws.cell(row=8, column=2).value or 0
    mcap = ws.cell(row=9, column=2).value or 0

    data_cols = [col for col in range(2, 20) if ws.cell(row=16, column=col).value is not None]
    fiscal_years = []
    for col in data_cols:
        dt = ws.cell(row=16, column=col).value
        fiscal_years.append(f"FY{str(dt.year)[2:]}" if hasattr(dt, "year") else str(dt))

    def rd(row_num):
        return [ws.cell(row=row_num, column=col).value for col in data_cols]

    nh = len(data_cols)
    sales = rd(17)
    rm = rd(18)

    data = {
        "company_name": company_name,
        "face_value": face_value,
        "cmp": cmp,
        "mcap": mcap,
        "fiscal_years": fiscal_years,
        "data_cols": data_cols,
        "num_years": nh,
        "pl": {
            "sales": sales, "raw_material": rm,
            "change_in_inventory": rd(19), "power_fuel": rd(20), "other_mfg": rd(21),
            "employee_cost": rd(22), "selling_admin": rd(23), "other_expenses": rd(24),
            "other_income": rd(25), "depreciation": rd(26), "interest": rd(27),
            "pbt": rd(28), "tax": rd(29), "net_profit": rd(30), "dividend_amount": rd(31),
        },
        "bs": {
            "share_capital": rd(57), "reserves": rd(58), "borrowings": rd(59),
            "other_liabilities": rd(60), "total_liab": rd(61), "net_block": rd(62),
            "cwip": rd(63), "investments": rd(64), "other_assets": rd(65),
            "total_assets": rd(66), "receivables": rd(67), "inventory": rd(68),
            "cash": rd(69), "shares_outstanding": rd(70),
        },
        "cf": {"cfo": rd(82), "cfi": rd(83), "cff": rd(84), "net_cash_flow": rd(85)},
    }
    data["derived"] = {
        "ebitda": [
            (sales[i] or 0) - sum(filter(None, [
                rm[i], data["pl"]["change_in_inventory"][i], data["pl"]["power_fuel"][i],
                data["pl"]["other_mfg"][i], data["pl"]["employee_cost"][i],
                data["pl"]["selling_admin"][i], data["pl"]["other_expenses"][i],
            ])) if sales[i] else None
            for i in range(nh)
        ]
    }
    wb.close()
    logger.info(f"✅ Extracted {nh} years: {fiscal_years}")
    return data


# ══════════════════════════════════════════════════════════════════
# MARKET DATA
# ══════════════════════════════════════════════════════════════════
def fetch_market_data(ticker: str) -> str:
    parts = []
    try:
        import yfinance as yf
        info = yf.Ticker(f"{ticker}.NS").info
        parts.append(
            f"yfinance | CMP:₹{info.get('currentPrice','NA')} "
            f"MCap:₹{(info.get('marketCap',0) or 0)/1e7:.0f}Cr "
            f"PE:{info.get('trailingPE','NA')} PB:{info.get('priceToBook','NA')} "
            f"ROE:{((info.get('returnOnEquity',0) or 0)*100):.1f}% "
            f"Beta:{info.get('beta','NA')} 52WH:₹{info.get('fiftyTwoWeekHigh','NA')} "
            f"52WL:₹{info.get('fiftyTwoWeekLow','NA')}"
        )
    except Exception as e:
        parts.append(f"yfinance: {e}")

    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        for u in [
            f"https://www.screener.in/company/{ticker}/consolidated/",
            f"https://www.screener.in/company/{ticker}/",
        ]:
            r = req_lib.get(u, headers=headers, timeout=10)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "html.parser")
                rows = []
                for t in soup.find_all("section", class_="card"):
                    h = t.find("h2")
                    if h:
                        rows.append(f"\n=== {h.get_text(strip=True)} ===")
                    for row in t.find_all("tr"):
                        cells = row.find_all(["th", "td"])
                        if cells:
                            rows.append(" | ".join(c.get_text(strip=True) for c in cells))
                if rows:
                    parts.append("\n".join(rows[:60]))
                break
    except Exception:
        pass
    return "\n\n".join(parts)


# ══════════════════════════════════════════════════════════════════
# JSON EXTRACTION
# ══════════════════════════════════════════════════════════════════
def extract_json_from_text(text: str) -> dict:
    text = re.sub(r"<thinking>.*?</thinking>", "", text, flags=re.DOTALL).strip()
    if text.startswith("{"):
        try:
            return json.loads(text)
        except Exception:
            pass
    for pat in [r"```json\s*(.*?)\s*```", r"```\s*(\{.*?\})\s*```"]:
        m = re.search(pat, text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except Exception:
                pass
    start = text.find("{")
    if start != -1:
        depth = 0
        for i, ch in enumerate(text[start:], start):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start : i + 1])
                    except Exception:
                        break
    raise ValueError(f"No valid JSON. Preview: {text[:500]}")


# ══════════════════════════════════════════════════════════════════
# CLAUDE API — YEAR-AGNOSTIC PROMPT
# ══════════════════════════════════════════════════════════════════
def call_claude_api(
    company_name: str,
    ticker: str,
    sector: str,
    screener_data: dict,
    market_text: str,
    api_key: str,
    tracker: CostTracker,
) -> dict:
    client = anthropic.Anthropic(api_key=api_key)

    hist_json = json.dumps(
        {
            "years": screener_data["fiscal_years"],
            "revenue": screener_data["pl"]["sales"],
            "raw_material": screener_data["pl"]["raw_material"],
            "change_in_inventory": screener_data["pl"]["change_in_inventory"],
            "power_fuel": screener_data["pl"]["power_fuel"],
            "other_mfg": screener_data["pl"]["other_mfg"],
            "employee_cost": screener_data["pl"]["employee_cost"],
            "selling_admin": screener_data["pl"]["selling_admin"],
            "other_expenses": screener_data["pl"]["other_expenses"],
            "ebitda": screener_data["derived"]["ebitda"],
            "net_profit": screener_data["pl"]["net_profit"],
            "depreciation": screener_data["pl"]["depreciation"],
            "interest": screener_data["pl"]["interest"],
            "other_income": screener_data["pl"]["other_income"],
            "pbt": screener_data["pl"]["pbt"],
            "tax": screener_data["pl"]["tax"],
            "share_capital": screener_data["bs"]["share_capital"],
            "reserves": screener_data["bs"]["reserves"],
            "borrowings": screener_data["bs"]["borrowings"],
            "other_liabilities": screener_data["bs"]["other_liabilities"],
            "net_block": screener_data["bs"]["net_block"],
            "cwip": screener_data["bs"]["cwip"],
            "investments": screener_data["bs"]["investments"],
            "receivables": screener_data["bs"]["receivables"],
            "inventory": screener_data["bs"]["inventory"],
            "cash": screener_data["bs"]["cash"],
            "total_assets": screener_data["bs"]["total_assets"],
            "cfo": screener_data["cf"]["cfo"],
            "cfi": screener_data["cf"]["cfi"],
            "cff": screener_data["cf"]["cff"],
        },
        indent=2,
        default=str,
    )

    prompt = f"""You are a senior equity research analyst building a complete financial model.

COMPANY: {company_name} ({ticker}) | Sector: {sector}

━━ HISTORICAL DATA (Screener.in audited, ₹ Crores) ━━
{hist_json[:6000]}

━━ MARKET DATA ━━
{market_text[:3000]}

━━ INSTRUCTIONS ━━

STEP 0 — DETERMINE TIMELINE:
- Look at the historical data above. Identify the LAST fiscal year with actual data (call it "Base Year").
- Your projection period = next 5 fiscal years AFTER the Base Year.
- Example: if last actual year is FY25 → project FY26E, FY27E, FY28E, FY29E, FY30E
- Example: if last actual year is FY24 → project FY25E, FY26E, FY27E, FY28E, FY29E
- Use this dynamically — do NOT assume any specific year.

STEP 1 — WEB SEARCHES (use current/recent context, not hardcoded years):
- "{company_name} latest quarterly results earnings revenue profit"
- "{company_name} most recent annual results full year earnings"
- "{company_name} management guidance outlook capacity expansion next 3 years"
- "{company_name} analyst consensus target price estimates"
- "{company_name} broker reports earnings estimates next 5 years"
- "{sector} sector India peer comparison PE EV/EBITDA current multiples"
- "India 10 year government bond yield current"
- "{ticker} consensus estimates earnings forecasts"
- "{ticker} GoIndiaStocks consensus estimates"
- "site:trendlyne.com {ticker} forecasts estimates"
- "{company_name} order book revenue visibility pipeline"

STEP 2 — BUILD ASSUMPTIONS (5 projected years):
For EVERY projected year, provide these as PLAIN NUMBERS (15.5 not 0.155):
- revenue_growth_pct, rm_pct, employee_pct, power_fuel_pct, other_mfg_pct,
  selling_admin_pct, other_exp_pct, chg_inventory_pct
- depreciation_cr, interest_cr, other_income_cr (₹ Cr absolute)
- tax_rate_pct, capex_cr, receivable_days, inventory_days

Excel formulas applied by the builder:
  Revenue Year N = Revenue Year N-1 × (1 + revenue_growth_pct / 100)
  Each expense = Revenue × pct / 100
  EBITDA = Revenue - all operating expenses
  PBT = EBITDA + Other Income - Depreciation - Interest
  Tax = PBT × tax_rate_pct / 100
  PAT = PBT - Tax
  Receivables = Revenue × receivable_days / 365
  Inventory = RM Cost × inventory_days / 365
Every assumption % MUST be realistic and internally consistent.

STEP 3 — BALANCE SHEET PROJECTIONS (absolute values):
- net_block, cwip, investments, other_assets
- share_capital, reserves, borrowings, other_liabilities
Cash = PLUG. Balance Sheet MUST balance every year.

STEP 4 — HISTORICAL FINANCIALS IN ABSOLUTE TERMS (last 6 years):
Use web_search to find the company's annual revenue, EBITDA, and PAT in ₹ Crores for each
of the last 6 fiscal years. This is CRITICAL for recently-IPO'd or platform businesses
(e.g., Swiggy, Zomato) where the Screener data above may show nulls. Search for:
  "{company_name} annual revenue EBITDA PAT FY20 FY21 FY22 FY23 FY24 FY25 crores"
  "{company_name} annual report revenue profit loss"

STEP 5 — HISTORICAL RATIOS (last 6 years):
- ebitda_margin_pct, pat_margin_pct, roe_pct, roce_pct
- debt_equity, receivable_days, inventory_days, asset_turnover
- rm_pct, employee_pct, tax_rate_pct

STEP 5 — CASH FLOW PROJECTIONS: CFO, CFI, CFF for each projected year.

STEP 6 — VALUATION (3 methods):
- DCF (FCFE discounted, terminal via Gordon Growth)
- PE-based (Target PE × EPS year-2)
- EV/EBITDA (Target × EBITDA year-2 - net debt)
- Blended (DCF 40%, PE 30%, EV/EBITDA 30%)
- Sensitivity 5x5 grid: PE Multiple vs EPS Growth %

STEP 7 — PEER COMPARISON: 4+ listed peers with current MCap, PE, EV/EBITDA, EBITDA Margin, ROE.

STEP 8 — INVESTMENT THESIS with SAARTHI scoring (S+A+A+R+T+H+I = 100 max):
  S—Sector(0-15) A—Accounting(0-15) A—Asset(0-15) R—Revenue(0-15)
  T—Track Record(0-10) H—BS Health(0-15) I—Valuation(0-15)
  STRONG BUY≥80 | BUY 65-79 | ACCUMULATE 55-64 | HOLD 45-54 | UNDERPERFORM 35-44 | SELL<35

━━ RULES ━━
1. All monetary values in ₹ Crores. EPS/Target in ₹/share.
2. Percentages as PLAIN NUMBERS: 15.5, NOT 0.155.
3. null for missing. NEVER fabricate historical numbers.
4. BS must balance. Cash is plug.
5. Projection years = 5 years after last actual year.
6. shares_cr = share_capital of last year ÷ face_value ({screener_data['face_value']}).
7. Provide ALL individual expense % — builder NEEDS them for Excel formulas.

━━ OUTPUT ━━
Return ONLY valid JSON. Start with {{ end with }}. No prose, no markdown fences.

{{
  "sector": str,
  "base_year": str,
  "cmp": number,
  "target_price": number,
  "rating": "STRONG BUY|BUY|ACCUMULATE|HOLD|UNDERPERFORM|SELL",
  "upside_pct": number,
  "shares_cr": number,
  "assumptions": {{
    "projection_years": ["FY__E", ...5],
    "revenue_growth_pct": {{"FY__E": number, ...5}},
    "rm_pct": {{...}}, "employee_pct": {{...}}, "power_fuel_pct": {{...}},
    "other_mfg_pct": {{...}}, "selling_admin_pct": {{...}}, "other_exp_pct": {{...}},
    "chg_inventory_pct": {{...}},
    "depreciation_cr": {{...}}, "interest_cr": {{...}}, "other_income_cr": {{...}},
    "tax_rate_pct": {{...}}, "capex_cr": {{...}},
    "receivable_days": {{...}}, "inventory_days": {{...}},
    "dividend_payout_pct": number,
    "target_pe": number, "target_ev_ebitda": number,
    "wacc_pct": number, "terminal_growth_pct": number,
    "rationale": str
  }},
  "projections": {{
    "years": ["FY__E", ...5],
    "net_block":[5], "cwip":[5], "investments":[5], "other_assets":[5],
    "share_capital":[5], "reserves":[5], "borrowings":[5], "other_liabilities":[5],
    "cfo":[5], "cfi":[5], "cff":[5]
  }},
  "historical_ratios": {{
    "years": [last 6 FY],
    "ebitda_margin_pct":[6], "pat_margin_pct":[6], "roe_pct":[6], "roce_pct":[6],
    "debt_equity":[6], "receivable_days":[6], "inventory_days":[6], "asset_turnover":[6],
    "rm_pct":[6], "employee_pct":[6], "tax_rate_pct":[6]
  }},
  "historical_pl_absolute": {{
    "years": [last 6 FY — same order as historical_ratios.years],
    "revenue": [6 values in ₹ Cr, use null only if truly unavailable after web search],
    "ebitda":  [6 values in ₹ Cr, null if unavailable],
    "pat":     [6 values in ₹ Cr, null if unavailable]
  }},
  "valuation": {{
    "dcf_fair_value":number, "pe_fair_value":number,
    "ev_ebitda_fair_value":number, "blended_fair_value":number,
    "sensitivity_pe": {{
      "row_label":"PE Multiple", "col_label":"EPS Growth %",
      "row_values":[5], "col_values":[5], "grid":[[5x5]]
    }}
  }},
  "peers": [{{"name":str,"mcap_cr":n,"pe":n,"ev_ebitda":n,"ebitda_margin_pct":str,"roe_pct":n}}],
  "thesis": {{
    "investment_thesis":str, "bull_case":str, "bear_case":str,
    "key_catalysts":[4], "key_risks":[4],
    "saarthi_scores":{{
      "S_sector_quality":n,"A_accounting_quality":n,"A_asset_quality":n,
      "R_revenue_visibility":n,"T_track_record":n,"H_balance_sheet_health":n,"I_intrinsic_valuation":n
    }},
    "saarthi_total":number, "saarthi_rating":str
  }}
}}"""

    logger.info(f"🤖 Calling Claude API ({MODEL_NAME}) with web_search...")
    messages = [{"role": "user", "content": prompt}]
    raw = ""

    for iteration in range(30):
        response = client.messages.create(
            model=MODEL_NAME, max_tokens=16000, messages=messages, tools=WEB_SEARCH_TOOL,
        )
        tracker.track(response)

        if response.stop_reason == "end_turn":
            raw = "".join(b.text for b in response.content if hasattr(b, "text") and b.text)
            break
        elif response.stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": response.content})
            tool_results = [
                {"type": "tool_result", "tool_use_id": b.id, "content": ""}
                for b in response.content if hasattr(b, "type") and b.type == "tool_use"
            ]
            if tool_results:
                messages.append({"role": "user", "content": tool_results})
            logger.info(f"   🔍 Web search #{iteration+1}...")
        else:
            raw = "".join(b.text for b in response.content if hasattr(b, "text") and b.text)
            break
    else:
        raw = "".join(b.text for b in response.content if hasattr(b, "text") and b.text)

    logger.info(
        f"✅ Claude done | Tokens: in={tracker.total_input_tokens:,} out={tracker.total_output_tokens:,}"
    )

    try:
        return extract_json_from_text(raw)
    except ValueError:
        logger.warning("⚠️ JSON parse failed, requesting repair...")
        repair = client.messages.create(
            model=MODEL_NAME, max_tokens=16000,
            messages=[{
                "role": "user",
                "content": f"Fix into valid JSON starting with {{ ending with }}:\n{raw[:12000]}",
            }],
        )
        tracker.track(repair)
        return extract_json_from_text(
            "".join(b.text for b in repair.content if hasattr(b, "text"))
        )


# ══════════════════════════════════════════════════════════════════
# FORMULA-BASED EXCEL BUILDER
# ══════════════════════════════════════════════════════════════════
def build_model(screener_path: str, screener_data: dict, model: dict, out_path: str):
    is_xlsm = screener_path.endswith(".xlsm")
    wb = load_workbook(screener_path, keep_vba=is_xlsm)
    orig_sheets = list(wb.sheetnames)
    logger.info(f"📂 Loaded: {len(orig_sheets)} sheets")

    def mk(name):
        n = name + " - Model" if name in wb.sheetnames else name
        return wb.create_sheet(n), n

    def hdr_r(ws, r, labels, nc):
        for ci, lbl in enumerate(labels, 1):
            c = ws.cell(row=r, column=ci, value=lbl)
            c.fill = navy_fill
            c.font = hdr_font
            c.alignment = center
            c.border = thin_bdr

    def set_w(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def add_borders(ws, mr, mc):
        for row in ws.iter_rows(min_row=1, max_row=mr, min_col=1, max_col=mc):
            for cell in row:
                cell.border = thin_bdr

    sd = screener_data
    hist_years = sd["fiscal_years"]
    data_cols = sd["data_cols"]
    ds_letters = [get_column_letter(c) for c in data_cols]
    nh = len(hist_years)

    asmp = model.get("assumptions", {})
    proj = model.get("projections", {})
    proj_years = asmp.get("projection_years", proj.get("years", ["FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]))
    np_ = len(proj_years)

    dh = min(nh, 6)
    h_start = nh - dh
    disp_hist = hist_years[h_start:]
    disp_ds = ds_letters[h_start:]

    year_labels = ["Particulars"] + disp_hist + proj_years
    nc = len(year_labels)
    h_cols = list(range(2, 2 + dh))
    p_cols = list(range(2 + dh, nc + 1))

    shares_cr = model.get("shares_cr", sd["bs"]["share_capital"][-1] or 10)
    thesis = model.get("thesis", {})
    val = model.get("valuation", {})
    peers = model.get("peers", [])
    rating = model.get("rating", "HOLD")
    target = model.get("target_price", 0)
    cmp = model.get("cmp", sd["cmp"])
    upside = model.get("upside_pct", 0)
    hist_ratios = model.get("historical_ratios", {})

    def av(key, yr):
        d = asmp.get(key, {})
        if isinstance(d, dict):
            return d.get(yr, 0)
        return d or 0

    def pv(key, i, default=0):
        vals = proj.get(key, [])
        return vals[i] if i < len(vals) and vals[i] is not None else default

    # ═══ COVER ═══
    ws, _ = mk("Cover")
    set_w(ws, [25, 30, 5, 25, 30])
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = (sd["company_name"] or "").upper()
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=20)
    c.fill = navy_fill
    c.alignment = center
    ws.row_dimensions[1].height = 45
    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value = f"{model.get('sector','')} | TIKONA CAPITAL"
    c.font = Font(name="Arial", color="FFFFFF", size=12)
    c.fill = blue_fill
    c.alignment = center

    for i, (k, v) in enumerate(
        [
            ("CMP (₹)", cmp), ("Target (₹)", target), ("Upside", f"{upside:.1f}%"),
            ("Rating", rating), ("Market Cap (₹ Cr)", sd["mcap"]),
            ("Shares (Cr)", f"{shares_cr:.2f}"), ("Base Year", model.get("base_year", "")),
        ], 4,
    ):
        ws.cell(row=i, column=1, value=k).font = sec_font
        ws.cell(row=i, column=2, value=v).font = data_font
        for c in range(1, 3):
            ws.cell(row=i, column=c).border = thin_bdr

    r = 12
    ws.merge_cells(f"A{r}:E{r}")
    ws[f"A{r}"].value = f"RATING: {rating} | Target: ₹{target:,.0f} | Upside: {upside:.1f}%"
    color = COLORS["green"] if upside > 0 else COLORS["red"]
    ws[f"A{r}"].font = Font(name="Arial", bold=True, color="FFFFFF", size=14)
    ws[f"A{r}"].fill = PatternFill("solid", fgColor=color)
    ws[f"A{r}"].alignment = center
    r = 14
    ws.merge_cells(f"A{r}:E{r+3}")
    ws[f"A{r}"].value = thesis.get("investment_thesis", "")
    ws[f"A{r}"].font = data_font
    ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    r = 19
    for label, items, col in [
        ("KEY CATALYSTS", thesis.get("key_catalysts", []), COLORS["green"]),
        ("KEY RISKS", thesis.get("key_risks", []), COLORS["red"]),
    ]:
        ws[f"A{r}"].value = label
        ws[f"A{r}"].font = Font(name="Arial", bold=True, color=col, size=12)
        r += 1
        for item in items:
            ws[f"A{r}"].value = f"• {item}"
            r += 1
        r += 1
    ws[f"A{r}"].value = (
        f"SAARTHI SCORE: {thesis.get('saarthi_total','NA')}/100 → {thesis.get('saarthi_rating','NA')}"
    )
    ws[f"A{r}"].font = Font(name="Arial", bold=True, color=COLORS["navy"], size=13)
    ws[f"A{r}"].fill = PatternFill("solid", fgColor=COLORS["peach"])
    logger.info("  ✅ Cover")

    # ═══ ASSUMPTIONS ═══
    ws, asn = mk("Assumptions")
    set_w(ws, [35] + [15] * (nc - 1))
    hdr_r(ws, 1, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=1, column=ci).fill = orange_fill

    assum_items = [
        ("Revenue Growth %", "revenue_growth_pct", None, RATIO),
        ("RM % of Revenue", "rm_pct", "rm_pct", RATIO),
        ("Employee % of Revenue", "employee_pct", "employee_pct", RATIO),
        ("Power & Fuel %", "power_fuel_pct", None, RATIO),
        ("Other Mfg %", "other_mfg_pct", None, RATIO),
        ("Selling & Admin %", "selling_admin_pct", None, RATIO),
        ("Other Expenses %", "other_exp_pct", None, RATIO),
        ("Chg in Inventory %", "chg_inventory_pct", None, RATIO),
        ("Depreciation (₹ Cr)", "depreciation_cr", None, INR),
        ("Interest (₹ Cr)", "interest_cr", None, INR),
        ("Other Income (₹ Cr)", "other_income_cr", None, INR),
        ("Tax Rate %", "tax_rate_pct", "tax_rate_pct", RATIO),
        ("Capex (₹ Cr)", "capex_cr", None, INR),
        ("Receivable Days", "receivable_days", "receivable_days", "0"),
        ("Inventory Days", "inventory_days", "inventory_days", "0"),
    ]
    for r_idx, (label, assum_key, hist_key, fmt) in enumerate(assum_items, 2):
        ws.cell(row=r_idx, column=1, value=label).font = data_font
        if hist_key and hist_key in hist_ratios:
            hr_vals = hist_ratios[hist_key]
            hr_years = hist_ratios.get("years", [])
            for ci, yr in zip(h_cols, disp_hist):
                if yr in hr_years:
                    idx = hr_years.index(yr)
                    if idx < len(hr_vals) and hr_vals[idx] is not None:
                        c = ws.cell(row=r_idx, column=ci, value=hr_vals[idx])
                        c.number_format = fmt
        for ci, yr in zip(p_cols, proj_years):
            v = av(assum_key, yr)
            if v is not None and v != 0:
                c = ws.cell(row=r_idx, column=ci, value=v)
                c.number_format = fmt
                c.font = input_font
                c.fill = peach_fill

    r_val = len(assum_items) + 3
    ws.cell(row=r_val, column=1, value="VALUATION ASSUMPTIONS").font = sec_font
    for c in range(1, nc + 1):
        ws.cell(row=r_val, column=c).fill = grey_fill
    r_val += 1
    for label, key, fmt in [
        ("WACC %", "wacc_pct", RATIO), ("Terminal Growth %", "terminal_growth_pct", RATIO),
        ("Target PE (x)", "target_pe", RATIO), ("Target EV/EBITDA (x)", "target_ev_ebitda", RATIO),
        ("Dividend Payout %", "dividend_payout_pct", RATIO),
    ]:
        ws.cell(row=r_val, column=1, value=label)
        v = asmp.get(key)
        if v:
            c = ws.cell(row=r_val, column=2, value=v)
            c.number_format = fmt
            c.font = input_font
        r_val += 1
    r_val += 1
    ws.cell(row=r_val, column=1, value="RATIONALE").font = sec_font
    r_val += 1
    ws.merge_cells(start_row=r_val, start_column=1, end_row=r_val + 3, end_column=nc)
    ws.cell(row=r_val, column=1, value=asmp.get("rationale", "")).font = data_font
    ws.cell(row=r_val, column=1).alignment = Alignment(wrap_text=True, vertical="top")

    add_borders(ws, r_val + 3, nc)
    ws.freeze_panes = "B2"
    logger.info("  ✅ Assumptions")

    # ═══ P&L ═══
    ws, pln = mk("P&L")
    set_w(ws, [30] + [15] * (nc - 1))
    ws.cell(row=1, column=1, value="All figures in ₹ Cr").font = Font(name="Arial", italic=True, size=9)
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill

    A_GROWTH, A_RM, A_EMP, A_PF, A_OMFG, A_SA, A_OE, A_CI = 2, 3, 4, 5, 6, 7, 8, 9
    A_DEP, A_INT, A_OI, A_TAX, A_CAPEX, A_RECV, A_INV = 10, 11, 12, 13, 14, 15, 16

    r = 3
    ws.cell(row=r, column=1, value="Revenue").font = sec_font
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}17").number_format = INR
    for ci in p_cols:
        prev = get_column_letter(ci - 1)
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={prev}{r}*(1+'{asn}'!{cl}{A_GROWTH}/100)").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    REV = r
    r += 1

    expense_lines = [
        ("Raw Material Cost", 18, A_RM), ("Chg in Inventory", 19, A_CI),
        ("Employee Cost", 22, A_EMP), ("Power & Fuel", 20, A_PF),
        ("Other Mfg Exp", 21, A_OMFG), ("Selling & Admin", 23, A_SA),
        ("Other Expenses", 24, A_OE),
    ]
    exp_rows = []
    for label, ds_row, assum_row in expense_lines:
        ws.cell(row=r, column=1, value=label).font = data_font
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci in p_cols:
            cl = get_column_letter(ci)
            ws.cell(row=r, column=ci, value=f"={cl}{REV}*'{asn}'!{cl}{assum_row}/100").number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        exp_rows.append(r)
        r += 1

    ws.cell(row=r, column=1, value="Total Expenses").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=SUM({cl}{exp_rows[0]}:{cl}{exp_rows[-1]})").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    TOTEXP = r
    r += 1

    ws.cell(row=r, column=1, value="EBITDA").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{REV}-{cl}{TOTEXP}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EBITDA = r
    r += 1

    ws.cell(row=r, column=1, value="EBITDA Margin %")
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=IF({cl}{REV}=0,0,{cl}{EBITDA}/{cl}{REV})").number_format = PCT
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    r += 1

    ws.cell(row=r, column=1, value="Other Income")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}25").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{asn}'!{cl}{A_OI}").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    OI = r
    r += 1

    ws.cell(row=r, column=1, value="Depreciation")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}26").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{asn}'!{cl}{A_DEP}").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    DEP = r
    r += 1

    ws.cell(row=r, column=1, value="EBIT").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{EBITDA}+{cl}{OI}-{cl}{DEP}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EBIT = r
    r += 1

    ws.cell(row=r, column=1, value="Finance Cost")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}27").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{asn}'!{cl}{A_INT}").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    INT_R = r
    r += 1

    ws.cell(row=r, column=1, value="Profit Before Tax").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{EBIT}-{cl}{INT_R}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    PBT = r
    r += 1

    ws.cell(row=r, column=1, value="Tax")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}29").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=MAX(0,{cl}{PBT}*'{asn}'!{cl}{A_TAX}/100)").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    TAX = r
    r += 1

    ws.cell(row=r, column=1, value="Profit After Tax (PAT)").font = Font(
        name="Arial", bold=True, color=COLORS["navy"], size=12
    )
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{PBT}-{cl}{TAX}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    PAT = r
    r += 1

    ws.cell(row=r, column=1, value="PAT Margin %")
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=IF({cl}{REV}=0,0,{cl}{PAT}/{cl}{REV})").number_format = PCT
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    r += 1

    ws.cell(row=r, column=1, value="PAT Growth %")
    for ci in range(3, nc + 1):
        cl = get_column_letter(ci)
        prev = get_column_letter(ci - 1)
        ws.cell(row=r, column=ci, value=f"=IF({prev}{PAT}<=0,0,({cl}{PAT}/{prev}{PAT})-1)").number_format = PCT
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    r += 1

    ws.cell(row=r, column=1, value="EPS (₹)").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{PAT}/{shares_cr:.2f}").number_format = INR2
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EPS_R = r

    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ P&L (formula-based)")

    # ═══ BALANCE SHEET ═══
    ws, bsn = mk("Balance Sheet")
    set_w(ws, [30] + [15] * (nc - 1))
    ws.cell(row=1, column=1, value="All figures in ₹ Cr").font = Font(name="Arial", italic=True, size=9)
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill

    r = 3
    bs_src = [("Share Capital", 57, "share_capital"), ("Reserves", 58, "reserves")]
    bs_rr = {}
    for label, ds_row, pk in bs_src:
        ws.cell(row=r, column=1, value=label)
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        bs_rr[pk] = r
        r += 1

    ws.cell(row=r, column=1, value="Total Equity").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{bs_rr['share_capital']}+{cl}{bs_rr['reserves']}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EQ_R = r
    r += 1

    for label, ds_row, pk in [("Borrowings", 59, "borrowings"), ("Other Liabilities", 60, "other_liabilities")]:
        ws.cell(row=r, column=1, value=label)
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        bs_rr[pk] = r
        r += 1

    ws.cell(row=r, column=1, value="Total Liabilities & Equity").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(
            row=r, column=ci,
            value=f"={cl}{EQ_R}+{cl}{bs_rr['borrowings']}+{cl}{bs_rr['other_liabilities']}",
        ).number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    TL_R = r
    r += 2

    ws.cell(row=r, column=1, value="APPLICATION OF FUNDS").font = sec_font
    for c in range(1, nc + 1):
        ws.cell(row=r, column=c).fill = blue_fill
        ws.cell(row=r, column=c).font = sub_font
    r += 1

    asset_items = [("Net Block", 62, "net_block"), ("CWIP", 63, "cwip"), ("Investments", 64, "investments")]
    asset_rows = []
    for label, ds_row, pk in asset_items:
        ws.cell(row=r, column=1, value=label)
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        asset_rows.append(r)
        r += 1

    ws.cell(row=r, column=1, value="Trade Receivables")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}67").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{pln}'!{cl}{REV}*'{asn}'!{cl}{A_RECV}/365").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    RECV_R = r
    asset_rows.append(r)
    r += 1

    ws.cell(row=r, column=1, value="Inventory")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}68").number_format = INR
    RM_ROW = exp_rows[0]
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{pln}'!{cl}{RM_ROW}*'{asn}'!{cl}{A_INV}/365").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    INVTY_R = r
    asset_rows.append(r)
    r += 1

    ws.cell(row=r, column=1, value="Cash & Bank (Plug)")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}69").number_format = INR
    CASH_R = r
    r += 1

    ws.cell(row=r, column=1, value="Other Assets")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(
            row=r, column=ci,
            value=f"='Data Sheet'!{dc}65-'Data Sheet'!{dc}67-'Data Sheet'!{dc}68-'Data Sheet'!{dc}69",
        ).number_format = INR
    for ci, i in zip(p_cols, range(np_)):
        ws.cell(row=r, column=ci, value=pv("other_assets", i)).number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    OA_R = r
    r += 1

    for ci in p_cols:
        cl = get_column_letter(ci)
        asset_sum = "+".join(f"{cl}{ar}" for ar in asset_rows)
        ws.cell(row=CASH_R, column=ci, value=f"={cl}{TL_R}-{asset_sum}-{cl}{OA_R}").number_format = INR
        ws.cell(row=CASH_R, column=ci).fill = peach_fill

    ws.cell(row=r, column=1, value="Total Assets").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        all_a = "+".join(f"{cl}{ar}" for ar in asset_rows) + f"+{cl}{CASH_R}+{cl}{OA_R}"
        ws.cell(row=r, column=ci, value=f"={all_a}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    TA_R = r
    r += 1

    ws.cell(row=r, column=1, value="CHECK (TA - TL&E = 0)").font = red_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{TA_R}-{cl}{TL_R}").number_format = INR2

    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ Balance Sheet (formula-based)")

    # ═══ CASH FLOW ═══
    ws, _ = mk("Cash Flow")
    set_w(ws, [30] + [15] * (nc - 1))
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill
    r = 3
    cf_items = [("CFO", 82, "cfo"), ("CFI", 83, "cfi"), ("CFF", 84, "cff")]
    cf_rr = {}
    for label, ds_row, pk in cf_items:
        ws.cell(row=r, column=1, value=label).font = sec_font
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        cf_rr[pk] = r
        r += 1
    ws.cell(row=r, column=1, value="Net Cash Flow").font = sec_font
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}85").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(
            row=r, column=ci,
            value=f"={cl}{cf_rr['cfo']}+{cl}{cf_rr['cfi']}+{cl}{cf_rr['cff']}",
        ).number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    r += 2
    ws.cell(row=r, column=1, value="Free Cash Flow (FCF)").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{cf_rr['cfo']}+{cl}{cf_rr['cfi']}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill

    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ Cash Flow (formula-based)")

    # ═══ RATIOS ═══
    ws, _ = mk("Ratios")
    set_w(ws, [30] + [15] * (nc - 1))
    ws.cell(row=1, column=1, value="Key Financial Ratios").font = title_font
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill
    r = 3
    ratio_formulas = [
        ("EBITDA Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{EBITDA}/'{pln}'!{{cl}}{REV})", PCT),
        ("PAT Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{PAT}/'{pln}'!{{cl}}{REV})", PCT),
        ("ROE %", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{pln}'!{{cl}}{PAT}/'{bsn}'!{{cl}}{EQ_R})", PCT),
        ("ROCE %", f"=IF(('{bsn}'!{{cl}}{EQ_R}+'{bsn}'!{{cl}}{bs_rr['borrowings']})=0,0,'{pln}'!{{cl}}{EBIT}/('{bsn}'!{{cl}}{EQ_R}+'{bsn}'!{{cl}}{bs_rr['borrowings']}))", PCT),
        ("Debt/Equity (x)", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{bsn}'!{{cl}}{bs_rr['borrowings']}/'{bsn}'!{{cl}}{EQ_R})", RATIO),
        ("Receivable Days", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{bsn}'!{{cl}}{RECV_R}/'{pln}'!{{cl}}{REV}*365)", "0"),
        ("Inventory Days", f"=IF('{pln}'!{{cl}}{exp_rows[0]}=0,0,'{bsn}'!{{cl}}{INVTY_R}/'{pln}'!{{cl}}{exp_rows[0]}*365)", "0"),
        ("Asset Turnover (x)", f"=IF('{bsn}'!{{cl}}{TA_R}=0,0,'{pln}'!{{cl}}{REV}/'{bsn}'!{{cl}}{TA_R})", RATIO),
        ("EPS (₹)", f"='{pln}'!{{cl}}{PAT}/{shares_cr:.2f}", INR2),
        ("BVPS (₹)", f"='{bsn}'!{{cl}}{EQ_R}/{shares_cr:.2f}", INR2),
    ]
    for label, tmpl, fmt in ratio_formulas:
        ws.cell(row=r, column=1, value=label)
        for ci in range(2, nc + 1):
            cl = get_column_letter(ci)
            ws.cell(row=r, column=ci, value=tmpl.replace("{cl}", cl)).number_format = fmt
            if ci in p_cols:
                ws.cell(row=r, column=ci).fill = peach_fill
        r += 1
    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ Ratios (formula-based)")

    # ═══ VALUATION ═══
    ws, _ = mk("Valuation")
    set_w(ws, [35, 18, 18, 18, 18, 18])
    ws["A1"].value = "VALUATION SUMMARY"
    ws["A1"].font = title_font
    r = 3
    for label, v in [
        ("DCF Fair Value (₹)", val.get("dcf_fair_value")),
        ("PE Fair Value (₹)", val.get("pe_fair_value")),
        ("EV/EBITDA Fair Value (₹)", val.get("ev_ebitda_fair_value")),
        ("Blended Fair Value (₹)", val.get("blended_fair_value")),
        ("CMP (₹)", cmp), ("Rating", rating),
        ("Target Price (₹)", target), ("Upside %", upside),
    ]:
        ws.cell(row=r, column=1, value=label).font = sec_font
        ws.cell(row=r, column=2, value=v)
        for c in range(1, 3):
            ws.cell(row=r, column=c).border = thin_bdr
        r += 1
    r += 1

    ws.cell(row=r, column=1, value="PEER COMPARISON").font = sec_font
    r += 1
    for ci, h in enumerate(["Company", "MCap", "PE", "EV/EBITDA", "Margin", "ROE"], 1):
        ws.cell(row=r, column=ci, value=h).font = hdr_font
        ws.cell(row=r, column=ci).fill = navy_fill
    r += 1
    for p in peers:
        ws.cell(row=r, column=1, value=p.get("name", ""))
        ws.cell(row=r, column=2, value=p.get("mcap_cr")).number_format = INR
        ws.cell(row=r, column=3, value=p.get("pe")).number_format = RATIO
        ws.cell(row=r, column=4, value=p.get("ev_ebitda")).number_format = RATIO
        ws.cell(row=r, column=5, value=p.get("ebitda_margin_pct", ""))
        ws.cell(row=r, column=6, value=p.get("roe_pct")).number_format = RATIO
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = thin_bdr
        r += 1

    sens = val.get("sensitivity_pe", {})
    if sens and sens.get("grid"):
        r += 2
        ws.cell(row=r, column=1, value="SENSITIVITY TABLE").font = sec_font
        r += 1
        ws.cell(row=r, column=1, value=f"{sens.get('row_label','')} \\ {sens.get('col_label','')}")
        for ci, cv in enumerate(sens.get("col_values", []), 2):
            ws.cell(row=r, column=ci, value=cv).font = hdr_font
            ws.cell(row=r, column=ci).fill = navy_fill
        r += 1
        for rv, gv in zip(sens.get("row_values", []), sens.get("grid", [])):
            ws.cell(row=r, column=1, value=rv).font = sec_font
            for ci, v in enumerate(gv, 2):
                ws.cell(row=r, column=ci, value=v).number_format = INR
                ws.cell(row=r, column=ci).fill = peach_fill
            r += 1
    logger.info("  ✅ Valuation")

    # ═══ KPI DASHBOARD ═══
    ws, _ = mk("KPI Dashboard")
    set_w(ws, [35] + [15] * (nc - 1))
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill
    r = 3
    kpi_formulas = [
        ("Revenue (₹ Cr)", f"='{pln}'!{{cl}}{REV}", INR),
        ("EBITDA (₹ Cr)", f"='{pln}'!{{cl}}{EBITDA}", INR),
        ("PAT (₹ Cr)", f"='{pln}'!{{cl}}{PAT}", INR),
        ("EPS (₹)", f"='{pln}'!{{cl}}{EPS_R}", INR2),
        ("EBITDA Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{EBITDA}/'{pln}'!{{cl}}{REV})", PCT),
        ("ROE %", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{pln}'!{{cl}}{PAT}/'{bsn}'!{{cl}}{EQ_R})", PCT),
        ("D/E (x)", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{bsn}'!{{cl}}{bs_rr['borrowings']}/'{bsn}'!{{cl}}{EQ_R})", RATIO),
    ]
    for label, tmpl, fmt in kpi_formulas:
        ws.cell(row=r, column=1, value=label)
        for ci in range(2, nc + 1):
            cl = get_column_letter(ci)
            ws.cell(row=r, column=ci, value=tmpl.replace("{cl}", cl)).number_format = fmt
            if ci in p_cols:
                ws.cell(row=r, column=ci).fill = peach_fill
        r += 1
    add_borders(ws, r, nc)
    logger.info("  ✅ KPI Dashboard (formula-based)")

    # Verify originals intact + save
    final = wb.sheetnames
    assert all(s in final for s in orig_sheets), "❌ Original sheets missing!"
    wb.save(out_path)
    logger.info(f"✅ Saved → {out_path} ({os.path.getsize(out_path)//1024} KB)")


# ══════════════════════════════════════════════════════════════════
# P&L SERIES EMBEDDER
# ══════════════════════════════════════════════════════════════════
def _embed_pl_series(model_json: dict, screener_data: dict) -> None:
    """Compute Revenue/EBITDA/PAT series and store them in model_json in-place.

    The Excel model derives these via cell formulas that are never persisted to
    JSON. Without this step the PPTX generator has no absolute P&L numbers to
    put in charts, so slides show zeros or placeholder data.

    We use:
      - screener_data for the historical series (audited actuals)
      - model_json["assumptions"] growth/margin for the projection series
    Both halves are stored so the chart can show a full historical+forward view.
    """
    hist_years: list[str] = screener_data.get("fiscal_years") or []
    pl = screener_data.get("pl") or {}
    derived = screener_data.get("derived") or {}

    hist_revenue = pl.get("sales") or []
    hist_ebitda  = derived.get("ebitda") or []
    hist_pat     = pl.get("net_profit") or []

    # For recently-IPO'd / platform businesses (e.g. Swiggy, Zomato) Screener
    # often returns null for row 17 (sales). Fall back to the absolute values
    # Claude fetched via web_search, aligned to the same year list.
    web_pl = model_json.get("historical_pl_absolute") or {}
    web_years: list[str] = [str(y) for y in (web_pl.get("years") or [])]

    def _align_web(screener_vals: list, web_vals: list) -> list:
        """Return screener_vals, patching any None entries from web_vals aligned by year."""
        if not web_vals or len(web_vals) != len(web_years):
            return screener_vals
        web_by_yr = dict(zip(web_years, web_vals))
        out = []
        for yr, sv in zip(hist_years, screener_vals):
            out.append(sv if sv is not None else web_by_yr.get(yr))
        if len(screener_vals) < len(hist_years):
            out = screener_vals
        return out

    hist_revenue = _align_web(hist_revenue, web_pl.get("revenue") or [])
    hist_ebitda  = _align_web(hist_ebitda,  web_pl.get("ebitda")  or [])
    hist_pat     = _align_web(hist_pat,     web_pl.get("pat")     or [])

    # If Screener returned no years at all but web search has data, use web data directly
    if not hist_years and web_years:
        hist_years   = web_years
        hist_revenue = web_pl.get("revenue") or []
        hist_ebitda  = web_pl.get("ebitda")  or []
        hist_pat     = web_pl.get("pat")     or []

    model_json["historical_pl"] = {
        "years":   hist_years,
        "revenue": hist_revenue,
        "ebitda":  hist_ebitda,
        "pat":     hist_pat,
    }

    asmp = model_json.get("assumptions") or {}
    proj_years: list[str] = asmp.get("projection_years") or []
    revenue_growth: dict   = asmp.get("revenue_growth_pct") or {}

    hist_ratios = model_json.get("historical_ratios") or {}
    ebitda_margins = [v for v in (hist_ratios.get("ebitda_margin_pct") or []) if v is not None]
    pat_margins    = [v for v in (hist_ratios.get("pat_margin_pct")    or []) if v is not None]
    ebitda_m = ebitda_margins[-1] if ebitda_margins else 15.0
    pat_m    = pat_margins[-1]    if pat_margins    else 8.0

    # Base revenue = last non-null historical value
    base_rev = next((v for v in reversed(hist_revenue) if v is not None), None)

    proj_revenues: list = []
    proj_ebitdas:  list = []
    proj_pats:     list = []
    prev = base_rev
    for yr in proj_years:
        growth = (revenue_growth.get(yr) or 0) / 100
        rev = round(prev * (1 + growth), 1) if prev is not None else None
        proj_revenues.append(rev)
        proj_ebitdas.append(round(rev * ebitda_m / 100, 1) if rev is not None else None)
        proj_pats.append(   round(rev * pat_m    / 100, 1) if rev is not None else None)
        prev = rev

    model_json["projected_pl"] = {
        "years":   proj_years,
        "revenue": proj_revenues,
        "ebitda":  proj_ebitdas,
        "pat":     proj_pats,
    }


# ══════════════════════════════════════════════════════════════════
# PUBLIC ENTRY POINT
# ══════════════════════════════════════════════════════════════════
def generate_financial_model(
    nse_code: str,
    company_name: str,
    sector: str,
    output_dir: str,
    anthropic_api_key: Optional[str] = None,
    screener_username: Optional[str] = None,
    screener_password: Optional[str] = None,
) -> dict:
    """Generate a full financial model. Returns dict with file paths + metadata."""
    api_key = anthropic_api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY not provided (env var or argument)")

    sc_user = screener_username or os.environ.get("SCREENER_USERNAME")
    sc_pass = screener_password or os.environ.get("SCREENER_PASSWORD")
    if not sc_user or not sc_pass:
        raise ValueError("SCREENER_USERNAME and SCREENER_PASSWORD required (env vars or arguments)")

    os.makedirs(output_dir, exist_ok=True)
    tracker = CostTracker()
    t_start = time.time()

    nse_code = nse_code.upper()
    logger.info(f"🚀 PIPELINE START | {company_name} ({nse_code}) | Sector: {sector}")

    screener_path = download_screener_excel(nse_code, output_dir, sc_user, sc_pass)
    screener_data = extract_screener_data(screener_path)
    market_text = fetch_market_data(nse_code)

    model_json = call_claude_api(
        company_name, nse_code, sector, screener_data, market_text, api_key, tracker,
    )

    # Embed computed P&L series so the PPTX generator can build real charts.
    # The Excel model uses formulas that aren't stored in the JSON, so we
    # recompute Revenue/EBITDA/PAT here from the Screener historical base +
    # Claude's growth/margin assumptions.
    _embed_pl_series(model_json, screener_data)

    json_path = os.path.join(output_dir, f"{nse_code}_model.json")
    with open(json_path, "w") as f:
        json.dump(model_json, f, indent=2, default=str)

    out_path = os.path.join(output_dir, f"{nse_code}_financial_model.xlsx")
    shutil.copy2(screener_path, out_path)
    build_model(out_path, screener_data, model_json, out_path)

    elapsed = time.time() - t_start
    logger.info(tracker.summary())

    return {
        "file_path": out_path,
        "json_path": json_path,
        "model_json": model_json,
        "cost_summary": tracker.to_dict(),
        "elapsed_seconds": round(elapsed, 1),
        "nse_code": nse_code,
        "rating": model_json.get("rating"),
        "target_price": model_json.get("target_price"),
        "upside_pct": model_json.get("upside_pct"),
        "base_year": model_json.get("base_year"),
    }


# ══════════════════════════════════════════════════════════════════
# CLI / COLAB ENTRY
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate financial model for an NSE stock")
    parser.add_argument("--nse", required=True, help="NSE code (e.g., GRAVITA)")
    parser.add_argument("--name", required=True, help="Company name")
    parser.add_argument("--sector", default="General", help="Sector")
    parser.add_argument("--out", default="./output", help="Output directory")
    args = parser.parse_args()

    try:
        result = generate_financial_model(args.nse, args.name, args.sector, args.out)
        print(f"\n✅ DONE in {result['elapsed_seconds']}s")
        print(f"   File:        {result['file_path']}")
        print(f"   Rating:      {result['rating']}")
        print(f"   Target:      ₹{result['target_price']}")
        print(f"   Upside:      {result['upside_pct']}%")
        print(f"   Cost:        ${result['cost_summary']['cost_usd']} (₹{result['cost_summary']['cost_inr']})")
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        traceback.print_exc()
