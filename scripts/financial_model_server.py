"""
Financial Model API Server
===========================
Deploy this on your VPS alongside financial_model_v5.py.
Exposes a single POST endpoint that n8n (or your frontend) calls.

Usage:
  pip install fastapi uvicorn anthropic yfinance beautifulsoup4 requests openpyxl pandas
  ANTHROPIC_API_KEY="sk-ant-..." python3 financial_model_server.py

The server runs on port 8500 by default.

Deployment setup (one-time, as root):
  sudo mkdir -p /var/lib/financial_models
  sudo chown <deploy-user> /var/lib/financial_models
  # Optional: prevent systemd-tmpfiles from clearing it
  echo "x /var/lib/financial_models" | sudo tee /etc/tmpfiles.d/financial_models.conf
"""

import os
import sys
import uuid
import time
import traceback
from fastapi import FastAPI, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
import requests
import uvicorn

# ── Make sure financial_model_v3 is importable from the same directory ──
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from financial_model_v5 import __version__ as FM_VERSION

app = FastAPI(title="Financial Model Generator", version="1.0")

# Output directory for generated models
OUTPUT_DIR = os.environ.get("MODEL_OUTPUT_DIR", "/var/lib/financial_models")
try:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
except PermissionError as e:
    raise RuntimeError(
        f"Cannot create {OUTPUT_DIR}. Run: sudo mkdir -p {OUTPUT_DIR} && sudo chown $(whoami) {OUTPUT_DIR}"
    ) from e
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
MODEL_STORAGE_BUCKET = os.environ.get("SUPABASE_FINANCIAL_MODEL_BUCKET", "research-reports-html")

# Track job status for async mode
jobs: dict[str, dict] = {}


# ========================
# Request / Response Models
# ========================

class GenerateRequest(BaseModel):
    nse_symbol: str
    company_name: str
    sector: str
    folder_id: str | None = None  # Google Drive company vault folder — model xlsx is mirrored here


class GenerateResponse(BaseModel):
    status: str  # "success" | "error"
    file_name: str | None = None
    file_path: str | None = None
    storage_path: str | None = None
    storage_url: str | None = None
    json_storage_path: str | None = None
    json_storage_url: str | None = None
    file_id: str | None = None
    file_url: str | None = None
    message: str | None = None
    duration_seconds: int | None = None


class JobStatus(BaseModel):
    job_id: str
    status: str  # "processing" | "completed" | "failed"
    state: str | None = None  # alias for status, for frontend compatibility
    file_name: str | None = None
    file_path: str | None = None
    storage_path: str | None = None
    storage_url: str | None = None
    json_storage_path: str | None = None
    json_storage_url: str | None = None
    file_id: str | None = None
    file_url: str | None = None
    message: str | None = None
    duration_seconds: int | None = None


class StorageMirrorResponse(BaseModel):
    status: str  # "success" | "error"
    file_name: str | None = None
    file_path: str | None = None
    storage_path: str | None = None
    storage_url: str | None = None
    json_storage_path: str | None = None
    json_storage_url: str | None = None
    message: str | None = None


# ========================
# Health Check
# ========================

@app.get("/health")
def health():
    return {
        "status": "ok",
        "model_version": FM_VERSION,
        "anthropic_key_set": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "screener_creds_set": bool(os.environ.get("SCREENER_USERNAME") and os.environ.get("SCREENER_PASSWORD")),
        "output_dir": OUTPUT_DIR,
    }


def _storage_public_url(path: str) -> str:
    return f"{SUPABASE_URL}/storage/v1/object/public/{MODEL_STORAGE_BUCKET}/{path}"


def upload_model_to_supabase(file_path: str, ticker: str) -> dict[str, str | None]:
    """Upload the v5 xlsx and (when present) the .json sidecar produced by
    financial_model_v5.generate_financial_model. The html-report pipeline
    prefers the JSON sidecar because the xlsx is formula-only and openpyxl
    can't read computed cell values without Excel having recalculated them.
    """
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise RuntimeError("SUPABASE_URL / SUPABASE_SERVICE_KEY not configured")

    ticker_u = ticker.upper()
    path = f"financial-models/{ticker_u}/{ticker_u}_model.xlsx"
    url = f"{SUPABASE_URL}/storage/v1/object/{MODEL_STORAGE_BUCKET}/{path}"
    headers = {
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "apikey": SUPABASE_SERVICE_KEY,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "x-upsert": "true",
        "Cache-Control": "no-cache",
    }

    with open(file_path, "rb") as f:
        resp = requests.put(url, data=f, headers=headers, timeout=120)

    if resp.status_code >= 400:
        raise RuntimeError(f"Supabase storage upload failed: {resp.status_code} {resp.text[:400]}")

    # Also upload the JSON sidecar if it exists alongside the xlsx
    json_local = os.path.splitext(file_path)[0] + ".json"
    json_path = None
    json_public_url = None
    if os.path.exists(json_local):
        json_path = f"financial-models/{ticker_u}/{ticker_u}_model.json"
        json_url = f"{SUPABASE_URL}/storage/v1/object/{MODEL_STORAGE_BUCKET}/{json_path}"
        json_headers = {
            **headers,
            "Content-Type": "application/json; charset=utf-8",
        }
        try:
            with open(json_local, "rb") as f:
                jr = requests.put(json_url, data=f, headers=json_headers, timeout=60)
            if jr.status_code >= 400:
                raise RuntimeError(f"JSON sidecar upload failed: {jr.status_code} {jr.text[:200]}")
            print(f"  Uploaded JSON sidecar -> {json_path}")
            json_public_url = _storage_public_url(json_path)
        except Exception as e:  # noqa: BLE001
            raise RuntimeError(f"JSON sidecar upload failed ({type(e).__name__}: {e})") from e
    else:
        print(f"  NOTE: no JSON sidecar at {json_local}; html pipeline will fall back to CSV/xlsx")

    return {
        "storage_path": path,
        "storage_url": _storage_public_url(path),
        "json_storage_path": json_path,
        "json_storage_url": json_public_url,
    }


def upload_model_to_gdrive(file_path: str, ticker: str, folder_id: str | None) -> dict[str, str | None]:
    """Mirror the generated xlsx into the company's Google Drive vault folder
    (same n8n /upload-document webhook used by the PPTX pipeline), into a
    "Financial Model" subfolder. Falls back to uploading straight into the
    company folder if the subfolder-aware call fails. Never raises — a Drive
    upload failure must not fail the overall generation, since the model is
    already safely in Supabase Storage by the time this runs.
    """
    if not folder_id:
        return {"file_id": None, "file_url": None}

    import base64

    n8n_base = os.environ.get("N8N_BASE_URL") or "https://n8n.tikonacapital.com/webhook"
    url = f"{n8n_base.rstrip('/')}/upload-document"
    file_name = f"{ticker.upper()}_model.xlsx"

    with open(file_path, "rb") as fh:
        file_base64 = base64.b64encode(fh.read()).decode("utf-8")

    def _post(subfolder_name: str | None) -> dict[str, str | None] | None:
        payload = {"folder_id": folder_id, "file_name": file_name, "file_base64": file_base64}
        if subfolder_name:
            payload["subfolder_name"] = subfolder_name
        try:
            res = requests.post(url, json=payload, timeout=60)
        except Exception as e:  # noqa: BLE001
            print(f"  Drive upload error (subfolder={subfolder_name}): {e}")
            return None
        if res.status_code != 200:
            print(f"  Drive upload failed (subfolder={subfolder_name}): {res.status_code} {res.text[:200]}")
            return None
        data = res.json()
        if data.get("status") != "success" and "file" not in data and "id" not in data:
            print(f"  Drive upload unexpected response (subfolder={subfolder_name}): {data}")
            return None
        file_info = data.get("file", data)
        if isinstance(file_info, list) and file_info:
            file_info = file_info[0]
        file_id = file_info.get("id")
        file_url = file_info.get("webViewLink") or (
            f"https://drive.google.com/file/d/{file_id}/view" if file_id else None
        )
        return {"file_id": file_id, "file_url": file_url}

    result = _post("Financial Model")
    if result is None:
        print("  Retrying Drive upload directly into company folder (no subfolder)...")
        result = _post(None)

    return result or {"file_id": None, "file_url": None}


# ========================
# Synchronous Generation (n8n calls this with long timeout)
# ========================

@app.post("/generate", response_model=GenerateResponse)
def generate_sync(req: GenerateRequest):
    """
    Synchronous endpoint — blocks until the model is generated.
    n8n should call this with a 15-minute timeout.
    Returns the file path so n8n can read and upload it.
    """
    ticker = req.nse_symbol.strip().upper()
    start = time.time()

    print(f"[generate] Received: nse_symbol='{req.nse_symbol}', ticker='{ticker}', company='{req.company_name}', sector='{req.sector}'")

    if not ticker:
        return GenerateResponse(
            status="error",
            message="nse_symbol is empty — check the webhook payload",
            duration_seconds=0,
        )

    try:
        # v5: formula-based, year-agnostic, cost-tracked
        from financial_model_v5 import generate_financial_model

        out_dir = os.path.join(OUTPUT_DIR, ticker)
        os.makedirs(out_dir, exist_ok=True)

        result = generate_financial_model(
            nse_code=ticker,
            company_name=req.company_name,
            sector=req.sector,
            output_dir=out_dir,
        )

        file_path = result["file_path"]
        json_src = result.get("json_path")

        # Copy xlsx + JSON sidecar into the canonical OUTPUT_DIR slot so
        # upload_model_to_supabase can find both side-by-side.
        final_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")
        final_json = os.path.join(OUTPUT_DIR, f"{ticker}_model.json")
        if file_path != final_path:
            import shutil
            shutil.copy2(file_path, final_path)
            file_path = final_path
        if json_src and os.path.exists(json_src) and json_src != final_json:
            import shutil
            shutil.copy2(json_src, final_json)

        storage_result = upload_model_to_supabase(file_path, ticker)
        gdrive_result = upload_model_to_gdrive(file_path, ticker, req.folder_id)

        return GenerateResponse(
            status="success",
            file_name=f"{ticker}_model.xlsx",
            file_path=file_path,
            storage_path=storage_result["storage_path"],
            storage_url=storage_result["storage_url"],
            json_storage_path=storage_result["json_storage_path"],
            json_storage_url=storage_result["json_storage_url"],
            file_id=gdrive_result["file_id"],
            file_url=gdrive_result["file_url"],
            duration_seconds=int(time.time() - start),
            message=(
                f"rating={result.get('rating')} target={result.get('target_price')} "
                f"upside={result.get('upside_pct')}% cost=${result['cost_summary']['cost_usd']}"
            ),
        )

    except Exception as e:
        traceback.print_exc()
        return GenerateResponse(
            status="error",
            message=str(e),
            duration_seconds=int(time.time() - start),
        )


# ========================
# Async Generation (if you want non-blocking)
# ========================

@app.post("/generate-async")
def generate_async(req: GenerateRequest, background_tasks: BackgroundTasks):
    """
    Async endpoint — returns immediately with a job_id.
    Poll /job/{job_id} to check status.
    """
    job_id = str(uuid.uuid4())[:8]
    jobs[job_id] = {"status": "processing", "started_at": time.time()}

    background_tasks.add_task(_run_generation, job_id, req)

    return {"job_id": job_id, "status": "processing", "message": "Generation started"}


@app.get("/job/{job_id}", response_model=JobStatus)
@app.get("/status/{job_id}", response_model=JobStatus)
def get_job_status(job_id: str):
    """Check the status of an async generation job."""
    if job_id not in jobs:
        return JSONResponse(status_code=404, content={"error": "Job not found"})

    job = jobs[job_id]
    return JobStatus(
        job_id=job_id,
        status=job["status"],
        state=job["status"],
        file_name=job.get("file_name"),
        file_path=job.get("file_path"),
        storage_path=job.get("storage_path"),
        storage_url=job.get("storage_url"),
        json_storage_path=job.get("json_storage_path"),
        json_storage_url=job.get("json_storage_url"),
        message=job.get("message"),
        duration_seconds=job.get("duration_seconds"),
    )


def _run_generation(job_id: str, req: GenerateRequest):
    """Background task for async generation."""
    ticker = req.nse_symbol.upper()
    start = time.time()

    try:
        from financial_model_v5 import generate_financial_model

        result = generate_financial_model(
            nse_code=ticker,
            company_name=req.company_name,
            sector=req.sector,
            output_dir=os.path.join(OUTPUT_DIR, ticker),
        )
        file_path = result["file_path"]
        json_src = result.get("json_path")

        final_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")
        final_json = os.path.join(OUTPUT_DIR, f"{ticker}_model.json")
        if file_path != final_path:
            import shutil
            shutil.copy2(file_path, final_path)
            file_path = final_path
        if json_src and os.path.exists(json_src) and json_src != final_json:
            import shutil
            shutil.copy2(json_src, final_json)

        storage_result = upload_model_to_supabase(file_path, ticker)
        gdrive_result = upload_model_to_gdrive(file_path, ticker, req.folder_id)

        jobs[job_id] = {
            "status": "completed",
            "file_name": f"{ticker}_model.xlsx",
            "file_path": file_path,
            "storage_path": storage_result["storage_path"],
            "storage_url": storage_result["storage_url"],
            "json_storage_path": storage_result["json_storage_path"],
            "json_storage_url": storage_result["json_storage_url"],
            "file_id": gdrive_result["file_id"],
            "file_url": gdrive_result["file_url"],
            "duration_seconds": int(time.time() - start),
            "rating": result.get("rating"),
            "target_price": result.get("target_price"),
            "upside_pct": result.get("upside_pct"),
            "cost_usd": result["cost_summary"]["cost_usd"],
        }

    except Exception as e:
        traceback.print_exc()
        jobs[job_id] = {
            "status": "failed",
            "message": str(e),
            "duration_seconds": int(time.time() - start),
        }


# ========================
# Download the generated file
# ========================

@app.get("/download/{ticker}")
def download_model(ticker: str):
    """Download the generated Excel file."""
    ticker = ticker.upper()
    file_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")

    if not os.path.exists(file_path):
        return JSONResponse(status_code=404, content={"error": f"No model found for {ticker}"})

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{ticker}_model.xlsx",
    )


@app.post("/storage/{ticker}", response_model=StorageMirrorResponse)
def mirror_model_to_storage(ticker: str):
    ticker = ticker.upper()
    file_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")

    if not os.path.exists(file_path):
        return StorageMirrorResponse(
            status="error",
            message=f"No model found for {ticker}",
        )

    try:
        storage_result = upload_model_to_supabase(file_path, ticker)
        return StorageMirrorResponse(
            status="success",
            file_name=f"{ticker}_model.xlsx",
            file_path=file_path,
            storage_path=storage_result["storage_path"],
            storage_url=storage_result["storage_url"],
            json_storage_path=storage_result["json_storage_path"],
            json_storage_url=storage_result["json_storage_url"],
        )
    except Exception as e:
        traceback.print_exc()
        return StorageMirrorResponse(
            status="error",
            file_name=f"{ticker}_model.xlsx",
            file_path=file_path,
            message=str(e),
        )


@app.post("/sync-screener/{ticker}")
def sync_screener(ticker: str):
    """
    Programmatically logs into Screener, downloads the audited financials Excel workbook,
    scrapes live stock indicators from yfinance / Screener HTML, and updates the equity_universe
    database table in Supabase.
    """
    ticker = ticker.strip().upper()
    print(f"[sync-screener] Syncing {ticker} from Screener.in...")
    try:
        username = os.environ.get("SCREENER_USERNAME")
        password = os.environ.get("SCREENER_PASSWORD")
        if not username or not password:
            return {"status": "error", "message": "SCREENER_USERNAME or SCREENER_PASSWORD is not set in server environment"}

        # 1. Download screener excel workbook
        out_dir = os.path.join(OUTPUT_DIR, ticker)
        os.makedirs(out_dir, exist_ok=True)
        
        from financial_model_v5 import download_screener_excel, extract_screener_data
        
        excel_path = download_screener_excel(ticker, out_dir, username, password)
        screener_data = extract_screener_data(excel_path)
        
        # 2. Scrape live market ratios using yfinance / screener HTML
        live_data = {}
        try:
            import yfinance as yf
            info = yf.Ticker(f"{ticker}.NS").info
            live_data["current_price"] = info.get("currentPrice")
            live_data["market_cap"] = info.get("marketCap")
            live_data["pe_ttm"] = info.get("trailingPE")
            live_data["roe"] = (info.get("returnOnEquity") or 0) * 100 if info.get("returnOnEquity") else None
            live_data["fifty_two_week_high"] = info.get("fiftyTwoWeekHigh")
            live_data["fifty_two_week_low"] = info.get("fiftyTwoWeekLow")
            live_data["volume"] = info.get("volume")
        except Exception as e:
            print(f"[sync-screener] yfinance lookup failed for {ticker}: {e}")

        # Supplement with details from Screener HTML page
        from bs4 import BeautifulSoup
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            for u in [
                f"https://www.screener.in/company/{ticker}/consolidated/",
                f"https://www.screener.in/company/{ticker}/",
            ]:
                r = requests.get(u, headers=headers, timeout=10)
                if r.status_code == 200:
                    soup = BeautifulSoup(r.text, "html.parser")
                    top_stats = soup.find("div", class_="company-ratios")
                    if top_stats:
                        for li in top_stats.find_all("li"):
                            name_el = li.find("span", class_="name")
                            val_el = li.find("span", class_="number")
                            if name_el and val_el:
                                name_text = name_el.get_text(strip=True).lower()
                                val_text = val_el.get_text(strip=True).replace(",", "").replace("%", "")
                                try:
                                    val_float = float(val_text)
                                    if "price" in name_text or "current price" in name_text:
                                        if not live_data.get("current_price"):
                                            live_data["current_price"] = val_float
                                    elif "market cap" in name_text:
                                        if not live_data.get("market_cap"):
                                            live_data["market_cap"] = val_float * 1e7  # Screener stores in Cr
                                    elif "stock p/e" in name_text or "p/e" in name_text:
                                        if not live_data.get("pe_ttm"):
                                            live_data["pe_ttm"] = val_float
                                    elif "roce" in name_text:
                                        live_data["roce"] = val_float
                                    elif "roe" in name_text:
                                        if not live_data.get("roe"):
                                            live_data["roe"] = val_float
                                    elif "promoter holding" in name_text:
                                        live_data["promoter_holding_pct"] = val_float
                                    elif "dividend yield" in name_text:
                                        live_data["dividend_yield"] = val_float
                                    elif "face value" in name_text:
                                        live_data["face_value"] = val_float
                                except ValueError:
                                    pass
                    break
        except Exception as e:
            print(f"[sync-screener] Screener HTML parse failed for {ticker}: {e}")

        # 3. Map parsed Excel data to equity_universe columns
        pl = screener_data.get("pl", {})
        bs = screener_data.get("bs", {})
        cf = screener_data.get("cf", {})
        derived = screener_data.get("derived", {})
        fiscal_years = screener_data.get("fiscal_years", [])
        
        def get_fy_val(series, fy):
            try:
                idx = fiscal_years.index(fy)
                v = series[idx]
                return float(v) if v is not None else None
            except (ValueError, IndexError, TypeError):
                return None

        def get_last_val(series):
            if not series:
                return None
            for v in reversed(series):
                if v is not None:
                    return float(v)
            return None

        # Get ISIN from master_company table using nse_symbol (ticker)
        isin_code = None
        if SUPABASE_URL and SUPABASE_SERVICE_KEY:
            headers_api = {
                "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                "apikey": SUPABASE_SERVICE_KEY,
                "Content-Type": "application/json",
            }
            try:
                mc_url = f"{SUPABASE_URL}/rest/v1/master_company?nse_symbol=eq.{ticker}&select=isin"
                mc_resp = requests.get(mc_url, headers=headers_api, timeout=10)
                if mc_resp.status_code == 200 and mc_resp.json():
                    isin_code = mc_resp.json()[0].get("isin")
            except Exception as e:
                print(f"[sync-screener] master_company lookup failed for {ticker}: {e}")
        
        # Fallback to avoid violating not-null constraint on equity_universe.isin_code
        if not isin_code:
            isin_code = f"INE{ticker:<9}".replace(" ", "0")[:12]

        payload = {
            "company_name": screener_data.get("company_name"),
            "nse_code": ticker,
            "isin_code": isin_code,
            "current_price": live_data.get("current_price") or screener_data.get("cmp"),
            "market_cap": live_data.get("market_cap") or screener_data.get("mcap"),
            "pe_ttm": live_data.get("pe_ttm"),
            "roe": live_data.get("roe"),
            "roce": live_data.get("roce"),
            "promoter_holding_pct": live_data.get("promoter_holding_pct"),
            "dividend_yield": live_data.get("dividend_yield"),
            "face_value": live_data.get("face_value") or screener_data.get("face_value"),
            
            # Historical Revenue
            "revenue_fy2023": get_fy_val(pl.get("sales"), "FY23"),
            "revenue_fy2024": get_fy_val(pl.get("sales"), "FY24"),
            "revenue_fy2025": get_fy_val(pl.get("sales"), "FY25"),
            "revenue_ttm": get_last_val(pl.get("sales")),
            
            # Historical PAT
            "pat_fy2023": get_fy_val(pl.get("net_profit"), "FY23"),
            "pat_fy2024": get_fy_val(pl.get("net_profit"), "FY24"),
            "pat_fy2025": get_fy_val(pl.get("net_profit"), "FY25"),
            "pat_ttm": get_last_val(pl.get("net_profit")),
            
            # Balance Sheet / Debt / Cash
            "debt": get_last_val(bs.get("borrowings")),
            "cash_equivalents": get_last_val(bs.get("cash")),
            "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S")
        }
        
        # Calculate derived margins
        sales_ttm = payload["revenue_ttm"]
        pat_ttm = payload["pat_ttm"]
        if sales_ttm and pat_ttm:
            payload["pat_margin_ttm"] = (pat_ttm / sales_ttm) * 100
        
        ebitda_ttm = get_last_val(derived.get("ebitda"))
        if sales_ttm and ebitda_ttm:
            payload["ebitda_margin_ttm"] = (ebitda_ttm / sales_ttm) * 100

        # 4. Write to Supabase using REST API
        if SUPABASE_URL and SUPABASE_SERVICE_KEY:
            headers = {
                "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                "apikey": SUPABASE_SERVICE_KEY,
                "Content-Type": "application/json",
            }
            # Query existing company_id
            query_url = f"{SUPABASE_URL}/rest/v1/equity_universe?nse_code=eq.{ticker}&select=company_id"
            q_resp = requests.get(query_url, headers=headers, timeout=10)
            if q_resp.status_code == 200 and q_resp.json():
                company_id = q_resp.json()[0]["company_id"]
                # Update (PATCH)
                up_url = f"{SUPABASE_URL}/rest/v1/equity_universe?company_id=eq.{company_id}"
                up_resp = requests.patch(up_url, json=payload, headers=headers, timeout=10)
                if up_resp.status_code >= 400:
                    raise Exception(f"Failed to update equity_universe: {up_resp.status_code} {up_resp.text}")
                payload["company_id"] = company_id
            else:
                # Insert (POST)
                ins_url = f"{SUPABASE_URL}/rest/v1/equity_universe"
                ins_resp = requests.post(ins_url, json=payload, headers=headers, timeout=10)
                if ins_resp.status_code >= 400:
                    raise Exception(f"Failed to insert into equity_universe: {ins_resp.status_code} {ins_resp.text}")

        return {"status": "success", "data": payload}

    except Exception as e:
        traceback.print_exc()
        return {"status": "error", "message": str(e)}


# ========================
# Run Server
# ========================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8500))
    print(f"\n🚀 Financial Model Server starting on port {port}")
    print(f"📁 Output directory: {OUTPUT_DIR}")
    print(f"🔑 Anthropic key: {'✓ Set' if os.environ.get('ANTHROPIC_API_KEY') else '✗ MISSING'}\n")

    uvicorn.run(app, host="0.0.0.0", port=port)
