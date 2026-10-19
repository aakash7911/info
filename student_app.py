import os
import re
import logging
import asyncio
import resource
import gc
from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import Optional

from datetime import datetime, timezone

from student_scraper import scrape_student_data, generate_styled_excel, DOSSIERS_DIR, cleanup_browser_processes
from mongo_service import save_student_dossier, check_mongo_health, find_cached_student_dossier, get_mongo_db

logger = logging.getLogger("student_app")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# Global concurrency lock to guarantee strictly ONE Chromium instance at a time (protects 512MB RAM limit)
SCRAPE_LOCK = asyncio.Lock()

def get_current_memory_mb() -> float:
    """Returns current resident set size (RSS) in MB."""
    try:
        usage = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024
        if hasattr(os, "uname") and os.uname().sysname == "Darwin":
            usage = usage / 1024
        return round(usage, 2)
    except Exception:
        return 0.0

def build_cached_response(cached_doc: dict, enrollment: str, dob: str, request: Request) -> JSONResponse:
    """Builds and returns cached dossier response without touching browser engine."""
    filename = cached_doc.get("latest_excel_filename") or f"Student_Dossier_{enrollment}.xlsx"
    file_path = os.path.join(DOSSIERS_DIR, filename)
    sections = cached_doc.get("sections") or {}

    summary = {
        "enrollment_no": enrollment,
        "dob": dob,
        "student_name": cached_doc.get("student_name") or "Unknown",
        "course": cached_doc.get("course") or "Unknown",
        "institute": "University",
        "contact_mobile": "N/A",
        "email": "N/A",
        "domicile": cached_doc.get("domicile") or "Unknown"
    }

    # If file missing from ephemeral disk, regenerate Excel directly from cached sections
    if not os.path.exists(file_path) and sections:
        all_sections_data = []
        for sec_name, field_list in sections.items():
            for item in field_list:
                all_sections_data.append({
                    "Section": sec_name,
                    "Field Name": item.get("field", ""),
                    "Value": item.get("value", "")
                })
        if all_sections_data:
            try:
                file_path = generate_styled_excel(all_sections_data, enrollment)
                filename = os.path.basename(file_path)
            except Exception as ex:
                logger.warning(f"Could not regenerate cached Excel file: {ex}")

    # Update access timestamp in MongoDB
    client_ip = request.client.host if request.client else "unknown"
    user_agent = request.headers.get("user-agent", "unknown")
    try:
        db = get_mongo_db()
        if db is not None:
            db.students.update_one(
                {"enrollment_no": enrollment},
                {"$set": {"last_accessed_at": datetime.now(timezone.utc), "last_client_ip": client_ip}}
            )
            db.scrape_history.insert_one({
                "enrollment_no": enrollment,
                "dob": dob,
                "student_name": summary.get("student_name"),
                "course": summary.get("course"),
                "total_fields": cached_doc.get("total_fields", 0),
                "excel_filename": filename,
                "extracted_at": datetime.now(timezone.utc),
                "client_ip": client_ip,
                "user_agent": user_agent,
                "status": "cache_hit"
            })
    except Exception:
        pass

    return JSONResponse(status_code=200, content={
        "success": True,
        "enrollment_no": enrollment,
        "dob": dob,
        "excel_path": file_path,
        "filename": filename,
        "total_fields": cached_doc.get("total_fields", 0),
        "summary": summary,
        "sections": sections,
        "cached_from_db": True,
        "saved_to_mongo": True,
        "error": None
    })

app = FastAPI(
    title="University — Student Dossier Gateway",
    description="Automated Student Verification & Academic Dossier Gateway Engine",
    version="1.0.0"
)

# CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "student_scanner.html")

class ScrapeRequest(BaseModel):
    enrollment_no: str = Field(..., description="Student Enrollment Number / UG Number")
    dob: str = Field(..., description="Student Date of Birth (DD/MM/YYYY or DD-MM-YYYY)")
    headless: bool = Field(default=True, description="Run invisible headless browser (default True)")

async def verify_mongo_in_background():
    """Asynchronously verifies MongoDB in background so the port opens immediately."""
    await asyncio.sleep(1.0)
    try:
        print("\n" + "=" * 65, flush=True)
        print("🚀 [STARTUP] Verifying MongoDB Connection Status...", flush=True)
        health = await asyncio.to_thread(check_mongo_health)
        status = health.get("status")
        if status == "connected":
            print(f"✅ [MONGODB SUCCESS] Connected to cluster successfully!", flush=True)
            print(f"   Database: {health.get('database')}", flush=True)
            print(f"   Total Saved Students: {health.get('total_students_saved')}", flush=True)
        elif status == "not_configured":
            print("⚠️ [MONGODB WARNING] MONGODB_URI is not set in Render Environment.", flush=True)
            print("   Set MONGODB_URI to automatically save student profiles.", flush=True)
        else:
            print(f"❌ [MONGODB ERROR] Connection failed: {health.get('message')}", flush=True)
            if health.get("hint"):
                print(f"👉 [ACTION REQUIRED] {health.get('hint')}", flush=True)
        print("=" * 65 + "\n", flush=True)
    except Exception as exc:
        logger.warning(f"Background MongoDB health check error: {exc}")

async def keep_alive_background_worker():
    """
    Periodic keep-alive task that pings the application's external URL
    every 10 minutes (600s) to prevent Render free-tier from idling/sleeping (15-min timeout).
    """
    # Wait 45 seconds after startup before first ping
    await asyncio.sleep(45)
    
    while True:
        target_url = (
            os.getenv("KEEP_ALIVE_URL")
            or os.getenv("RENDER_EXTERNAL_URL")
            or (f"https://{os.getenv('RENDER_EXTERNAL_HOSTNAME')}" if os.getenv("RENDER_EXTERNAL_HOSTNAME") else None)
        )
        
        if target_url:
            ping_url = target_url.rstrip("/") + "/api/ping"
            try:
                def _do_ping():
                    import urllib.request
                    req = urllib.request.Request(
                        ping_url,
                        headers={"User-Agent": "Render-KeepAlive-Worker/1.0"}
                    )
                    with urllib.request.urlopen(req, timeout=12) as response:
                        return response.getcode()

                status_code = await asyncio.to_thread(_do_ping)
                logger.info(f"⚡ [KEEP-ALIVE] Heartbeat pinged {ping_url} - Status: {status_code} (Prevented sleep mode)")
            except Exception as e:
                logger.warning(f"⚠️ [KEEP-ALIVE] Ping to {ping_url} failed: {e}")
        else:
            logger.info("ℹ️ [KEEP-ALIVE] Standby: Set KEEP_ALIVE_URL in Render Environment to auto-ping.")

        # Sleep for 10 minutes (600s) - Render free tier sleeps after 15 minutes (900s)
        await asyncio.sleep(600)

@app.on_event("startup")
async def startup_event():
    """Fast, non-blocking startup: spawns background DB check and keep-alive ping worker."""
    logger.info("Application starting up. Scheduling background services...")
    asyncio.create_task(verify_mongo_in_background())
    asyncio.create_task(keep_alive_background_worker())
    logger.info("Startup complete. Ready to bind port.")

@app.get("/api/ping")
@app.get("/ping")
async def ping():
    """Ultra-lightweight ping endpoint for keep-alive monitoring and preventing sleep mode."""
    return {
        "status": "awake",
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "service": "student-dossier-gateway",
        "rss_mb": get_current_memory_mb()
    }

@app.get("/", response_class=HTMLResponse)
async def serve_scanner_ui():
    """Serves the bespoke ID card scanner and dossier web application."""
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=404, detail="Scanner frontend template not found.")
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    return HTMLResponse(
        content=content,
        headers={
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0"
        }
    )

@app.get("/health")
@app.get("/api/health")
async def health_check():
    """Health check endpoint for Render / monitoring with live memory stats."""
    db_health = await asyncio.to_thread(check_mongo_health)
    current_mem_mb = get_current_memory_mb()
    return {
        "status": "healthy",
        "service": "student-dossier-gateway",
        "browser_engine": "headless-chromium",
        "memory": {
            "current_rss_mb": current_mem_mb,
            "render_limit_mb": 512,
            "status": "safe" if current_mem_mb < 400 else "warning"
        },
        "mongodb": db_health
    }

@app.get("/api/db/status")
async def db_status():
    """Returns detailed status of MongoDB connection for admin/mongosh verification."""
    return await asyncio.to_thread(check_mongo_health)

@app.post("/api/scrape")
async def trigger_student_scrape(payload: ScrapeRequest, request: Request):
    """
    Retrieves or generates student dossier.
    First checks MongoDB for cached dossier (bypassing portal login completely).
    If not cached, launches headless browser automation, extracts data, and persists to MongoDB.
    Strictly serialized with SCRAPE_LOCK to prevent memory from ever exceeding 512MB RAM.
    """
    enrollment = payload.enrollment_no.strip().upper()
    dob = payload.dob.strip()

    if not enrollment or not dob:
        raise HTTPException(status_code=400, detail="Both enrollment_no and dob are required.")

    clean_dob = re.sub(r'[^0-9]', '', dob)
    logger.info(f"Incoming dossier request for Enrollment: {enrollment} | DOB: {dob} (Sanitized digits: {clean_dob})")

    # -------------------------------------------------------------
    # 1. CHECK MONGODB CACHE (FAST PATH - ZERO BROWSER RAM)
    # -------------------------------------------------------------
    cached_doc = find_cached_student_dossier(enrollment_no=enrollment, dob=dob)
    if cached_doc:
        logger.info(f"Instant cache hit in MongoDB for [{enrollment}]! Skipping browser login entirely.")
        return build_cached_response(cached_doc, enrollment, dob, request)

    # -------------------------------------------------------------
    # 2. RUN BROWSER EXTRACTION IF NOT YET CACHED
    # -------------------------------------------------------------
    logger.info(f"[MEM BEFORE SCRAPE] Current RSS: {get_current_memory_mb():.2f} MB / 512 MB")

    if SCRAPE_LOCK.locked():
        logger.info(f"Engine busy. Request for [{enrollment}] queued. Waiting for turn...")

    result = None
    try:
        # Serialize browser execution so strictly ONE headless Chrome instance runs at a time
        async with SCRAPE_LOCK:
            # Double-check cache inside lock in case a previous queued request just saved it
            cached_doc_inner = find_cached_student_dossier(enrollment_no=enrollment, dob=dob)
            if cached_doc_inner:
                logger.info(f"Double-check cache hit in lock for [{enrollment}]! Skipping browser.")
                return build_cached_response(cached_doc_inner, enrollment, dob, request)

            logger.info(f"Acquired scrape lock for {enrollment}. Launching Chrome in thread pool...")
            try:
                result = await asyncio.wait_for(
                    asyncio.to_thread(
                        scrape_student_data,
                        enrollment_no=enrollment,
                        dob=dob,
                        headless=payload.headless
                    ),
                    timeout=75.0
                )
            except asyncio.TimeoutError:
                logger.error(f"Scrape timed out after 75s for {enrollment}.")
                return JSONResponse(
                    status_code=504,
                    content={
                        "success": False,
                        "enrollment_no": enrollment,
                        "dob": dob,
                        "excel_path": None,
                        "filename": None,
                        "total_fields": 0,
                        "summary": {},
                        "sections": {},
                        "error": "Portal verification timed out. Please try again."
                    }
                )
    finally:
        # Aggressively reclaim memory after every browser run
        cleanup_browser_processes()
        gc.collect()
        logger.info(f"[MEM AFTER SCRAPE & GC] Current RSS: {get_current_memory_mb():.2f} MB / 512 MB")

    if not result or not result.get("success"):
        error_msg = (result.get("error") if result else None) or "Failed to extract student profile."
        logger.error(f"Scraping failed: {error_msg}")
        return JSONResponse(status_code=422, content=result or {"success": False, "error": error_msg})

    # Persist student data & audit log into MongoDB
    client_ip = request.client.host if request.client else "unknown"
    user_agent = request.headers.get("user-agent", "unknown")
    saved_to_db = save_student_dossier(result=result, client_ip=client_ip, user_agent=user_agent)
    result["saved_to_mongo"] = saved_to_db

    return JSONResponse(status_code=200, content=result)

@app.get("/api/download/{filename}")
async def download_dossier_excel(filename: str):
    """Securely downloads the generated Excel dossier."""
    # Sanitize to prevent path traversal
    safe_filename = os.path.basename(filename)
    file_path = os.path.join(DOSSIERS_DIR, safe_filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Requested dossier file not found or expired.")

    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=safe_filename
    )

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("student_app:app", host="0.0.0.0", port=port, reload=True)
