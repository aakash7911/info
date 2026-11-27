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

from student_scraper import scrape_student_data, generate_styled_excel, DOSSIERS_DIR
from mongo_service import (
    save_student_dossier,
    check_mongo_health,
    find_cached_student_dossier,
    get_mongo_db,
    purge_blocked_students,
    BLOCKED_ENROLLMENT_CLEAN
)

logger = logging.getLogger("student_app")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

def normalize_dob(dob_str: str) -> str:
    if not dob_str:
        return ""
    s = dob_str.strip().lower()
    month_map = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }
    for m_name, m_num in month_map.items():
        if m_name in s:
            s = re.sub(rf'\b{m_name}[a-z]*\b', m_num, s)
            break
    digits = re.sub(r'[^0-9]', '', s)
    if len(digits) == 8:
        return digits
    return dob_str.strip()

SCRAPE_LOCK = asyncio.Lock()

def get_current_memory_mb() -> float:
    try:
        usage = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024
        if hasattr(os, "uname") and os.uname().sysname == "Darwin":
            usage = usage / 1024
        return round(usage, 2)
    except Exception:
        return 0.0

def build_cached_response(cached_doc: dict, enrollment: str, dob: str, request: Request) -> JSONResponse:
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
    title="Student Dossier Gateway",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "student_scanner.html")

class ScrapeRequest(BaseModel):
    enrollment_no: str
    dob: str
    headless: bool = True

async def verify_mongo_in_background():
    await asyncio.sleep(1.0)
    try:
        await asyncio.to_thread(purge_blocked_students)
        health = await asyncio.to_thread(check_mongo_health)
        status = health.get("status")
        if status == "connected":
            logger.info(f"MongoDB connected: {health.get('database')}, Saved: {health.get('total_students_saved')}")
        elif status == "not_configured":
            logger.warning("MONGODB_URI is not set.")
        else:
            logger.warning(f"MongoDB connection issue: {health.get('message')}")
    except Exception as exc:
        logger.warning(f"Background MongoDB health check error: {exc}")

async def keep_alive_background_worker():
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
                        headers={"User-Agent": "KeepAlive-Worker/1.0"}
                    )
                    with urllib.request.urlopen(req, timeout=12) as response:
                        return response.getcode()

                status_code = await asyncio.to_thread(_do_ping)
                logger.info(f"Pinged {ping_url} - Status: {status_code}")
            except Exception as e:
                logger.warning(f"Ping to {ping_url} failed: {e}")
        await asyncio.sleep(600)

@app.on_event("startup")
async def startup_event():
    asyncio.create_task(verify_mongo_in_background())
    asyncio.create_task(keep_alive_background_worker())

@app.get("/api/ping")
@app.get("/ping")
async def ping():
    return {
        "status": "awake",
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "rss_mb": get_current_memory_mb()
    }

@app.get("/", response_class=HTMLResponse)
async def serve_scanner_ui():
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=404, detail="Scanner template not found.")
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
    db_health = await asyncio.to_thread(check_mongo_health)
    current_mem_mb = get_current_memory_mb()
    return {
        "status": "healthy",
        "service": "student-dossier-gateway",
        "memory": {
            "current_rss_mb": current_mem_mb,
            "render_limit_mb": 512,
            "status": "safe" if current_mem_mb < 400 else "warning"
        },
        "mongodb": db_health
    }

@app.get("/api/db/status")
async def db_status():
    return await asyncio.to_thread(check_mongo_health)

@app.post("/api/scrape")
async def trigger_student_scrape(payload: ScrapeRequest, request: Request):
    enrollment = payload.enrollment_no.strip().upper()
    dob = payload.dob.strip()

    if not enrollment or not dob:
        raise HTTPException(status_code=400, detail="Both enrollment_no and dob are required.")

    clean_enroll = re.sub(r'[^0-9A-Za-z]', '', enrollment).upper()
    if clean_enroll == BLOCKED_ENROLLMENT_CLEAN:
        logger.warning(f"Blocked access for enrollment: {enrollment}")
        await asyncio.to_thread(purge_blocked_students)
        return JSONResponse(
            status_code=403,
            content={
                "success": False,
                "error": f"FUCK YOU. Access permanently denied: Enrollment {BLOCKED_ENROLLMENT_CLEAN} is banned on this system."
            }
        )

    clean_dob = normalize_dob(dob)

    cached_doc = find_cached_student_dossier(enrollment_no=enrollment, dob=clean_dob)
    if cached_doc:
        return build_cached_response(cached_doc, enrollment, clean_dob, request)

    result = None
    try:
        async with SCRAPE_LOCK:
            cached_doc_inner = find_cached_student_dossier(enrollment_no=enrollment, dob=clean_dob)
            if cached_doc_inner:
                return build_cached_response(cached_doc_inner, enrollment, clean_dob, request)

            try:
                result = await asyncio.wait_for(
                    asyncio.to_thread(
                        scrape_student_data,
                        enrollment_no=enrollment,
                        dob=clean_dob,
                        headless=payload.headless
                    ),
                    timeout=180.0
                )
            except asyncio.TimeoutError:
                logger.error(f"Scrape timed out after 180s for {enrollment}")
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
        gc.collect()

    if not result or not result.get("success"):
        error_msg = (result.get("error") if result else None) or "Failed to extract student profile."
        return JSONResponse(status_code=422, content=result or {"success": False, "error": error_msg})

    client_ip = request.client.host if request.client else "unknown"
    user_agent = request.headers.get("user-agent", "unknown")
    saved_to_db = save_student_dossier(result=result, client_ip=client_ip, user_agent=user_agent)
    result["saved_to_mongo"] = saved_to_db

    return JSONResponse(status_code=200, content=result)

@app.get("/api/download/{filename}")
async def download_dossier_excel(filename: str):
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
