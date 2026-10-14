import os
import re
import logging
from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import Optional

from student_scraper import scrape_student_data, DOSSIERS_DIR

logger = logging.getLogger("student_app")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

app = FastAPI(
    title="Parul University — Student Dossier Gateway",
    description="Automated Headless Extraction and Verification Engine for Parul University UMS Portal",
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

@app.get("/", response_class=HTMLResponse)
async def serve_scanner_ui():
    """Serves the bespoke ID card scanner and dossier web application."""
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=404, detail="Scanner frontend template not found.")
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    return HTMLResponse(content=content)

@app.get("/health")
@app.get("/api/health")
async def health_check():
    """Health check endpoint for Render / monitoring."""
    return {
        "status": "healthy",
        "service": "student-dossier-gateway",
        "browser_engine": "headless-chromium"
    }

@app.post("/api/scrape")
async def trigger_student_scrape(payload: ScrapeRequest):
    """
    Triggers the headless browser automation to log into UMS,
    harvest all student tabs, and compile an Excel workbook.
    """
    enrollment = payload.enrollment_no.strip()
    dob = payload.dob.strip()

    if not enrollment or not dob:
        raise HTTPException(status_code=400, detail="Both enrollment_no and dob are required.")

    logger.info(f"Incoming scrape request for Enrollment: {enrollment}")
    
    # Run scraper
    result = scrape_student_data(enrollment_no=enrollment, dob=dob, headless=payload.headless)

    if not result.get("success"):
        error_msg = result.get("error") or "Failed to extract student profile from UMS."
        logger.error(f"Scraping failed: {error_msg}")
        return JSONResponse(status_code=422, content=result)

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
