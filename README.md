# University — Student Dossier & Verification Portal

Automated student profile extraction and verification portal using ID Card OCR and Headless Browser Automation.

## Features
- **Live Camera & Photo Upload Scanner**: Real-time OCR powered by Tesseract.js to detect Enrollment Number (UG No.) and Date of Birth.
- **Headless Browser Gateway**: Invisible Selenium Chrome automated login and data harvesting from UMS Portal.
- **Academic Dossier Compilation**: Formats all sections (Personal, Admission, Contact, Parents, Qualification) into an official Segoe UI styled Excel sheet (`.xlsx`).
- **Render Ready**: Optimized for deployment on Render (Docker or Python 3).

## Local Development
```bash
pip install -r requirements.txt
python student_app.py
```
Open [http://localhost:8000](http://localhost:8000)

## Render Deployment
- **Runtime**: Docker (Recommended) or Python 3
- **Start Command**: `uvicorn student_app:app --host 0.0.0.0 --port $PORT`
