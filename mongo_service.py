import os
import re
import time
import logging
from datetime import datetime, timezone
from typing import Dict, Any, Optional
import pymongo
from pymongo import MongoClient

logger = logging.getLogger("mongo_service")

# Global cached client & error diagnostics
_mongo_client: Optional[MongoClient] = None
_db_name: Optional[str] = None
_last_mongo_error: Optional[str] = None
_last_mongo_hint: Optional[str] = None
_last_attempt_time: float = 0.0
_RETRY_INTERVAL: float = 15.0  # seconds cooldown between failed connection attempts


def get_mongo_uri() -> Optional[str]:
    """Retrieves MongoDB connection URI from environment variables."""
    return (
        os.getenv("MONGODB_URI")
        or os.getenv("MONGO_URI")
        or os.getenv("MONGODB_URL")
    )


def get_mongo_db():
    """Initializes and returns cached MongoDB database connection with detailed error diagnostics."""
    global _mongo_client, _db_name, _last_mongo_error, _last_mongo_hint, _last_attempt_time
    uri = get_mongo_uri()
    if not uri:
        _last_mongo_error = "MONGODB_URI environment variable is not set."
        _last_mongo_hint = "Add MONGODB_URI in Render dashboard -> Environment."
        return None

    # Detect unreplaced password placeholders
    if "<db_password>" in uri or "<password>" in uri:
        _last_mongo_error = "Found unreplaced '<db_password>' placeholder in connection string."
        _last_mongo_hint = "In Render Environment, replace '<db_password>' with your actual database user password."
        logger.error(_last_mongo_error)
        return None

    if _mongo_client is None:
        now = time.time()
        # If previous attempt failed, enforce cooldown to avoid repeated blocking
        if _last_mongo_error and (now - _last_attempt_time < _RETRY_INTERVAL):
            return None
        _last_attempt_time = now

        try:
            logger.info("Attempting MongoDB connection...")
            client = MongoClient(
                uri,
                serverSelectionTimeoutMS=3000,
                connectTimeoutMS=3000,
                socketTimeoutMS=5000,
                appname="university_student_portal"
            )
            # Test ping to verify authentication & network reachability
            client.admin.command("ping")
            
            db_name = os.getenv("MONGODB_DB_NAME") or "student_dossier_db"
            _db_name = db_name
            _mongo_client = client
            _last_mongo_error = None
            _last_mongo_hint = None
            logger.info(f"MongoDB successfully connected. Active database: {db_name}")
        except Exception as e:
            err_str = str(e)
            _last_mongo_error = err_str
            logger.warning(f"MongoDB connection failed: {err_str}")

            if "Authentication failed" in err_str or "bad auth" in err_str:
                _last_mongo_hint = "Authentication failed: Please verify your MongoDB database username and password in Render Environment."
            elif "ServerSelectionTimeoutError" in err_str or "timed out" in err_str:
                _last_mongo_hint = "Connection timed out. In MongoDB Atlas, go to Network Access -> Add IP Address -> Select '0.0.0.0/0' (Allow Access from Anywhere) and click Confirm."
            elif "ConfigurationError" in err_str:
                _last_mongo_hint = "Invalid connection URI format. Ensure your connection string starts with mongodb+srv:// or mongodb://."
            else:
                _last_mongo_hint = "Verify your MongoDB Atlas cluster status and Network Access IP whitelist (0.0.0.0/0)."

            _mongo_client = None
            return None

    db_name = _db_name or os.getenv("MONGODB_DB_NAME") or "student_dossier_db"
    return _mongo_client[db_name]


def save_student_dossier(
    result: Dict[str, Any],
    client_ip: Optional[str] = None,
    user_agent: Optional[str] = None
) -> bool:
    """
    Saves the extracted student dossier to MongoDB collections:
    1. 'students' - Upserts the complete student profile with all academic sections.
    2. 'scrape_history' - Records an immutable audit log of every extraction event.
    """
    uri = get_mongo_uri()
    if not uri:
        logger.debug("MONGODB_URI not configured in environment. Skipping database save.")
        return False

    try:
        db = get_mongo_db()
        if db is None:
            logger.warning("MongoDB is currently unreachable. Dossier not saved to database.")
            return False

        enrollment = result.get("enrollment_no", "").strip().upper()
        dob = result.get("dob", "").strip()
        summary = result.get("summary", {})
        sections = result.get("sections", {})
        total_fields = result.get("total_fields", 0)
        filename = result.get("filename", "")
        now = datetime.now(timezone.utc)

        # 1. Update / Upsert into 'students' collection
        student_doc = {
            "enrollment_no": enrollment,
            "dob": dob,
            "student_name": summary.get("student_name") or "Unknown",
            "course": summary.get("course") or "Unknown",
            "domicile": summary.get("domicile") or "Unknown",
            "total_fields": total_fields,
            "sections": sections,
            "latest_excel_filename": filename,
            "last_extracted_at": now,
            "last_client_ip": client_ip,
            "last_user_agent": user_agent,
        }

        db.students.update_one(
            {"enrollment_no": enrollment},
            {
                "$set": student_doc,
                "$setOnInsert": {"created_at": now},
                "$inc": {"scrape_count": 1}
            },
            upsert=True
        )

        # 2. Append event to 'scrape_history' collection for audit trail
        history_doc = {
            "enrollment_no": enrollment,
            "dob": dob,
            "student_name": summary.get("student_name") or "Unknown",
            "course": summary.get("course") or "Unknown",
            "total_fields": total_fields,
            "excel_filename": filename,
            "extracted_at": now,
            "client_ip": client_ip,
            "user_agent": user_agent,
            "status": "success"
        }
        db.scrape_history.insert_one(history_doc)

        logger.info(f"Student record [{enrollment}] successfully saved/updated in MongoDB.")
        return True

    except Exception as e:
        logger.error(f"Failed to save student dossier to MongoDB: {e}")
        return False


def find_cached_student_dossier(enrollment_no: str, dob: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """
    Checks if student record already exists in MongoDB 'students' collection.
    If found and DOB matches (or if sanitized digits match), returns the cached student profile.
    """
    uri = get_mongo_uri()
    if not uri:
        return None

    try:
        db = get_mongo_db()
        if db is None:
            return None

        enrollment = enrollment_no.strip().upper()
        doc = db.students.find_one({"enrollment_no": enrollment})
        if not doc:
            return None

        # Verify DOB if provided
        if dob:
            clean_search_dob = re.sub(r'[^0-9]', '', dob)
            clean_db_dob = re.sub(r'[^0-9]', '', doc.get("dob", ""))
            if clean_search_dob and clean_db_dob and clean_search_dob != clean_db_dob:
                logger.info(f"Student [{enrollment}] found in DB, but search DOB ({clean_search_dob}) != DB DOB ({clean_db_dob}). Skipping cache.")
                return None

        logger.info(f"Instant cache hit: Student [{enrollment}] already exists in MongoDB! Serving from cache.")
        return doc

    except Exception as e:
        logger.warning(f"Error checking student cache in MongoDB: {e}")
        return None


def check_mongo_health() -> Dict[str, Any]:
    """Returns the current MongoDB connection health with actionable diagnostic hints."""
    global _last_mongo_error, _last_mongo_hint
    uri = get_mongo_uri()
    if not uri:
        return {
            "configured": False,
            "status": "not_configured",
            "message": "MONGODB_URI environment variable is not set.",
            "hint": "Set MONGODB_URI in Render dashboard -> Environment tab."
        }
    try:
        db = get_mongo_db()
        if db is None:
            return {
                "configured": True,
                "status": "disconnected",
                "message": _last_mongo_error or "Could not connect to MongoDB instance.",
                "hint": _last_mongo_hint or "Check Network Access (0.0.0.0/0) and credentials in MongoDB Atlas."
            }
        count = db.students.estimated_document_count()
        return {
            "configured": True,
            "status": "connected",
            "database": db.name,
            "total_students_saved": count,
            "message": "MongoDB is actively connected and saving dossiers."
        }
    except Exception as e:
        return {
            "configured": True,
            "status": "error",
            "message": str(e),
            "hint": "Verify your MongoDB Atlas cluster status."
        }
