from pathlib import Path
import os
import shutil
import uuid

from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

try:
    from .processing import process_excel
except ImportError:
    from processing import process_excel

app = FastAPI()

cors_origins = [
    origin.strip()
    for origin in os.getenv("CORS_ORIGINS", "*").split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=False,
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs" 

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")


def cleanup_files(*paths):
    for path in paths:
        try:
            Path(path).unlink(missing_ok=True)
        except OSError:
            pass


@app.get("/")
def home():
    return FileResponse(BASE_DIR / "index.html")


@app.get("/style.css", include_in_schema=False)
def stylesheet():
    return FileResponse(BASE_DIR / "style.css", media_type="text/css")


@app.get("/config.js", include_in_schema=False)
def config_script():
    return FileResponse(BASE_DIR / "config.js", media_type="application/javascript")


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    allowed_extensions = [".xlsx", ".xls"]

    if file.filename is None:
        return {
            "success": False,
            "message": "No file name found."
        }

    original_filename = file.filename
    file_ext = Path(original_filename).suffix.lower()

    if file_ext not in allowed_extensions:
        return {
            "success": False,
            "message": "Only Excel files are allowed."
        }

    # Create unique filenames to avoid overwriting files
    unique_id = uuid.uuid4().hex

    uploaded_filename = f"uploaded_{unique_id}{file_ext}"
    output_filename = f"modified_{unique_id}.xlsx"

    uploaded_path = UPLOAD_DIR / uploaded_filename
    output_path = OUTPUT_DIR / output_filename

    try:
        # Save uploaded file
        with open(uploaded_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Process Excel file
        process_excel(uploaded_path, output_path)
    except Exception:
        cleanup_files(uploaded_path, output_path)
        raise
    finally:
        await file.close()

    # Return modified file as automatic download
    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="sắp thất nghiệp.xlsx",
        background=BackgroundTask(cleanup_files, uploaded_path, output_path),
    )
