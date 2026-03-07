from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse

from formatter import format_document

APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
INDEX_FILE = STATIC_DIR / "index.html"
OUTPUT_DIR = APP_DIR / "output"
ACCESS_PASSWORD = os.getenv("APP_PASSWORD", "").strip()

app = FastAPI(title="文件格式整理系统", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    if not INDEX_FILE.exists():
        raise HTTPException(status_code=404, detail="index.html not found")
    return INDEX_FILE.read_text(encoding="utf-8")


def _validate_password(x_app_password: str | None) -> None:
    if not ACCESS_PASSWORD:
        return
    if (x_app_password or "").strip() != ACCESS_PASSWORD:
        raise HTTPException(status_code=401, detail="访问口令不正确。")


@app.get("/api/download/{filename}")
def download_file(filename: str, x_app_password: str | None = Header(default=None)) -> FileResponse:
    _validate_password(x_app_password)
    target = OUTPUT_DIR / Path(filename).name
    if not target.exists():
        raise HTTPException(status_code=404, detail="Formatted file not found.")
    return FileResponse(
        path=target,
        filename=target.name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/api/format")
async def format_word(file: UploadFile = File(...), x_app_password: str | None = Header(default=None)) -> dict:
    _validate_password(x_app_password)
    suffix = Path(file.filename or "").suffix.lower()
    if suffix != ".docx":
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_name = f"{Path(file.filename or 'document').stem}-formatted-{uuid4().hex[:8]}.docx"
    output_path = OUTPUT_DIR / output_name

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        temp_input = Path(tmp.name)

    try:
        with temp_input.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        summary = format_document(temp_input, output_path)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Failed to format document: {exc}") from exc
    finally:
        temp_input.unlink(missing_ok=True)
        await file.close()

    return {
        "filename": output_name,
        "download_url": f"/api/download/{output_name}",
        "report": summary.to_dict(),
    }
