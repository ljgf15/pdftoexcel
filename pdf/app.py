# app.py
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import json
import os
import uuid
from pathlib import Path

from extractor import (
    parse_pdf,
    write_excel,
    ALLOWED_COLUMNS
)

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
def index():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/api/preview")
async def api_preview(files: list[UploadFile] = File(...), columns: str = Form(...)):
    try:
        cols = json.loads(columns)
        all_rows = []
        errors = []
        for f in files:
            try:
                suffix = Path(f.filename).suffix
                tmp_path = f"tmp_{uuid.uuid4()}{suffix}"
                with open(tmp_path, "wb") as fp:
                    fp.write(await f.read())
                rows = parse_pdf(tmp_path)
                all_rows.extend(rows)
                os.remove(tmp_path)
            except Exception as e:
                errors.append({"file": f.filename, "error": str(e)})
        return {
            "rows": all_rows,
            "errors": errors,
            "summary": {
                "file_count": len(files),
                "row_count": len(all_rows),
                "error_count": len(errors)
            }
        }
    except Exception as e:
        return {"message": str(e)}, 500

@app.post("/api/export")
async def api_export(files: list[UploadFile] = File(...), columns: str = Form(...)):
    cols = json.loads(columns)
    all_rows = []
    errors = []
    for f in files:
        try:
            suffix = Path(f.filename).suffix
            tmp_path = f"tmp_{uuid.uuid4()}{suffix}"
            with open(tmp_path, "wb") as fp:
                fp.write(await f.read())
            rows = parse_pdf(tmp_path)
            all_rows.extend(rows)
            os.remove(tmp_path)
        except Exception as e:
            errors.append({"file": f.filename, "error": str(e)})
    out_path = os.path.join(OUTPUT_DIR, f"result_{uuid.uuid4()}.xlsx")
    write_excel(all_rows, cols, out_path, errors)
    return FileResponse(out_path, filename="PDF_extract_result.xlsx")
