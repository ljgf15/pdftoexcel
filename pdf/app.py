# -*- coding: utf-8 -*-
import json
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from extractor import (
    ALLOWED_COLUMNS,
    parse_pdf,
    write_excel,
    get_row_value,
)


app = FastAPI(title="PDF采购订单提取到Excel")


def parse_columns(columns_raw: str):
    try:
        columns = json.loads(columns_raw)
    except Exception:
        return []

    if not isinstance(columns, list):
        return []

    result = []
    for col in columns:
        name = str(col).strip()
        if name in ALLOWED_COLUMNS and name not in result:
            result.append(name)
    return result


async def save_uploads(files: List[UploadFile]):
    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_excel_"))
    pdf_paths = []

    for file in files:
        filename = Path(file.filename or "upload.pdf").name
        if not filename.lower().endswith(".pdf"):
            continue

        target = temp_dir / filename
        content = await file.read()
        target.write_bytes(content)
        pdf_paths.append(target)

    return temp_dir, pdf_paths


def filter_rows(rows, columns):
    filtered = []
    for row in rows:
        filtered.append({col: get_row_value(row, col) for col in columns})
    return filtered


@app.post("/api/preview")
async def preview(files: List[UploadFile] = File(...), columns: str = Form(...)):
    selected_columns = parse_columns(columns)
    if not selected_columns:
        return JSONResponse(status_code=400, content={"message": "请至少选择一个允许的导出字段。"})

    temp_dir, pdf_paths = await save_uploads(files)
    all_rows = []
    errors = []

    try:
        if not pdf_paths:
            return JSONResponse(status_code=400, content={"message": "请上传 PDF 文件。"})

        for pdf_path in pdf_paths:
            try:
                rows = parse_pdf(pdf_path)
                if not rows:
                    errors.append({"file": pdf_path.name, "error": "未提取到数据，请确认 PDF 是文字型且格式已覆盖。"})
                else:
                    all_rows.extend(rows)
            except Exception as exc:
                errors.append({"file": pdf_path.name, "error": str(exc)})

        preview_rows = filter_rows(all_rows, selected_columns)
        return {
            "rows": preview_rows,
            "errors": errors,
            "summary": {
                "file_count": len(pdf_paths),
                "row_count": len(preview_rows),
                "error_count": len(errors),
            },
        }
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@app.post("/api/export")
async def export_excel(files: List[UploadFile] = File(...), columns: str = Form(...)):
    selected_columns = parse_columns(columns)
    if not selected_columns:
        return JSONResponse(status_code=400, content={"message": "请至少选择一个允许的导出字段。"})

    temp_dir, pdf_paths = await save_uploads(files)
    all_rows = []
    errors = []

    try:
        if not pdf_paths:
            return JSONResponse(status_code=400, content={"message": "请上传 PDF 文件。"})

        for pdf_path in pdf_paths:
            try:
                rows = parse_pdf(pdf_path)
                if not rows:
                    errors.append({"file": pdf_path.name, "error": "未提取到数据"})
                else:
                    all_rows.extend(rows)
            except Exception as exc:
                errors.append({"file": pdf_path.name, "error": str(exc)})

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = temp_dir / f"PDF_extract_result_{timestamp}.xlsx"
        write_excel(all_rows, selected_columns, output_path, errors)

        return FileResponse(
            output_path,
            filename=output_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    finally:
        # FileResponse 会在响应后读取文件。这里不立即删除 temp_dir，Render 临时目录会自动清理。
        pass


app.mount("/", StaticFiles(directory="static", html=True), name="static")
