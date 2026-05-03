from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import StreamingResponse
import json
from typing import List
from urllib.parse import quote
from pydantic import ValidationError
from models.excel import ExcelUpdateItem
from services.excel_service import apply_excel_highlights


router = APIRouter()

@router.post("/highlight")
async def highlight_excel(
    file: UploadFile = File(...),
    json_data: str = Form(...)
):
    # 1. 파일 및 데이터 읽기
    content = await file.read()

    try:
        raw_data = json.loads(json_data)
        update_data: List[ExcelUpdateItem] = [ExcelUpdateItem(**item) for item in raw_data]
    except (json.JSONDecodeError, ValidationError) as e:
        return {"error": "잘못된 데이터 형식입니다.", "details": str(e)}

    output = apply_excel_highlights(content, update_data)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={quote(file.filename)}_highlighted.xlsx"}
    )