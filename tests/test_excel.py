# python -m pytest
# python -m pytest tests/test_excel.py
import pytest
from fastapi.testclient import TestClient
from main import app  # FastAPI 앱 객체 임포트
import io
import json

client = TestClient(app)

def test_highlight_excel_success():
    # 1. 가짜 엑셀 파일 생성 (테스트용)
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Original Value"
    
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    # 2. API 호출 데이터 준비
    json_payload = [
        {"row": 1, "column": "A", "value": "Updated Value"}
    ]
    
    # 3. API 호출 (multipart/form-data)
    response = client.post(
        "/excel/highlight",
        files={"file": ("test.xlsx", excel_file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"json_data": json.dumps(json_payload)}
    )

    # 4. 검증
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"