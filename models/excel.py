from pydantic import BaseModel, Field

class ExcelUpdateItem(BaseModel):
    row: int = Field(..., description="수정할 엑셀 행 번호", json_schema_extra={"example": 2})
    column: str = Field(..., description="수정할 엑셀 열 문자 (예: A, B, C)", json_schema_extra={"example": "C"})
    value: str = Field(..., description="셀에 입력할 새로운 값", json_schema_extra={"example": "C12345"})
    # sheet_name: str = Field(..., description="The name of the sheet to update")
    # cell: str = Field(..., description="The cell reference (e.g., A1) to update")
    # value: str = Field(..., description="The new value to set in the specified cell")