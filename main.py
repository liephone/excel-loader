# fastapi dev main.py
from fastapi import FastAPI
from api.excel_highlighter import router as excel_router

app = FastAPI()

# prefix를 "/excel"로 주면, excel.py의 모든 경로 앞에 /excel이 자동으로 붙습니다.
# tags를 지정하면 Swagger UI에서 예쁘게 그룹화됩니다.
app.include_router(excel_router, prefix="/excel", tags=["Excel"])

@app.get("/")
def root():
    return {"message": "Welcome to Excel API Server"}
