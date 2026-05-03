# fastapi dev main.py
from fastapi import FastAPI
from api.excel_highlighter import router as excel_router
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    servers=[
        {"url": "https://excel-loader.onrender.com", "description": "Render Service server"},
        {"url": "http://127.0.0.1:8000", "description": "Local server"}
    ]
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # 모든 도메인 허용 (테스트용)
    allow_methods=["*"],
    allow_headers=["*"],
)

# prefix를 "/excel"로 주면, excel.py의 모든 경로 앞에 /excel이 자동으로 붙습니다.
# tags를 지정하면 Swagger UI에서 예쁘게 그룹화됩니다.
app.include_router(excel_router, prefix="/excel", tags=["Excel"])

@app.get("/")
def root():
    return {"message": "Welcome to Excel API Server"}
