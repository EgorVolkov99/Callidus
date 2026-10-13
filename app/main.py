from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from app.api.routers import docs
from fastapi.responses import HTMLResponse

app = FastAPI(title="Callidus", version="2.0")

# Настройка CORS для работы с фронтендом
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # В продакшене укажите точный домен фронтенда
    allow_methods=["*"],
    allow_headers=["*"],
)

# Монтируем папки хранилища, чтобы файлы были доступны для скачивания по прямым ссылкам
app.mount("/downloads", StaticFiles(directory="storage/outputs"), name="downloads")
app.mount("/archives", StaticFiles(directory="storage/archives"), name="archives")

# Подключаем роутеры
app.include_router(docs.router, prefix="/api/docs", tags=["Documents"])

@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    with open("frontend/index.html", "r", encoding="utf-8") as f:
        return f.read()