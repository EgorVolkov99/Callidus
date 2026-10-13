import uvicorn

if __name__ == "__main__":
    print("=== Запуск сервера Callidus ===")
    print("Откройте в браузере: http://127.0.0.1:8000")
    # Запускаем приложение из папки app
    uvicorn.run("app.main:app", host="127.0.0.1", port=8000, reload=True)