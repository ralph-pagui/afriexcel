"""
Script de démarrage du serveur AfriExcel FastAPI
Port : 8000
"""

import uvicorn

if __name__ == "__main__":
    print("=" * 60)
    print("🚀 Démarrage du Copilote IA Business Plan OHADA (Backend Python)")
    print("📍 URL Locale : http://127.0.0.1:8000")
    print("📖 Documentation API : http://127.0.0.1:8000/docs")
    print("=" * 60)
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
