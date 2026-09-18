#!/usr/bin/env python3
"""
Entry point para PyInstaller - inicia servidor y abre navegador
"""
import sys
import os
import webbrowser
import threading
import time
from pathlib import Path

# Asegurar que podemos importar módulos locales cuando se ejecute como .exe
if getattr(sys, 'frozen', False):
    # Ejecutándose como .exe compilado
    BASE_DIR = Path(sys._MEIPASS)
else:
    # Ejecutándose como script normal
    BASE_DIR = Path(__file__).parent

os.chdir(BASE_DIR)
sys.path.insert(0, str(BASE_DIR))

from server import app
import uvicorn


def open_browser():
    """Abre el navegador después de un breve delay para que el servidor inicie"""
    time.sleep(1.5)
    webbrowser.open("http://localhost:8000")


def main():
    print("\n" + "=" * 50)
    print("  GENERADOR DE OFICIOS JUDICIALES")
    print("  " + "=" * 50)
    print("\n🚀 Iniciando servidor en http://localhost:8000")
    print("   Se abrirá automáticamente en tu navegador")
    print("   Presiona Ctrl+C para salir\n")

    # Abrir navegador en hilo separado
    threading.Thread(target=open_browser, daemon=True).start()

    # Iniciar servidor (bloquea hasta Ctrl+C)
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")


if __name__ == "__main__":
    main()