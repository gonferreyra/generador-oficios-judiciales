#!/usr/bin/env python3
"""
Script para compilar el .exe con PyInstaller
Uso: python build_exe.py
"""
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent

def build():
    import platform
    sep = ";" if platform.system() == "Windows" else ":"
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=GeneradorOficios",
        "--onefile",
        "--noconsole",
        "--icon=NONE",
        f"--add-data=templates{sep}templates",
        f"--add-data=data{sep}data",
        f"--add-data=web{sep}web",
        "--hidden-import=jinja2",
        "--hidden-import=jinja2.ext",
        "--hidden-import=docx",
        "--hidden-import=docx.shared",
        "--hidden-import=docx.oxml.ns",
        "--hidden-import=fastapi",
        "--hidden-import=uvicorn",
        "--hidden-import=pydantic",
        "--hidden-import=python_multipart",
        "--hidden-import=starlette",
        "main.py"
    ]

    print("🔨 Compilando .exe con PyInstaller...")
    print(f"   Comando: {' '.join(cmd)}")

    result = subprocess.run(cmd, cwd=BASE_DIR)

    if result.returncode == 0:
        import platform
        exe_name = "GeneradorOficios.exe" if platform.system() == "Windows" else "GeneradorOficios"
        exe_path = BASE_DIR / "dist" / exe_name
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✅ Éxito! Binario generado en: {exe_path}")
            print(f"   Tamaño: {size_mb:.1f} MB")
        else:
            print("\n⚠️  Build completado pero no se encontró el binario")
    else:
        print("\n❌ Error en la compilación")
        sys.exit(1)


if __name__ == "__main__":
    build()