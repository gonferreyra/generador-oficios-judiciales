#!/usr/bin/env python3
"""
Servidor web local para generar oficios
Ejecuta: python3 server.py
Abre: http://localhost:8000
"""

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse, StreamingResponse
from pathlib import Path
from jinja2 import Environment, FileSystemLoader, select_autoescape
import io

from generator import CAMPOS, generar, generar_todos, OUTPUT_DIR

app = FastAPI(title="Generador de Oficios")

# Jinja2 directo (evita bug en Starlette 1.6.0 + Jinja2Templates)
templates = Environment(
    loader=FileSystemLoader("web/templates"),
    autoescape=select_autoescape(['html', 'xml']),
    enable_async=False
)

Path("web/templates").mkdir(parents=True, exist_ok=True)


def render(template_name: str, **context) -> HTMLResponse:
    """Renderiza template y devuelve HTMLResponse"""
    template = templates.get_template(template_name)
    return HTMLResponse(template.render(**context))


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return render("index.html", request=request)


@app.get("/formulario", response_class=HTMLResponse)
async def formulario_unificado(request: Request):
    """Formulario unificado para generar los 3 oficios"""
    # Combinar todos los campos únicos de los 3 tipos
    todos_campos = {}
    todas_labels = {}
    for tipo_info in CAMPOS.values():
        for campo in tipo_info["required"]:
            if campo not in todos_campos:
                todos_campos[campo] = True
            if campo not in todas_labels:
                todas_labels[campo] = tipo_info["labels"].get(campo, campo)
    
    # Ordenar campos: primero los comunes, luego los específicos
    orden_prioridad = [
        "dia", "mes", "anio",
        "autos_caratulados", "juzgado",
        "matricula", "titular",
        "descripcion_inmueble",
        "nomenclatura_catastral", "ciudad",
        "que_se_subasta", "designado",
        "destinatario_constatacion",
        "numero_cuenta",
        "exento_pago", "tipo_exencion", "texto_exencion_otros"
    ]
    campos_ordenados = [c for c in orden_prioridad if c in todos_campos]
    
    return render("formulario.html", request=request,
                  campos=campos_ordenados, labels=todas_labels)


@app.post("/generar")
async def generar_oficios(
    request: Request,
    firma: str = Form("tribunal"),
):
    form_data = await request.form()
    datos = {k: v for k, v in form_data.items() if k != "firma"}

    try:
        zip_bytes = generar_todos(datos, firma)
        return StreamingResponse(
            io.BytesIO(zip_bytes),
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="oficios_{datos.get("matricula", "sin_matricula").replace(".", "_").replace("/", "_")}.zip"'}
        )
    except Exception as e:
        return HTMLResponse(f"<h1>Error: {e}</h1>", status_code=500)


if __name__ == "__main__":
    import uvicorn
    print("\n🚀 Servidor iniciado en http://localhost:8000")
    print("   Presiona Ctrl+C para detener\n")
    uvicorn.run(app, host="0.0.0.0", port=8000)