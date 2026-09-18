# Generador de Oficios Judiciales

Sistema minimalista para generar oficios judiciales (Constatación, Rentas, Municipalidad) a partir de plantillas Word (.docx) + datos variables.

---

## 🚀 Inicio Rápido (para usuarios finales - Windows)

**Si solo querés usar la app (sin instalar Python):**

1. Descargá el archivo `GeneradorOficios.exe` desde [Releases](../../releases)
2. Hacé **doble clic** en el `.exe`
3. Se abre automáticamente el navegador en `http://localhost:8000`
4. Completá el formulario y descargá los 3 oficios en un ZIP

> **Nota:** Si Windows muestra "Protegiste tu PC", clickeá "Más información" → "Ejecutar de todas formas". El archivo no está firmado digitalmente aún.

---

## 🛠️ Instalación para Desarrolladores (Python)

### Requisitos
- Python 3.10+
- Git

### Pasos
```bash
# 1. Clonar repo
git clone git@github.com:gonferreyra/generador-oficios-judiciales.git
cd generador-oficios-judiciales/moray

# 2. Instalar dependencias
pip install -r requirements.txt

# 3a. Usar versión web (desarrollo)
python server.py
# Abre http://localhost:8000

# 3b. Usar versión CLI
python generator.py constatacion --data data/ejemplo_constatacion.json
```

> **Linux/Ubuntu:** Usá `python3` y `pip3`. Si `pip` no está en PATH: `export PATH="$HOME/.local/bin:$PATH"`

---

## Estructura

```
poC-oficios/
├── templates/                 # Plantillas .docx base
│   ├── constatacion/
│   │   └── tribunal.docx
│   ├── rentas/
│   │   └── tribunal.docx
│   └── municipalidad/
│       └── tribunal.docx
├── data/                      # Datos de ejemplo (JSON)
├── output/                    # Oficios generados (se crea solo)
├── web/                       # Interfaz web
│   └── templates/             # HTML Jinja2
├── generator.py               # Motor CLI + librería
├── server.py                  # Servidor web local (FastAPI)
├── create_template.py         # Script para crear plantillas
├── requirements.txt
└── README.md
```

## Instalación (una sola vez)

```bash
cd /home/gon/github/POC-Emi
python3 -m pip install --break-system-packages -r requirements.txt
```

> **Nota:** En Linux/Ubuntu usar `python3` (no `python`). Si `pip` no está en PATH: `export PATH="/home/gon/.local/bin:$PATH"`

## Uso CLI

```bash
# 1. Constatación
python3 generator.py constatacion --data data/ejemplo_constatacion.json

# 2. Rentas
python3 generator.py rentas --data data/ejemplo_rentas.json

# 3. Municipalidad
python3 generator.py municipalidad --data data/ejemplo_municipalidad.json

# Interactivo (pide campos por consola)
python3 generator.py constatacion

# Especificar nombre de salida
python3 generator.py rentas --data data/ejemplo_rentas.json --output mi_rentas.docx
```

Los archivos `.docx` generados caen en `output/`

## Uso Web (interfaz visual)

```bash
python3 server.py
# Abre http://localhost:8000 en el navegador
# 1. Seleccionás tipo de oficio
# 2. Completás formulario
# 3. Descarga .docx automáticamente
```

## Tipos de oficio incluidos

| Tipo | Firma | Campos clave |
|------|-------|--------------|
| **constatacion** | Solo tribunal | fecha, autos, juzgado, matrícula, titular, qué se subasta, designado, descripción, nomenclatura catastral |
| **rentas** | Solo tribunal | fecha, autos, juzgado, matrícula, titular, descripción, número de cuenta |
| **municipalidad** | Solo tribunal | fecha, autos, juzgado, matrícula, titular, descripción, nomenclatura catastral, ciudad |

## Agregar nuevo tipo de oficio

1. Crear carpeta en `templates/nuevo_tipo/`
2. Poner `tribunal.docx` (y `martillero.docx` si aplica) con `{{placeholders}}`
3. Agregar entrada en `CAMPOS` dict en `generator.py` (línea ~20):

```python
"nuevo_tipo": {
    "required": ["campo1", "campo2", ...],
    "labels": {"campo1": "Etiqueta amigable", ...}
}
```

4. Listo: aparece en CLI y web automáticamente

## Crear plantillas propias

Opción A: Editar el `.docx` directamente en Word (usar `{{campo}}` para variables)

Opción B: Modificar `create_template.py` y correr:
```bash
python3 create_template.py
```

## Compartir el sistema

Copiar toda la carpeta `POC-Emi/` a otra PC → `python3 -m pip install --break-system-packages -r requirements.txt` → funciona igual. Sin Docker, sin cloud, sin costos.

## Ejecutable standalone (.exe para Windows)

Para usuarios no técnicos (martilleros, etc.) que no tienen Python instalado:

### Construir en Windows (recomendado)

```cmd
# En Windows con Python instalado:
pip install -r requirements.txt pyinstaller
python build_exe.py
```

Genera `dist/GeneradorOficios.exe` (~50-80 MB). El usuario solo hace **doble clic** → se abre el navegador en `http://localhost:8000`.

### Para desarrollador (Linux/macOS - solo test local)

```bash
python3 build_exe.py
# Genera dist/GeneradorOficios (binario Linux, para testear que compila bien)
```

### Notas importantes

- **Compilar en Windows** para obtener `.exe` nativo (cross-compile no funciona bien)
- Incluye templates, data y web automáticamente via `--add-data`
- `--noconsole` oculta la terminal (solo abre el navegador)
- Para evitar alertas de antivirus: firmar con certificado EV (opcional)

## Para producción futura

- Docker: `docker build -t oficios . && docker run -p 8000:8000 oficios`
- Agregar autenticación, logs, base de datos de plantillas, etc.