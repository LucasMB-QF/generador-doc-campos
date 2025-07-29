from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import Response, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from docx import Document
import re
from io import BytesIO
import os
from pathlib import Path
import logging

# Configuración básica de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuración de rutas
current_dir = Path(__file__).resolve().parent
templates_dir = current_dir / "templates"

app = FastAPI()

# Montar directorio estático (necesario para Vercel)
app.mount("/static", StaticFiles(directory="templates"), name="static")

# Configuración de templates
templates = Jinja2Templates(directory=str(templates_dir))

# Middleware CORS (opcional pero recomendado)
from fastapi.middleware.cors import CORSMiddleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Regex para detectar campos tipo {{Hoja!Celda}} o {{Hoja!Rango}}
campo_regex = re.compile(r"\{\{\s*([^\{\}]+?)\s*\}\}")

def obtener_valor(wb, hoja_nombre, celda):
    try:
        hoja = wb[hoja_nombre]
        celda_valor = hoja[celda].value
        return str(celda_valor) if celda_valor is not None else ""
    except Exception as e:
        logger.error(f"Error en celda {hoja_nombre}!{celda}: {str(e)}")
        return ""

def obtener_valores_rango(wb, hoja_nombre, rango):
    try:
        hoja = wb[hoja_nombre]
        celdas = hoja[rango]
        fila = celdas[0]
        return [str(c.value) if c.value is not None else "" for c in fila]
    except Exception as e:
        logger.error(f"Error en rango {hoja_nombre}!{rango}: {str(e)}")
        return []

def reemplazar_campos(texto, wb):
    def reemplazo(match):
        campo = match.group(1)
        if '!' in campo:
            hoja, celda_o_rango = campo.split('!', 1)
            hoja = hoja.strip()
            celda_o_rango = celda_o_rango.strip()
            if ':' in celda_o_rango:
                valores = obtener_valores_rango(wb, hoja, celda_o_rango)
                return ', '.join(valores)
            else:
                return obtener_valor(wb, hoja, celda_o_rango)
        return ""
    return campo_regex.sub(reemplazo, texto)

def procesar_documento(doc, wb):
    for p in doc.paragraphs:
        if campo_regex.search(p.text):
            nuevo_texto = reemplazar_campos(p.text, wb)
            p.text = nuevo_texto

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if campo_regex.search(celda.text):
                    nuevo_texto = reemplazar_campos(celda.text, wb)
                    celda.text = nuevo_texto

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/procesar")
async def procesar(
    archivo_excel: UploadFile = File(...),
    archivo_word: UploadFile = File(...)
):
    try:
        logger.info("Iniciando procesamiento de archivos...")
        
        # Validación básica de tipos de archivo
        if not archivo_excel.filename.endswith(('.xlsx', '.xlsm')):
            raise HTTPException(400, "El archivo Excel debe ser .xlsx o .xlsm")
        if not archivo_word.filename.endswith('.docx'):
            raise HTTPException(400, "El archivo Word debe ser .docx")

        # Leer contenido en memoria
        excel_content = await archivo_excel.read()
        word_content = await archivo_word.read()

        # Procesamiento en memoria
        with BytesIO(excel_content) as excel_stream:
            wb = load_workbook(filename=excel_stream, data_only=True)
            
            with BytesIO(word_content) as word_stream:
                doc = Document(word_stream)
                procesar_documento(doc, wb)
                
                output_stream = BytesIO()
                doc.save(output_stream)
                output_stream.seek(0)
                
                logger.info("Procesamiento completado con éxito")
                return Response(
                    content=output_stream.getvalue(),
                    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    headers={
                        "Content-Disposition": "attachment; filename=informe_generado.docx",
                        "Access-Control-Expose-Headers": "Content-Disposition"
                    }
                )
                
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error en el procesamiento: {str(e)}", exc_info=True)
        raise HTTPException(500, f"Error interno del servidor: {str(e)}")

# Manejo de errores personalizado
@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    return templates.TemplateResponse(
        "error.html",
        {
            "request": request,
            "status_code": exc.status_code,
            "detail": exc.detail
        },
        status_code=exc.status_code
    )
