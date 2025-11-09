from fastapi import FastAPI, UploadFile, File, Form, Request, HTTPException
from fastapi.responses import JSONResponse, Response, HTMLResponse
from fastapi.templating import Jinja2Templates
# from fastapi.staticfiles import StaticFiles  # Comentado porque no montamos estáticos
from docx import Document
from io import BytesIO
import re
from pathlib import Path
import json

app = FastAPI()

# Config directorios
current_dir = Path(__file__).parent.resolve()
templates = Jinja2Templates(directory=str(current_dir / "templates"))
# app.mount("/static", StaticFiles(directory=str(current_dir / "static")), name="static")  # Comentado para evitar error

# Regex para campos manuales {{campo}} (sin '!')
campo_regex = re.compile(r"\{\{\s*([^\{\}!]+?)\s*\}\}")

def extraer_campos_de_parrafos(parrafos):
    campos = set()
    for p in parrafos:
        for match in campo_regex.finditer(p.text):
            campos.add(match.group(1).strip())
    return campos

def extraer_campos_de_tablas(tablas):
    campos = set()
    for table in tablas:
        for row in table.rows:
            for cell in row.cells:
                campos |= extraer_campos_de_parrafos(cell.paragraphs)
    return campos

def extraer_todos_los_campos(doc: Document):
    campos = set()
    campos |= extraer_campos_de_parrafos(doc.paragraphs)
    campos |= extraer_campos_de_tablas(doc.tables)

    for section in doc.sections:
        campos |= extraer_campos_de_parrafos(section.header.paragraphs)
        campos |= extraer_campos_de_tablas(section.header.tables)
        campos |= extraer_campos_de_parrafos(section.footer.paragraphs)
        campos |= extraer_campos_de_tablas(section.footer.tables)

        if section.different_first_page_header_footer:
            campos |= extraer_campos_de_parrafos(section.first_page_header.paragraphs)
            campos |= extraer_campos_de_tablas(section.first_page_header.tables)
            campos |= extraer_campos_de_parrafos(section.first_page_footer.paragraphs)
            campos |= extraer_campos_de_tablas(section.first_page_footer.tables)
    return sorted(campos)

def reemplazar_texto_en_parrafo(parrafo, reemplazos):
    texto_completo = "".join(run.text for run in parrafo.runs)
    def reemplazo_match(match):
        campo = match.group(1).strip()
        return str(reemplazos.get(campo, match.group(0)))
    nuevo_texto = campo_regex.sub(reemplazo_match, texto_completo)

    if nuevo_texto != texto_completo:
        if parrafo.runs:
            parrafo.runs[0].text = nuevo_texto
            for run in parrafo.runs[1:]:
                run.text = ""

def reemplazar_campos(doc: Document, reemplazos: dict):
    for p in doc.paragraphs:
        reemplazar_texto_en_parrafo(p, reemplazos)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    reemplazar_texto_en_parrafo(p, reemplazos)

    for section in doc.sections:
        for p in section.header.paragraphs:
            reemplazar_texto_en_parrafo(p, reemplazos)
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        reemplazar_texto_en_parrafo(p, reemplazos)

        for p in section.footer.paragraphs:
            reemplazar_texto_en_parrafo(p, reemplazos)
        for table in section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        reemplazar_texto_en_parrafo(p, reemplazos)

        if section.different_first_page_header_footer:
            for p in section.first_page_header.paragraphs:
                reemplazar_texto_en_parrafo(p, reemplazos)
            for table in section.first_page_header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            reemplazar_texto_en_parrafo(p, reemplazos)

            for p in section.first_page_footer.paragraphs:
                reemplazar_texto_en_parrafo(p, reemplazos)
            for table in section.first_page_footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            reemplazar_texto_en_parrafo(p, reemplazos)

@app.post("/detectar-campos")
async def detectar_campos(archivo_word: UploadFile = File(...)):
    try:
        if not archivo_word.filename.endswith(".docx"):
            raise HTTPException(400, "Solo se aceptan archivos .docx")
        contenido = await archivo_word.read()
        doc = Document(BytesIO(contenido))
        campos = extraer_todos_los_campos(doc)
        return JSONResponse({"campos": campos})
    except Exception as e:
        raise HTTPException(500, f"Error al detectar campos: {str(e)}")

@app.post("/procesar-manual")
async def procesar_manual(
    archivo_word: UploadFile = File(...),
    replacements: str = Form(...)
):
    try:
        # Obtenemos el nombre del archivo original
        nombre_original = archivo_word.filename
        
        reemplazos = json.loads(replacements)
        contenido = await archivo_word.read()
        doc = Document(BytesIO(contenido))
        reemplazar_campos(doc, reemplazos)
        salida = BytesIO()
        doc.save(salida)
        salida.seek(0)
        
        # Usamos el nombre original en el header Content-Disposition
        return Response(
            content=salida.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{nombre_original}"'}
        )
    except Exception as e:
        raise HTTPException(500, f"Error en procesamiento manual: {str(e)}")

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    return templates.TemplateResponse(
        "error.html",
        {"request": request, "status_code": exc.status_code, "detail": exc.detail},
        status_code=exc.status_code,
    )
