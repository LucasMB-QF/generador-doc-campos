from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import Response, HTMLResponse
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from docx import Document
import re
from io import BytesIO

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# Regex para detectar campos tipo {{Hoja!Celda}} o {{Hoja!Rango}}
campo_regex = re.compile(r"\{\{\s*([^\{\}]+?)\s*\}\}")

def obtener_valor(wb, hoja_nombre, celda):
    try:
        hoja = wb[hoja_nombre]
        celda_valor = hoja[celda].value
        return str(celda_valor) if celda_valor is not None else ""
    except Exception as e:
        print(f"[ERROR] Celda {hoja_nombre}!{celda}: {e}")
        return ""

def obtener_valores_rango(wb, hoja_nombre, rango):
    try:
        hoja = wb[hoja_nombre]
        celdas = hoja[rango]
        fila = celdas[0]
        return [str(c.value) if c.value is not None else "" for c in fila]
    except Exception as e:
        print(f"[ERROR] Rango {hoja_nombre}!{rango}: {e}")
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
async def procesar(archivo_excel: UploadFile = File(...), archivo_word: UploadFile = File(...)):
    try:
        excel_content = await archivo_excel.read()
        word_content = await archivo_word.read()

        excel_stream = BytesIO(excel_content)
        word_stream = BytesIO(word_content)

        wb = load_workbook(filename=excel_stream, data_only=True)
        doc = Document(word_stream)

        procesar_documento(doc, wb)

        output_stream = BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)

        return Response(
            content=output_stream.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=informe_generado.docx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")
