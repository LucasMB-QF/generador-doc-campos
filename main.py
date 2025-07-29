from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from docx import Document
import re
import shutil
import os
import uuid

app = FastAPI()

@app.get("/")
async def root():
    return {"message": "La API está corriendo. Usa POST /procesar para generar informes."}

# Regex para detectar campos tipo {{Hoja!Celda}} o {{Hoja!Rango}}
campo_regex = re.compile(r"\{\{\s*([^\{\}]+?)\s*\}\}")

# Función para extraer valor de una celda
def obtener_valor(wb, hoja_nombre, celda):
    try:
        hoja = wb[hoja_nombre]
        celda_valor = hoja[celda].value
        return str(celda_valor) if celda_valor is not None else ""
    except Exception as e:
        print(f"[ERROR] Celda {hoja_nombre}!{celda}: {e}")
        return ""

# Función para extraer valores de un rango horizontal (una fila)
def obtener_valores_rango(wb, hoja_nombre, rango):
    try:
        hoja = wb[hoja_nombre]
        celdas = hoja[rango]
        fila = celdas[0]
        return [str(c.value) if c.value is not None else "" for c in fila]
    except Exception as e:
        print(f"[ERROR] Rango {hoja_nombre}!{rango}: {e}")
        return []

# Reemplaza campos del tipo {{Hoja!Celda}} o {{Hoja!P10:Y10}}
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

# Reemplaza en párrafos y celdas de tablas
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

@app.post("/procesar")
async def procesar(archivo_excel: UploadFile = File(...), archivo_word: UploadFile = File(...)):
    temp_id = str(uuid.uuid4())

    ruta_excel = f"temp_{temp_id}.xlsm"
    ruta_word = f"temp_{temp_id}.docx"
    ruta_salida = f"salida_{temp_id}.docx"

    # Guardar archivos temporales
    with open(ruta_excel, "wb") as f:
        shutil.copyfileobj(archivo_excel.file, f)
    with open(ruta_word, "wb") as f:
        shutil.copyfileobj(archivo_word.file, f)

    try:
        wb = load_workbook(filename=ruta_excel, data_only=True)
        doc = Document(ruta_word)

        procesar_documento(doc, wb)

        doc.save(ruta_salida)

        return FileResponse(
            ruta_salida,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="informe_generado.docx"
        )
    except Exception as e:
        return {"error": str(e)}
    finally:
        # Limpieza de archivos temporales
        for ruta in [ruta_excel, ruta_word]:
            try:
                os.remove(ruta)
            except:
                pass
