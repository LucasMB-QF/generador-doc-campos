from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response
from openpyxl import load_workbook
from docx import Document
import re
from io import BytesIO

app = FastAPI()

# ... (las funciones obtener_valor, obtener_valores_rango, reemplazar_campos y procesar_documento se mantienen igual) ...

@app.post("/procesar")
async def procesar(archivo_excel: UploadFile = File(...), archivo_word: UploadFile = File(...)):
    try:
        # Leer contenido directamente en memoria
        excel_content = await archivo_excel.read()
        word_content = await archivo_word.read()
        
        # Crear streams en memoria
        excel_stream = BytesIO(excel_content)
        word_stream = BytesIO(word_content)
        
        # Procesar documentos
        wb = load_workbook(filename=excel_stream, data_only=True)
        doc = Document(word_stream)
        
        procesar_documento(doc, wb)
        
        # Guardar resultado en memoria
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
