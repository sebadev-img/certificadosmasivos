import pandas as pd
import io
import os
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from pypdf import PdfReader, PdfWriter

# ================= CONFIGURACIÓN =================
ARCHIVO_EXCEL = "datos.xlsx"       
ARCHIVO_PDF_BASE = "templateDocente.pdf" 
CARPETA_SALIDA = "certificados_emitidos"

COL_NOMBRE = "Nombre"
COL_APELLIDO = "Apellido"
COL_DNI = "DNI"

# --- Coordenadas Verticales ---
# IMPORTANTE: Como el tamaño es desconocido, esta altura es desde el borde INFERIOR hacia arriba.
# Tendrás que probar: si el PDF es pequeño, quizás 290 sea muy alto.
COORDENADA_Y_ALTURA = 290  

# --- Estilo del Texto ---
TAMANO_FUENTE = 14
FUENTE = "Helvetica-Bold"
COLOR_TEXTO = HexColor('#333333')
# =================================================

def crear_capa_texto_adaptable(texto, ancho_real, alto_real):
    """
    Crea un PDF temporal con las medidas EXACTAS del template original.
    """
    packet = io.BytesIO()
    
    # Usamos el tamaño detectado del PDF base, no A4
    c = canvas.Canvas(packet, pagesize=(ancho_real, alto_real))
    
    c.setFont(FUENTE, TAMANO_FUENTE)
    c.setFillColor(COLOR_TEXTO)
    
    # Calculamos el centro matemático exacto de ESTE documento
    centro_x = ancho_real / 2
    
    # Dibujamos el texto centrado
    c.drawCentredString(centro_x, COORDENADA_Y_ALTURA, texto)
    
    c.save()
    packet.seek(0)
    return packet

def procesar_diplomas():
    if not os.path.exists(CARPETA_SALIDA):
        os.makedirs(CARPETA_SALIDA)

    try:
        print(f"Leyendo archivo: {ARCHIVO_EXCEL}...")
        df = pd.read_excel(ARCHIVO_EXCEL, dtype=str).fillna("")
        
        # --- PASO NUEVO: Analizar el PDF Base antes de empezar ---
        lector_base_inicial = PdfReader(ARCHIVO_PDF_BASE)
        pagina_modelo = lector_base_inicial.pages[0]
        
        # Obtenemos ancho y alto real (en puntos) del template
        ancho_pdf = float(pagina_modelo.mediabox.width)
        alto_pdf = float(pagina_modelo.mediabox.height)
        
        print(f"📏 Dimensiones detectadas del template: {ancho_pdf:.2f} x {alto_pdf:.2f} puntos")
        print(f"📍 El centro horizontal exacto será: {ancho_pdf / 2:.2f}")
        
        total = len(df)
        
        for index, fila in df.iterrows():
            nombre = fila[COL_NOMBRE].strip()
            apellido = fila[COL_APELLIDO].strip()
            dni = fila[COL_DNI].strip()
            
            if dni.endswith(".0"): dni = dni[:-2]

            texto_completo = f"{nombre} {apellido} | DNI {dni}".upper()
            
            if not nombre or not apellido:
                continue

            # 1. Preparar el PDF Base
            lector_base = PdfReader(ARCHIVO_PDF_BASE)
            escritor = PdfWriter()
            pagina_base = lector_base.pages[0]

            # 2. Crear la capa usando las dimensiones DETECTADAS (ancho_pdf, alto_pdf)
            capa_io = crear_capa_texto_adaptable(texto_completo, ancho_pdf, alto_pdf)
            lector_capa = PdfReader(capa_io)
            pagina_capa = lector_capa.pages[0]

            # 3. Fusionar (Ahora ambas capas tienen el mismo tamaño exacto)
            pagina_base.merge_page(pagina_capa)
            escritor.add_page(pagina_base)

            # 4. Guardar
            nombre_archivo = f"{apellido}_{nombre}_{dni}".replace(" ", "_")
            nombre_archivo = "".join([c for c in nombre_archivo if c.isalnum() or c in ('_','-')])
            ruta_salida = os.path.join(CARPETA_SALIDA, f"Certificado_{nombre_archivo}.pdf")
            
            with open(ruta_salida, "wb") as f_out:
                escritor.write(f_out)
            
            print(f"[{index + 1}/{total}] OK: {texto_completo}")

        print(f"\n✅ Proceso finalizado.")

    except Exception as e:
        print(f"❌ ERROR: {e}")

if __name__ == "__main__":
    procesar_diplomas()