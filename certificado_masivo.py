import pandas as pd
from PyPDFForm import PdfWrapper
import os

# CONFIGURACIÓN
archivo_excel = "datos.xlsx"      # Nombre de tu archivo Excel
archivo_plantilla = "template.pdf" # Nombre de tu PDF rellenable
campo_pdf = "Nombre, apellido y DNI" # El nombre exacto del campo en el PDF
carpeta_salida = "certificados_generados"

# Crear la carpeta de salida si no existe
if not os.path.exists(carpeta_salida):
    os.makedirs(carpeta_salida)
    print(f"📂 Carpeta '{carpeta_salida}' creada.")
else:
    print(f"📂 Usando carpeta existente: '{carpeta_salida}'")

# 1. Leemos el Excel
try:
    # dtype=str asegura que el DNI se lea como texto y no pierda ceros o formato
    df = pd.read_excel(archivo_excel, dtype=str) 
except FileNotFoundError:
    print(f"❌ Error: No se encontró el archivo {archivo_excel}")
    exit()

# 2. Iteramos por cada fila del Excel
for index, row in df.iterrows():
    # Obtenemos los datos limpiando espacios vacíos alrededor
    nombre = row['Nombre'].strip()
    apellido = row['Apellido'].strip()
    dni = row['DNI'].strip()

    # 3. Creamos el texto con el formato solicitado: "NOMBRE APELLIDO | DNI X"
    # Usamos .upper() para pasar nombre y apellido a mayúsculas (opcional)
    texto_combinado = f"{nombre.upper()} {apellido.upper()} | DNI {dni}"

    print(f"Procesando: {texto_combinado}")

    # 4. Cargamos el formulario (debe hacerse dentro del bucle para tener una copia limpia cada vez)
    formulario = PdfWrapper(archivo_plantilla)

    # 5. Rellenamos los datos
    datos_a_rellenar = {
        campo_pdf: texto_combinado
    }
    formulario.fill(datos_a_rellenar)

    # 6. Guardamos el archivo individual
    # Generamos un nombre único, por ejemplo usando el DNI
    nombre_archivo = f"certificado_{dni}.pdf"

    # os.path.join une la carpeta y el archivo con la barra correcta (\ en Win, / en Mac/Linux)
    ruta_completa = os.path.join(carpeta_salida, nombre_archivo)
    
    with open(ruta_completa, "wb+") as output:
        output.write(formulario.read())
    
    print(f"   -> Guardado como: {nombre_archivo}")

print("\n✅ Proceso finalizado.")