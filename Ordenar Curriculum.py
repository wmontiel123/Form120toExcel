import os
import docx2txt
import PyPDF2

# Solicitar los valores de los parámetros al usuario
años_experiencia = input("Ingrese años de experiencia: ")
titulo_grado = input("Ingrese título de grado: ")
maestria = input("Ingrese maestría: ")
doctorado = input("Ingrese doctorado: ")
experiencia_proyectos = input("Ingrese experiencia en proyectos: ")
cargos = input("Ingrese cargos: ")
palabras_clave = input("Ingrese palabras clave separadas por coma: ").split(",")

# Carpeta de entrada y salida
carpeta_entrada = r'C:\\Users\\Walter Montiel\\OneDrive - COVENTRY S.A\\Escritorio\\Curriculums Vitae'
carpeta_salida = r'C:\\Users\\Walter Montiel\\OneDrive - COVENTRY S.A\\Escritorio\\Aprobados'
# Crear carpeta de salida si no existe
if not os.path.exists(carpeta_salida):
    os.makedirs(carpeta_salida)

def buscar_requisitos(texto):
    # Función para buscar requisitos en el texto
    texto = texto.lower()
    return (
        años_experiencia in texto
        and titulo_grado.lower() in texto
        and maestria.lower() in texto
        and doctorado.lower() in texto
        and experiencia_proyectos.lower() in texto
        and any(palabra.lower() in texto for palabra in palabras_clave)
    )

# Recorrer los archivos en la carpeta de entrada
for archivo in os.listdir(carpeta_entrada):
    archivo_path = os.path.join(carpeta_entrada, archivo)
    
    # Procesar archivos Word
    if archivo.endswith(".docx"):
        texto = docx2txt.process(archivo_path)
        if buscar_requisitos(texto):
            # Guardar el archivo en la carpeta de salida
            destino = os.path.join(carpeta_salida, archivo)
            os.rename(archivo_path, destino)
            print(f"Archivo {archivo} cumple con los requisitos y ha sido guardado.")
    
    # Procesar archivos PDF
    elif archivo.endswith(".pdf"):
        with open(archivo_path, "rb") as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            texto = ""
            for page_num in pdf_reader.pages:
                texto += page_num.extract_text()
        if buscar_requisitos(texto):
            # Guardar el archivo en la carpeta de salida
            destino = os.path.join(carpeta_salida, archivo)
            os.rename(archivo_path, destino)
            print(f"Archivo {archivo} cumple con los requisitos y ha sido guardado.")
    
    else:
        print(f"Archivo {archivo} no es compatible y ha sido omitido.")

print("Proceso completado.")
