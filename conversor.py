import PyPDF2
import re
from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog


def leer_pdf(ruta):
    lineas = []
    with open(ruta, 'rb') as archivo:
        lector_pdf = PyPDF2.PdfReader(archivo)
        for pagina in lector_pdf.pages:
            contenido = pagina.extract_text()
            lineas.extend(contenido.split('\n'))
    return lineas

def extract_text(linea, inicio, fin1, fin2):
   # Find the position of the inicio string
    inicio_pos = linea.find(inicio)
    if inicio_pos == -1:
        return None
    fin_pos = len(linea)+2
    if fin1 != "" and fin2 !="":
        fin_pos = linea.find(fin1, inicio_pos)

    # Find the position of the fin2 string
    if fin_pos == -1:
        fin_pos = linea.find(fin2, inicio_pos)+1
        inicio_pos = inicio_pos
        
    if fin_pos == -1:
        return None 



    # Extract the text between the inicio_pos and fin_pos positions
    texto_extraido = linea[(inicio_pos) + len(inicio):fin_pos]
    return texto_extraido.strip()  # Remove leading/trailing whitespaces


# Ruta del archivo PDF
# Create a Tkinter root window
root = tk.Tk()
root.withdraw()  # Hide the root window

# Open the file dialog to select a file
file_path = filedialog.askopenfilename()

# Print the selected file path
print("Selected File Path:", file_path)



ruta_pdf = file_path
#Carga de Valores Para La Carga
Valores_Obtenidos = []
values = [
    ("vados con tasa del 10%10", " 22 ", " 022 "),
    (" 22 ", "", ""),
    (" 022 ", "", ""),
    ("con tasa del 5%150", " 156 ", "0156 "),
    (" 156 ", "", ""),
    (" 0156 ", "", ""),
    ("gravados con tasa del 5%151", " 157 ", "0157 "),
    (" 157 ", "", ""),
    (" 0157 ", "", ""),
    ("no alcanzados por el Impuesto12", "", ""),
    ("industrializacion152 ", "", ""),
    ("bienes153 ", "", ""),
    ("bienes 14", "", ""),
    ("tasa del 10%15", " 23 ", "023 "),
    (" 23 ", "", ""),
    (" 023 ", "", ""),
    ("proceso de elaboracion o industrializacion154", " 158 ", "0158 "),
     (" 158 ", "", ""),
    (" 0158 ", "", ""),
    ("servicios155 ", " 159 ", "0159 "),
    (" 0159 ", "", ""),
    (" 159 ", "", ""),
    ("Impuesto17 ", "", ""),
    ("Inc. a+h)18", " 21 ", "021 "),
    (" 21 ", " 24 ", "024 "),
    (" 021 ", " 24 ", "024 "),
    (" 24 ", "", ""),
    (" 024 ", "", ""),
    ("natural160", "", ""),
    ("interno161", "", ""),
    ("no alcanzados 26", "", ""),
    ("(Inc. a+b+c) 27", "", ""),
    ("industrializacion162", "", ""),
    ("bienes 163", "", ""),
    ("bienes 29", "", ""),
    ("(Inc. e+f+g) 30", "", ""),
    ("(Inc. d+h) 31", "", ""),
    ("interno32 ", " 35 ", "035 "),
    (" 35 ", " 38 ", "038 "),
    (" 035 ", " 38 ", "038 "),
    (" 038 ", "", ""),
    (" 38 ", "", ""),
    ("alcanzadas33", " 36 ", "036 "),
    (" 36 ", " 39 ", "039 "),
    (" 036 ", " 39 ", "039 "),
    (" 39 ", "", ""),
    (" 039 ", "", ""),
    ("Rubro 2 Inc. a+b/ Inc. d)164", "", ""),
    ("calculo)165", "", ""),
    ("incobrables34", " 37 ", "037 "),
    (" 37 ", " 42 ", "042 "),
    (" 037 ", " 42 ", "042 "),
    (" 42 ", "", ""),
    (" 042 ", "", ""),
    ("a+c+d+e)43", "", ""),
    ("Inc. l) 44", "", ""),
    (" Inc. f) 45", "", ""),
    ("anterior)46", "", ""),
    ("Inc. c 166", "", ""),
    ("d167 ", "", ""),
    ("e47", "", ""),
    ("Inc. c 48", "", ""),
    (" Exportador)49", "", ""),
    ("(No trasladable)168", "", ""),
    ("i) 50", "", ""),
    ("Inc. j) 55", "", ""),
    ("anterior)51", "", ""),
    ("recibidas52", "", ""),
    ("gravadas 169", "", ""),
    ("vencimiento 56", "", ""),
    (" a+e) 53", " 57 ", "057 "),
    (" 57 ", "", ""),
    (" 057 ", "", ""),
    ("Rubro 454", "", ""),
    ("Col. I)58", "", ""),
    ("Impuesto59", " 65 ", "065 "),
    (" 65 ", "", ""),
    (" 065 ", "", ""),
    ("Impuesto60", " 66 ", "066 "),
    (" 66 ", "", ""),
    (" 066 ", "", ""),
    ("exportaciones61", "", ""),
    ("Impuesto62", "", ""),
    ("Impuesto63", "", ""),
    ("IRP 64 ", "", ""),
    ("IRP170 ", "", ""),
    
    
]



# Llamada a la función para leer el PDF y guardar las líneas en una lista
lineas_pdf = leer_pdf(ruta_pdf)

# Imprimir las líneas del PDF


for linea in lineas_pdf:
    
    for inicio, fin1, fin2 in values:
        texto_extraido = extract_text(linea, inicio, fin1, fin2)
        if texto_extraido:
            #print(linea)
            print(inicio+' '+fin1+' '+fin2+' '+texto_extraido)
            Valores_Obtenidos.append(texto_extraido)

# Crear un nuevo libro de Excel y seleccionar la hoja activa
wb = Workbook()
sheet = wb.active

# Datos que deseas copiar en la columna B
datos = ["10",
         "22",
         "150",
         "156",
         "151",
         "157",
         "12",
         "152",
         "153",
         "14",
         "15",
         "23",
         "154",
         "158",
         "155",
         "159",
         "17",
         "18",
         "21",
         "24",
         "160",
         "161",
         "26",
         "27",
         "162",
         "163",
         "29",
         "30",
         "31",
         "32",
         "35",
         "38",
         "33",
         "36",
         "39",
         "164",
         "165",
         "34",
         "37",
         "42",
         "43",
         "44",
         "45",
         "46",
         "166",
         "167",
         "47",
         "48",
         "49",
         "168",
         "50",
         "55",
         "51",
         "52",
         "169",
         "56",
         "53",
         "57",
         "54",
         "58",
         "59",
         "65",
         "60",
         "66",
         "61",
         "62",
         "63",
         "64",
         "170"]

# Copiar los datos en la columna B de Excel
for i, dato in enumerate(datos, start=1):
    sheet.cell(row=i, column=1).value = dato
    
for i, dato in enumerate(Valores_Obtenidos, start=1):
    sheet.cell(row=i, column=2).value = dato

# Guardar el archivo de Excel con los datos copiados
# Create a Tkinter root window
root = tk.Tk()
root.withdraw()  # Hide the root window

# Open the file dialog to select a file path for saving
file_path = filedialog.asksaveasfilename()

# Print the selected file path
print("Selected File Path for Saving:", file_path)

nombre_archivo = file_path

wb.save(nombre_archivo)
            
    

