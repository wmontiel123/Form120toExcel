import tkinter as tk
from tkinter import filedialog
from tkinter import Menu
import PyPDF2

# Lista para almacenar las rutas de los archivos PDF seleccionados
pdf_files = []

def abrir_archivos():
    archivos = filedialog.askopenfilenames(
        title="Seleccionar archivo(s)",
        filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
    )
    
    for archivo in archivos:
        pdf_files.append(archivo)
        lista_archivos.insert(tk.END, archivo)  # Agregar el nombre del archivo a la lista

def combinar_archivos():
    if pdf_files:
        # Abre el archivo PDF de salida en modo escritura
        output_pdf = PyPDF2.PdfFileWriter()

        # Función para agregar un archivo PDF al PDF de salida
        def add_pdf_to_output(file_path):
            pdf = PyPDF2.PdfFileReader(file_path)
            for page_num in range(len(pdf.pages)):
                page = pdf.pages[page_num]
                output_pdf.addPage(page)

        # Agrega los archivos PDF en el orden especificado
        for pdf_file in pdf_files:
            add_pdf_to_output(pdf_file)

        archivo_salida = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivo PDF", "*.pdf")],
            title="Guardar PDF combinado como"
        )

        # Guarda el PDF combinado en un archivo
        output_file_path = archivo_salida
        with open(output_file_path, "wb") as output_file:
            output_pdf.write(output_file)

        print(f"PDF combinado exitosamente y guardado en {output_file_path}")
    else:
        print("No se han seleccionado archivos para combinar.")

def abrir_ventana_operaciones():
    # Crea una nueva ventana para las operaciones
    ventana_operaciones = tk.Toplevel(ventana)
    ventana_operaciones.title("Operaciones PDF")
    ventana_operaciones.geometry("400x200")  # Tamaño de la ventana de operaciones

    # Agrega contenido a la ventana de operaciones
    etiqueta = tk.Label(ventana_operaciones, text="Seleccione una operación:")
    etiqueta.pack(pady=10)

    def extraer_paginas():
        pass  # Implementa la extracción de páginas aquí

    def dividir_paginas():
        pass  # Implementa la división de páginas aquí

    # Crea un menú desplegable con las operaciones disponibles
    menu_operaciones = Menu(ventana_operaciones)
    menu_operaciones.add_command(label="Extraer Páginas", command=extraer_paginas)
    menu_operaciones.add_command(label="Dividir en Páginas Independientes", command=dividir_paginas)
    ventana_operaciones.config(menu=menu_operaciones)

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Combinar y Operar con Archivos PDF")
ventana.geometry("800x300")  # Establecer el tamaño de la ventana

# Crear una lista tkinter para mostrar los archivos seleccionados
lista_archivos = tk.Listbox(ventana, selectmode=tk.MULTIPLE)
lista_archivos.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Rellenar y expandir para llenar la ventana

# Crear un botón para abrir el cuadro de diálogo de selección de archivo(s)
boton_abrir = tk.Button(ventana, text="Seleccionar Archivo(s)", command=abrir_archivos)
boton_abrir.pack(pady=10)

# Crear un botón para combinar los archivos seleccionados
boton_combinar = tk.Button(ventana, text="Combinar Archivos", command=combinar_archivos)
boton_combinar.pack(pady=10)

# Crear un botón para abrir la ventana de operaciones
boton_operaciones = tk.Button(ventana, text="Operaciones PDF", command=abrir_ventana_operaciones)
boton_operaciones.pack(pady=10)

# Iniciar el bucle principal de la GUI
ventana.mainloop()
