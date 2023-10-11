import PyPDF2
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog

def dividir_pdf(input_pdf, output_folder, group_size=1):
    pdf = PyPDF2.PdfReader(open(input_pdf, 'rb'))

    for page_group in range(0, len(pdf.pages), group_size):
        pdf_writer = PyPDF2.PdfWriter()
        output_pdf = f"{output_folder}/pagina_{page_group + 1}-{page_group + group_size}.pdf"

        for page_num in range(page_group, min(page_group + group_size,len(pdf.pages))):
            pdf_writer.add_page(pdf.pages[page_num])

        with open(output_pdf, 'wb') as output_file:
            pdf_writer.write(output_file)

def main():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal

    input_pdf = filedialog.askopenfilename(title="Seleccionar archivo PDF de entrada")
    if not input_pdf:
        return

    group_size = simpledialog.askinteger("Tamaño del grupo de páginas", "Ingrese el tamaño del grupo de páginas:",
                                         initialvalue=1, minvalue=1)

    if group_size is None:
        return

    output_folder = filedialog.askdirectory(title="Seleccionar carpeta de salida")
    if not output_folder:
        return

    dividir_pdf(input_pdf, output_folder, group_size)
    tk.messagebox.showinfo("Proceso completado", "El PDF ha sido dividido con éxito.")

if __name__ == "__main__":
    main()
