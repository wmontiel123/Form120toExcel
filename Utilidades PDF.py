import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox

def rotate_pdf(input_pdf_paths, output_pdf_path, degrees, save_separately):
    try:
        if save_separately:
            for input_pdf_path in input_pdf_paths:
                pdf_file = open(input_pdf_path, 'rb')
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    page.rotate(degrees)
                    output_pdf_path = f"{input_pdf_path}_{page_num + 1}.pdf"
                    
                    pdf_writer = PyPDF2.PdfWriter()
                    pdf_writer.add_page(page)
                    
                    with open(output_pdf_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    
                pdf_file.close()
                
            print(f'Páginas rotadas {degrees} grados y guardadas en archivos distintos.')
        else:
            pdf_writer = PyPDF2.PdfWriter()
            
            for input_pdf_path in input_pdf_paths:
                pdf_file = open(input_pdf_path, 'rb')
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    page.rotate(degrees)
                    pdf_writer.add_page(page)
                
                pdf_file.close()
            
            with open(output_pdf_path, 'wb') as output_file:
                pdf_writer.write(output_file)
                
            print(f'Páginas rotadas {degrees} grados y guardadas en un solo archivo PDF: {output_pdf_path}')

    except Exception as e:
        print(f'Error: {e}')

def rotate_left():
    input_pdf_paths = filedialog.askopenfilenames(
        title="Seleccionar archivos PDF de entrada",
        filetypes=[("Archivos PDF", "*.pdf")]
    )

    if not input_pdf_paths:
        print("No se seleccionó ningún archivo PDF.")
        return
    output_pdf_path = input_pdf_paths
    save_separately = messagebox.askyesno("Guardar Separadamente", "¿Desea guardar las páginas rotadas en archivos distintos?")
    if save_separately is None:  # El usuario cerró la ventana de confirmación
        return
    if save_separately is False:
        output_pdf_path = filedialog.asksaveasfilename(
            title="Guardar archivo PDF de salida",
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")]
        )

        if not output_pdf_path:
            print("No se proporcionó un nombre de archivo de salida.")
            return


    rotate_pdf(input_pdf_paths, output_pdf_path, 90, save_separately)

def rotate_right():
    input_pdf_paths = filedialog.askopenfilenames(
        title="Seleccionar archivos PDF de entrada",
        filetypes=[("Archivos PDF", "*.pdf")]
    )

    if not input_pdf_paths:
        print("No se seleccionó ningún archivo PDF.")
        return
    output_pdf_path = input_pdf_paths
    save_separately = messagebox.askyesno("Guardar Separadamente", "¿Desea guardar las páginas rotadas en archivos distintos?")
    if save_separately is None:  # El usuario cerró la ventana de confirmación
        return

    if save_separately is False:
        output_pdf_path = filedialog.asksaveasfilename(
            title="Guardar archivo PDF de salida",
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")]
        )

        if not output_pdf_path:
            print("No se proporcionó un nombre de archivo de salida.")
            return


    rotate_pdf(input_pdf_paths, output_pdf_path, 270, save_separately)

def rotate_180():
    input_pdf_paths = filedialog.askopenfilenames(
        title="Seleccionar archivos PDF de entrada",
        filetypes=[("Archivos PDF", "*.pdf")]
    )

    if not input_pdf_paths:
        print("No se seleccionó ningún archivo PDF.")
        return

    save_separately = messagebox.askyesno("Guardar Separadamente", "¿Desea guardar las páginas rotadas en archivos distintos?")
    if save_separately is None:  # El usuario cerró la ventana de confirmación
        return
    output_pdf_path = input_pdf_paths
    if save_separately is False:
        output_pdf_path = filedialog.asksaveasfilename(
            title="Guardar archivo PDF de salida",
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")]
        )

        if not output_pdf_path:
            print("No se proporcionó un nombre de archivo de salida.")
            return

    rotate_pdf(input_pdf_paths, output_pdf_path, 180, save_separately)

def main():
    root = tk.Tk()
    root.title("Rotación de PDF")
    root.geometry("400x200")

    rotate_left_button = tk.Button(root, text="Rotar a la Izquierda", command=rotate_left)
    rotate_left_button.pack(pady=10)

    rotate_right_button = tk.Button(root, text="Rotar a la Derecha", command=rotate_right)
    rotate_right_button.pack(pady=10)

    rotate_180_button = tk.Button(root, text="Rotar 180°", command=rotate_180)
    rotate_180_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
