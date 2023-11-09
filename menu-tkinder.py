import os
from tkinter import *
#Buscar archivos
from tkinter import filedialog
#PDF a Word
from pdf2docx import Converter
#Word a pdf
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def ventana():
    global ventana
    ventana=Tk()                                                    #Crear ventana
    ventana.geometry("400x300") #ancho largo
    ventana.title("Programa")

    etiqueta=Label(ventana,text="Bienvenido Usuario", bg="#808080") #Agrega el texto #bg es el fondo
    etiqueta.pack(fill=X, pady=20)                                           #inserta el texto  #el Fill es para el relleno, pero se puede sacar

    boton_pdf_word=Button(ventana, text="PDF a Word", width=20, command=pdf_word) #command es la funcion #el place es la posicion del boton
    boton_pdf_word.pack()


    boton_word_pdf=Button(ventana, text="Word a PDF", width=20, command=word_pdf) #command es la funcion #el place es la posicion del boton
    boton_word_pdf.pack()

    boton_salir=Button(ventana, text="Salir", width=20, command=salida) #command es la funcion #el place es la posicion del boton
    boton_salir.pack()

    ventana.mainloop()
def salida():
        ventana.destroy()

def pdf_word():
    global pdf_word
    pdf_word=Toplevel(ventana)                                      #Se usa el Toplevel y no el Tk porque da BUGS
    pdf_word.geometry("400x300")
    pdf_word.title("PDF a Word")
    etiqueta=Label(pdf_word,text="Convertir de pdf a word", bg="#808080")
    etiqueta.pack(fill=X)

    ##Codigo
    

    def buscador_pdf():
        pdf_file = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        pdf_dir = os.path.dirname(pdf_file)

        if not pdf_file:
            print("No se seleccionó ningún archivo PDF.")
        else:
            # Obtener el directorio del archivo PDF
            pdf_dir = os.path.dirname(pdf_file)

            # Crear una nueva ventana para ingresar el nombre del archivo DOCX
            nombre_docx_window = Toplevel(pdf_word)
            nombre_docx_window.title("Nombre del archivo DOCX")
            nombre_etiqueta = Label(nombre_docx_window, text="Ingrese el nombre del archivo DOCX de salida (sin extensión .docx):")
            nombre_etiqueta.pack()

            nombre_docx_entry = Entry(nombre_docx_window)
            nombre_docx_entry.pack()

            def convertir_pdf_to_docx():
                docx_name = nombre_docx_entry.get()
                if docx_name:
                    # Generar la ruta completa del archivo DOCX de salida en el mismo directorio
                    docx_file = os.path.join(pdf_dir, docx_name + ".docx")

                    # Realiza la conversión de PDF a DOCX
                    cv = Converter(pdf_file)
                    cv.convert(docx_file, start=0, end=None)
                    cv.close()

                    print(f"Se ha convertido {pdf_file} a {docx_file}")

                nombre_docx_window.destroy()

            convertir_button = Button(nombre_docx_window, text="Convertir", command=convertir_pdf_to_docx)
            convertir_button.pack()

        buscador_pdf.mainloop()


    boton_pdf_word=Button(pdf_word, text="Buscar PDF", width=20, command=buscador_pdf) #command es la funcion #el place es la posicion del boton
    boton_pdf_word.pack(pady=20)
    ##

    boton=Button(pdf_word, text="volver al inicio", width=20, command=volver_ventana)
    boton.pack(pady=20)

    if(pdf_word):
        ventana.withdraw()

    pdf_word.mainloop()

def word_pdf():
    global word_pdf
    word_pdf=Toplevel(ventana)                                      #Se usa el Toplevel y no el Tk porque da BUGS
    word_pdf.geometry("400x300")
    word_pdf.title("Word a PDF")
    etiqueta=Label(word_pdf,text="Convertir de word a pdf", bg="#808080")
    etiqueta.pack(fill=X)

    ##Codigo
    
    def convert_docx_to_pdf():
        file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        
        if file_path:
            pdf_file_path = file_path.replace(".docx", ".pdf")
            doc = SimpleDocTemplate(pdf_file_path, pagesize=letter)
            styles = getSampleStyleSheet()
            pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
            
            try:
                docx = Document(file_path)
                pdf_content = []
                
                for para in docx.paragraphs:
                    text = para.text
                    p = Paragraph(text, styles['Normal'])
                    pdf_content.append(p)
                
                doc.build(pdf_content)
                result_label.config(text="Conversión exitosa: " + pdf_file_path)
            except Exception as e:
                result_label.config(text="Error en la conversión: " + str(e))

    result_label = Label(word_pdf, text="")
    result_label.pack(pady=20)

    convert_button = Button(word_pdf, text="Convertir DOCX a PDF", command=convert_docx_to_pdf)
    convert_button.pack(pady=20)

    
    ##

    boton=Button(word_pdf, text="volver al inicio", width=20, command=volver_ventana)
    boton.pack(pady=20)

    if(word_pdf):
        ventana.withdraw()

    word_pdf.mainloop()

def volver_ventana():
    ventana.iconify()
    ventana.deiconify()

    #destruir
    pdf_word.destroy()
    word_pdf.destroy()
    

ventana()                                                           #cierre app