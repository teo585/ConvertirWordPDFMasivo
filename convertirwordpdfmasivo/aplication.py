import os
import comtypes.client # type: ignore

def convertir_docx_a_pdf(input_file, output_file):
    # Crear una instancia de Word
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    
    # Abrir el archivo
    doc = word.Documents.Open(input_file)
    
    # Guardar el archivo como PDF
    doc.SaveAs(output_file, FileFormat=17)  # 17 es el formato PDF
    
    # Cerrar el archivo y la aplicación
    doc.Close()
    word.Quit()



def buscar_y_convertir(root_folder):
    pass 

# Carpeta principal donde se encuentran las carpetas con archivos Word
carpeta_principal = 'la ruta donde está la carpeta entera'

# Ejecutar la función
buscar_y_convertir(carpeta_principal)