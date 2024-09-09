import os
//import comtypes.client

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