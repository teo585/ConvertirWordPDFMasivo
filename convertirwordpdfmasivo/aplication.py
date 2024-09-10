import os
import comtypes.client
from tqdm import tqdm

def convertir_docx_a_pdf(input_file, output_file):
    """
    Convierte un archivo DOCX a PDF utilizando Microsoft Word.
    
    :param input_file: Ruta completa del archivo DOCX de entrada.
    :param output_file: Ruta completa del archivo PDF de salida.
    """
    if not os.path.exists(input_file):
        print(f"Error: No se encontró el archivo {input_file}")
        return

    # Crear una instancia de Word
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False

    # Normalizar las rutas
    input_file = os.path.normpath(input_file)
    output_file = os.path.normpath(output_file)

    try:
        # Abrir el documento
        doc = word.Documents.Open(input_file)
        # Guardar como PDF
        doc.SaveAs(output_file, FileFormat=17)  # 17 es el formato PDF
    except Exception as e:
        print(f"No se pudo convertir el archivo {input_file} a PDF.")
        print(f"Error: {e}")
    finally:
        # Cerrar el documento y salir de la aplicación
        doc.Close()
        word.Quit()

def buscar_y_convertir(root_folder):
    """
    Busca archivos DOCX en la carpeta raíz y sus subcarpetas, y los convierte a PDF.
    
    :param root_folder: Ruta de la carpeta raíz donde se buscan los archivos DOCX.
    """
    todos_los_archivos = buscar_archivos_docx(root_folder)
    
    # Usar tqdm para mostrar una barra de progreso durante la conversión
    for archivo_completo in tqdm(todos_los_archivos, desc="Convirtiendo archivos"):
        carpeta_raiz, archivo = os.path.split(archivo_completo)
        nombre_pdf = os.path.join(carpeta_raiz, archivo.replace("_", "-").replace(".docx", ".pdf"))
        convertir_docx_a_pdf(archivo_completo, nombre_pdf)
        print(f"Convertido: {archivo_completo} -> {nombre_pdf}")

    # Mostrar las rutas de las carpetas encontradas
    mostrar_carpetas(root_folder)

def buscar_archivos_docx(root_folder):
    """
    Busca todos los archivos DOCX en una carpeta y sus subcarpetas.
    
    :param root_folder: Ruta de la carpeta raíz donde se buscan los archivos DOCX.
    :return: Lista de rutas completas de los archivos DOCX encontrados.
    """
    archivos_docx = []
    for carpeta_raiz, _, archivos in os.walk(root_folder):
        for archivo in archivos:
            if archivo.endswith(".docx"):
                archivo_completo = os.path.join(carpeta_raiz, archivo)
                archivos_docx.append(archivo_completo)
    return archivos_docx

def mostrar_carpetas(root_folder):
    """
    Muestra las carpetas encontradas dentro de la carpeta raíz.
    
    :param root_folder: Ruta de la carpeta raíz donde se buscan las carpetas.
    """
    for carpeta_raiz, subcarpetas, _ in os.walk(root_folder):
        for subcarpeta in subcarpetas:
            print(f"Encontrada carpeta de trabajador o mes: {os.path.join(carpeta_raiz, subcarpeta)}")

def main():
    """
    Función principal para solicitar la ruta de la carpeta y ejecutar el proceso de conversión.
    """
    carpeta_principal = input("Por favor, introduce la ruta de la carpeta principal: ")
    buscar_y_convertir(carpeta_principal)

if __name__ == "__main__":
    main()
