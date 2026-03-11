import logging
import subprocess
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageEnhance

def choose_file(title="Selecciona un archivo", filetypes=(("Todos los archivos", "."),)):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=title, filetypes=filetypes)

def choose_directory(title="Selecciona una carpeta"):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title=title)

def procesar_imagenes(images_dir):
    """
    Recorta la parte central de cada imagen en la carpeta y aplica nitidez.
    Sobrescribe las imágenes originales.
    """
    for filename in os.listdir(images_dir):
        if filename.lower().endswith((".jpg", ".jpeg", ".png")):
            path = os.path.join(images_dir, filename)
            try:
                imagen = Image.open(path).convert("RGB")  # asegurar modo compatible
                ancho, alto = imagen.size

                # Definir recorte central (ejemplo: mitad de ancho y alto)
                nuevo_ancho = int(ancho / 1.40)
                nuevo_alto = int(alto / 1.30)
                x1 = int((ancho - nuevo_ancho) / 2)
                y1 = int((alto - nuevo_alto) / 2)
                x2 = x1 + nuevo_ancho
                y2 = y1 + nuevo_alto


                recorte = imagen.crop((x1, y1, x2, y2)).convert("RGB")
                
                def oscurecer_sombras(valor):
                    # Si el píxel es oscuro (<128), lo reducimos más
                    if valor < 128:
                        return int(valor * 0.7)  # multiplica por 0.7 para oscurecer
                    else:
                        return valor

                #Aplicar sombras al recorte
                recorte_sombreado = recorte.point(oscurecer_sombras)

                # Aplicar nitidez dos veces al recorte
                enhancer1 = ImageEnhance.Sharpness(recorte_sombreado)
                recorte_nitido = enhancer1.enhance(0.0)
                recorte_nitido = enhancer1.enhance(0.0)

                #FormScanner por alguna razon ocupa que toda la bolita de respuesta
                #sea consistentemente llenada y negra por lo que estos filtros hace
                #el escaneo de respuestas mas consistente.

                # Pegar el recorte modificado en la imagen original
                imagen.paste(recorte_nitido, (x1, y1))

                # Guardar sobrescribiendo
                imagen.save(path)
                logging.info(f"Procesada imagen: {filename}")

            except Exception as e:
                logging.error(f"Error procesando {filename}: {e}")

def execute_FormScanner(formScanner_executable, formScanner_template, images_dir):
    logging.info(f"Escaneando imágenes con FormScanner usando la plantilla: {formScanner_template}")

    formScanner_cmd = [
        "java",
        "-jar", formScanner_executable,
        formScanner_template,
        images_dir,
    ]

    logging.debug("Comando preparado: " + ' '.join(formScanner_cmd))

    try:
        result = subprocess.run(
            formScanner_cmd,
            check=True,
            capture_output=True,
            text=True
        )
        logging.info("Proceso completado.")
        print("STDOUT:\n", result.stdout)
        print("STDERR:\n", result.stderr)

        save_path = os.path.join(images_dir, "resultado.csv")
        return save_path

    except subprocess.CalledProcessError as e:
        logging.error("La ejecución de FormScanner falló.")
        print("STDOUT:\n", e.stdout)
        print("STDERR:\n", e.stderr)
        raise

def run_formscanner_workflow():
    logging.basicConfig(level=logging.INFO)

    while True:
        formScanner_executable = r"C:\Program Files (x86)\FormScanner_1.1.4\lib\formscanner-main-1.1.4.jar"
        formScanner_template   = choose_file("Selecciona el archivo template (.xtmpl)", [("XTMPL files", "*.xtmpl")])
        images_dir             = choose_directory("Selecciona la carpeta con las imágenes")

        # Procesar imágenes antes de pasarlas a FormScanner
        procesar_imagenes(images_dir)

        save_path = execute_FormScanner(formScanner_executable, formScanner_template, images_dir)

        repetir = messagebox.askyesno(
            "Proceso completado",
            f"Archivo guardado como:\n{save_path}\n\n¿Deseas procesar otro archivo?"
        )
        if not repetir:
            print("Finalizando ejecución.")
            break