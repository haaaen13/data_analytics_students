import logging
import subprocess
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def choose_file(title="Selecciona un archivo", filetypes=(("Todos los archivos", "*.*"),)):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=title, filetypes=filetypes)

def choose_directory(title="Selecciona una carpeta"):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title=title)

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

        # Ejemplo de archivo generado
        save_path = os.path.join(images_dir, "resultado.csv")
        return save_path

    except subprocess.CalledProcessError as e:
        logging.error("La ejecución de FormScanner falló.")
        print("STDOUT:\n", e.stdout)
        print("STDERR:\n", e.stderr)
        raise

def run_formscanner_workflow():
    """
    Función principal que ejecuta el flujo de selección de archivos,
    ejecución de FormScanner y repetición opcional.
    Puede ser llamada desde el main de otra aplicación.
    """
    logging.basicConfig(level=logging.INFO)

    while True:
        formScanner_executable = r"C:\Program Files (x86)\FormScanner_1.1.4\lib\formscanner-main-1.1.4.jar"
        formScanner_template   = choose_file("Selecciona el archivo template (.xtmpl)", [("XTMPL files", "*.xtmpl")])
        images_dir             = choose_directory("Selecciona la carpeta con las imágenes")

        save_path = execute_FormScanner(formScanner_executable, formScanner_template, images_dir)

        repetir = messagebox.askyesno(
            "Proceso completado",
            f"Archivo guardado como:\n{save_path}\n\n¿Deseas procesar otro archivo?"
        )
        if not repetir:
            print("Finalizando ejecución.")
            break