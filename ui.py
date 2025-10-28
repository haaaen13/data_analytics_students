import pyodbc
import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog
import main 


# Configurar apariencia
ctk.set_appearance_mode("dark")  # "light", "dark", "system"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

# Crear ventana principal
ventana = ctk.CTk()
ventana.title("Graficador de Calificaciones")
ventana.geometry("1280x720")



# --- Función para mostrar gráfico dinámico ---
def mostrar_grafico(nombre_columna):
    global fig_actual

    # Validar si la columna existe
    if nombre_columna not in main.base_datos.columns:
        error_label.configure(text=f"⚠ La columna '{nombre_columna}' no existe.", text_color="red")
        return
    else:
        error_label.configure(text="")  # limpiar mensaje de error
    
    # Obtener el conteo de valores únicos
    conteo = main.base_datos[nombre_columna].value_counts().head(20)
    
    # Crear figura
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar(conteo.index.astype(str), conteo.values, color='skyblue')
    

    fig_actual = fig

    # Configurar el gráfico
    ax.set_title(f'Conteo de registros por "{nombre_columna}"')
    ax.set_xlabel(nombre_columna)
    ax.set_ylabel("Cantidad")
    ax.set_xticks(range(len(conteo.index)))
    ax.set_xticklabels(conteo.index, rotation=45, ha='right')
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    
    # Mostrar en la interfaz
    for widget in frame_grafico.winfo_children():
        widget.destroy()  # eliminar gráfico anterior

    canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack()

# --- Función auxiliar para obtener texto del usuario ---
def generar_grafico():
    nombre_columna = entry_columna.get().strip()
    mostrar_grafico(nombre_columna)

def exportar_grafico():
    if 'fig_actual' not in globals():
        error_label.configure(text="⚠ No hay gráfico para exportar.", text_color="red")
        return

    ruta_guardado = filedialog.asksaveasfilename(
        defaultextension=".png",
        filetypes=[("Imagen PNG", "*.png")],
        title="Guardar gráfico como imagen"
    )

    if ruta_guardado:
        fig_actual.savefig(ruta_guardado)
        error_label.configure(text=f"✅ Gráfico guardado en:\n{ruta_guardado}", text_color="green")

ctk.set_appearance_mode("dark")  # "light", "dark", "system"
ctk.set_default_color_theme("blue")

# Crear ventana principal
ventana = ctk.CTk()
ventana.title("Graficador de Calificaciones")
ventana.geometry("1280x720")

# Crear marco desplazable
scrollable_frame = ctk.CTkScrollableFrame(ventana, label_text="Panel Principal")
scrollable_frame.pack(fill="both", expand=True, padx=20, pady=20)

# --- Widgets dentro del marco desplazable ---
label_instruccion = ctk.CTkLabel(scrollable_frame, text="Ingrese el nombre de la columna a graficar:")
label_instruccion.pack(pady=10)

entry_columna = ctk.CTkEntry(scrollable_frame, width=300, placeholder_text="Ejemplo: nombre_preparatoria")
entry_columna.pack(pady=5)

boton = ctk.CTkButton(scrollable_frame, text="Mostrar gráfico", command=lambda: mostrar_grafico(entry_columna.get().strip()))
boton.pack(pady=10)

error_label = ctk.CTkLabel(scrollable_frame, text="", text_color="red")
error_label.pack()

# --- Frame donde irá el gráfico ---
frame_grafico = ctk.CTkFrame(scrollable_frame)
frame_grafico.pack(pady=20, fill="both", expand=True)

# --- Botones extra ---
boton_exportar = ctk.CTkButton(scrollable_frame, text="Exportar gráfico como imagen", command=exportar_grafico)
boton_exportar.pack(pady=10)

boton_cargar_m = ctk.CTkButton(scrollable_frame, text="📘 Cargar diagnóstico (mañana)", command=main.cargar_excel_diag_mañana)
boton_cargar_m.pack(pady=10)

boton_cargar_t = ctk.CTkButton(scrollable_frame, text="📗 Cargar diagnóstico (tarde)", command=main.cargar_excel_diag_tarde)
boton_cargar_t.pack(pady=10)

boton_exportar_m = ctk.CTkButton(scrollable_frame, text="💾 Exportar diagnóstico (mañana)", command=lambda: main.exportar_excel_diag("mañana"))
boton_exportar_m.pack(pady=10)

boton_exportar_t = ctk.CTkButton(scrollable_frame, text="💾 Exportar diagnóstico (tarde)", command=lambda: main.exportar_excel_diag("tarde"))
boton_exportar_t.pack(pady=10)

boton_respuestas_m = ctk.CTkButton(scrollable_frame, text="📘 Cargar respuestas correctas (mañana)", 
                                   command=lambda: setattr(main, 'respuestas_m', main.cargar_excel()))
boton_respuestas_m.pack(pady=10)

boton_respuestas_t = ctk.CTkButton(scrollable_frame, text="📗 Cargar respuestas correctas (tarde)", 
                                   command=lambda: setattr(main, 'respuestas_t', main.cargar_excel()))
boton_respuestas_t.pack(pady=10)

boton_calificar = ctk.CTkButton(scrollable_frame, text="⚙️ Generar archivos calificados (1=Correcto, 0=Incorrecto)", 
                                command=lambda: main.generar_excels_calificados(
                                    main.base_datos_diag_m if 'base_datos_diag_m' in main.__dict__ else None,
                                    main.base_datos_diag_t if 'base_datos_diag_t' in main.__dict__ else None,
                                    main.respuestas_m if 'respuestas_m' in main.__dict__ else None,
                                    main.respuestas_t if 'respuestas_t' in main.__dict__ else None
                                ))
boton_calificar.pack(pady=20)


boton_cargar_analisis = ctk.CTkButton(
    scrollable_frame,
    text="📊 Cargar archivo combinado para análisis",
    command=main.cargar_excel_analitico
)
boton_cargar_analisis.pack(pady=10)

boton_analizar = ctk.CTkButton(
    scrollable_frame,
    text="📈 Abrir ventana de análisis",
    command=main.analizar_datos
)
boton_analizar.pack(pady=10)


# Ejecutar ventana
ventana.mainloop()


