import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import main


# -----------------------
# CONFIGURACIÓN DE APARIENCIA
# -----------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

ventana = ctk.CTk()
ventana.title("Graficador de Calificaciones")
ventana.geometry("1280x720")


# -----------------------
# FUNCIONES
# -----------------------
def mostrar_grafico(nombre_columna):
    global fig_actual

    if main.base_datos is None:
        error_label.configure(text="⚠ Primero carga un archivo Excel.", text_color="red")
        return

    if nombre_columna not in main.base_datos.columns:
        error_label.configure(text=f"⚠ La columna '{nombre_columna}' no existe.", text_color="red")
        return
    else:
        error_label.configure(text="")

    conteo = main.base_datos[nombre_columna].value_counts().head(20)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar(conteo.index.astype(str), conteo.values, color='skyblue')
    fig_actual = fig

    ax.set_title(f'Conteo de registros por "{nombre_columna}"')
    ax.set_xlabel(nombre_columna)
    ax.set_ylabel("Cantidad")
    ax.set_xticks(range(len(conteo.index)))
    ax.set_xticklabels(conteo.index, rotation=45, ha='right')
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()

    for widget in frame_grafico.winfo_children():
        widget.destroy()

    canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack()


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


def actualizar_columnas():
    """Carga el Excel y actualiza las opciones del combobox"""
    main.cargar_excel_base()
    if main.base_datos is not None:
        combo["values"] = main.base_datos.columns.tolist()
        combo.set("Seleccione una columna para graficar")
        error_label.configure(text="✅ Archivo cargado correctamente.", text_color="green")
    else:
        combo["values"] = []
        combo.set("Seleccione una columna para graficar")
        error_label.configure(text="⚠ No se cargó ningún archivo.", text_color="red")


# -----------------------
# INTERFAZ
# -----------------------
scrollable_frame = ctk.CTkScrollableFrame(ventana, label_text="Panel Principal")
scrollable_frame.pack(fill="both", expand=True, padx=20, pady=20)

# --- Botón para cargar archivo principal ---
boton_cargar_excel = ctk.CTkButton(
    scrollable_frame,
    text="📂 Cargar archivo para graficar",
    command=actualizar_columnas
)
boton_cargar_excel.pack(pady=10)







# --- Combobox (vacío al inicio) ---
combo = ttk.Combobox(scrollable_frame, values=[])
combo.set("Seleccione una columna para graficar")
combo.pack(pady=10)

# --- Botones principales ---
boton = ctk.CTkButton(scrollable_frame, text="Mostrar gráfico", command=lambda: mostrar_grafico(combo.get().strip()))
boton.pack(pady=10)

error_label = ctk.CTkLabel(scrollable_frame, text="", text_color="red")
error_label.pack()

frame_grafico = ctk.CTkFrame(scrollable_frame)
frame_grafico.pack(pady=20, fill="both", expand=True)

boton_exportar = ctk.CTkButton(scrollable_frame, text="Exportar gráfico como imagen", command=exportar_grafico)
boton_exportar.pack(pady=10)

# --- Botones extra (funciones de main) ---
ctk.CTkButton(scrollable_frame, text="📘 Cargar carpeta con respuestas (mañana)",
              command=main.cargar_excel_diag_mañana).pack(pady=10)

ctk.CTkButton(scrollable_frame, text="📗 Cargar carpeta con respuestas (tarde)",
              command=main.cargar_excel_diag_tarde).pack(pady=10)


boton_respuestas_m = ctk.CTkButton(scrollable_frame, text="📘 Cargar respuestas correctas (mañana)", 
                                   command=lambda: setattr(main, 'respuestas_m', main.cargar_excel_respuestas()))
boton_respuestas_m.pack(pady=10)

boton_respuestas_t = ctk.CTkButton(scrollable_frame, text="📗 Cargar respuestas correctas (tarde)", 
                                   command=lambda: setattr(main, 'respuestas_t', main.cargar_excel_respuestas()))
boton_respuestas_t.pack(pady=10)


ctk.CTkButton(scrollable_frame, text="⚙️ Generar archivos calificados (1=Correcto, 0=Incorrecto)",
              command=lambda: main.generar_excels_calificados(
                  main.base_datos_diag_m, main.base_datos_diag_t,
                  main.respuestas_m, main.respuestas_t)).pack(pady=20)

ctk.CTkButton(scrollable_frame, text="📊 Cargar archivo combinado para análisis",
              command=main.cargar_excel_analitico).pack(pady=10)

ctk.CTkButton(scrollable_frame, text="📈 Abrir ventana de análisis",
              command=main.analizar_datos).pack(pady=10)

ctk.CTkButton(scrollable_frame, text="📈 Abrir ventana de análisis de errores por pregunta",
              command=main.analizar_datos_errores).pack(pady=10)


ctk.CTkButton(scrollable_frame, text="🔵🟠 Generar gráfica de dispersión (2 archivos calificados)",
              command=main.analizar_datos2).pack(pady=10)

ctk.CTkButton(scrollable_frame, text="📊 Comparar reprobados (2 archivos)",
              command=main.comparar_reprobados).pack(pady=10)

ctk.CTkButton(scrollable_frame, text="📊 Comparar aprobados y reprobados por carrera",
              command=main.comparar_por_carrera_unico).pack(pady=10)




ctk.CTkButton(scrollable_frame, text="🔵🟠 Generar gráfica de dispersión de errores (2 archivos calificados)",
              command=main.analizar_datos2_errores).pack(pady=10)



btn_combinar_diag_final = ttk.Button(scrollable_frame, text="Combinar archivos clave", command=main.combinar_diag_con_final)
btn_combinar_diag_final.pack(pady=10)

btn_combinar_combinado_base = ttk.Button(scrollable_frame, text="Combinar diagnostico-finales con base de datos", command=main.combinar_combinado_completo_con_base_datos)
btn_combinar_combinado_base.pack(pady=10)

# -----------------------
# INICIAR PROGRAMA
# -----------------------
ventana.mainloop()
