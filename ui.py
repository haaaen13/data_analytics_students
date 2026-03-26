import customtkinter as ctk
import tkinter as tk
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from formscanner_batch import run_formscanner_workflow
import main

import sys
import os




def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
# -----------------------
# CONFIGURACIÓN DE APARIENCIA
# -----------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme(resource_path("tema.json"))

ventana = ctk.CTk()
ventana.title("Graficador de Calificaciones")
ventana.geometry("1280x720")
ventana.iconbitmap(resource_path("huev.ico"))  

# Barra superior con botones (fila)
topbar = ctk.CTkFrame(ventana, height=50)
topbar.pack(side="top", fill="x")

# Frame contenedor donde se mostrarán las páginas
container = ctk.CTkFrame(ventana)
container.pack(fill="both", expand=True)

# -----------------------
# SCROLL HELPER
# -----------------------

def crear_pagina_scrollable(parent):
    """
    Crea un frame scrollable dentro del contenedor padre.
    Devuelve el frame interno donde se deben agregar los widgets.
    """
    # Canvas que contiene todo
    canvas = tk.Canvas(parent, bg=parent.cget("fg_color")[1] if hasattr(parent, "cget") else "#2b2b2b",
                       highlightthickness=0)
    scrollbar = ctk.CTkScrollbar(parent, orientation="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    # Frame interno donde van los widgets de la página
    inner_frame = ctk.CTkFrame(canvas)
    inner_frame_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    # Ajusta la región de scroll cuando cambia el tamaño del inner_frame
    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    # Ajusta el ancho del inner_frame al ancho del canvas
    def on_canvas_configure(event):
        canvas.itemconfig(inner_frame_id, width=event.width)

    inner_frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_configure)

    # Scroll con rueda del mouse (Windows y Linux)
    def on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_mousewheel_linux(event):
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

    canvas.bind_all("<MouseWheel>", on_mousewheel)
    canvas.bind_all("<Button-4>", on_mousewheel_linux)
    canvas.bind_all("<Button-5>", on_mousewheel_linux)

    return inner_frame


# -----------------------
# FUNCIONES
# -----------------------
frame_grafico = None
fig_actual = None
canvas_actual = None

def mostrar_grafico(nombre_columna, error_label, frame_grafico):
    global fig_actual, canvas_actual

    if main.base_datos is None:
        error_label.configure(text="⚠ Primero carga un archivo Excel.", text_color="red")
        return

    if nombre_columna not in main.base_datos.columns:
        error_label.configure(text=f"⚠ La columna '{nombre_columna}' no existe.", text_color="red")
        return
    else:
        error_label.configure(text="")

    if canvas_actual is not None:
        canvas_actual.get_tk_widget().destroy()
        canvas_actual = None
    if fig_actual is not None:
        plt.close(fig_actual)
        fig_actual = None

    conteo = main.base_datos[nombre_columna].value_counts().head(20)
    fig, ax = plt.subplots(figsize=(10, 6))
    barras = ax.bar(conteo.index.astype(str), conteo.values, color='skyblue')
    fig_actual = fig

    # ── Número exacto sobre cada barra ───────────────────────────────────────
    for barra, valor in zip(barras, conteo.values):
        ax.text(
            barra.get_x() + barra.get_width() / 2,
            barra.get_height() + 0.5,
            str(valor),
            ha="center", va="bottom", fontsize=9, fontweight="bold"
        )

    ax.set_title(f'Conteo de registros por "{nombre_columna}"')
    ax.set_xlabel(nombre_columna)
    ax.set_ylabel("Cantidad")
    ax.set_xticks(range(len(conteo.index)))
    ax.set_xticklabels(conteo.index, rotation=45, ha='right')
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()

    for widget in frame_grafico.winfo_children():
        widget.destroy()

    canvas_actual = FigureCanvasTkAgg(fig, master=frame_grafico)
    canvas_actual.draw()
    canvas_actual.get_tk_widget().pack()

def exportar_grafico(error_label_exportar):
    if fig_actual is None:
        error_label_exportar.configure(text="⚠ No hay gráfico para exportar.", text_color="red")
        return

    ruta_guardado = filedialog.asksaveasfilename(
        defaultextension=".png",
        filetypes=[("Imagen PNG", "*.png")],
        title="Guardar gráfico como imagen"
    )

    if ruta_guardado:
        fig_actual.savefig(ruta_guardado)
        error_label_exportar.configure(text=f"✅ Gráfico guardado en:\n{ruta_guardado}", text_color="green")

combo = None
error_label = None

def actualizar_columnas():
    """Carga el Excel y actualiza las opciones del combobox"""
    main.cargar_excel_base()
    
#Funcion para generar archivos calificados
def generar_archivos_calificados():
    try:
        main.generar_excels_calificados(
        main.base_datos_diag_m, 
        main.base_datos_diag_t,
        main.respuestas_m, 
        main.respuestas_t
    )
    except AttributeError as e:
        # Mostrar un mensaje de error
        messagebox.showerror("Error", f"No se selecciono los archivos necesarios para calificar")




# Función para mostrar una página
def show_page(page_name):
    global fig_actual, canvas_actual

    # Cierra cualquier figura activa al cambiar de pagina.
    if canvas_actual is not None:
        canvas_actual.get_tk_widget().destroy()
        canvas_actual = None
    if fig_actual is not None:
        plt.close(fig_actual)
        fig_actual = None
    # Eliminar cualquier contenido previo
    for p in container.winfo_children():
        p.destroy()

    # Crear la página según el nombre
    if page_name == "Inicio":
        # Frame raíz que ocupa todo el contenedor
        root_frame = ctk.CTkFrame(container)
        root_frame.pack(fill="both", expand=True)

        # Obtenemos el frame interno scrollable
        page = crear_pagina_scrollable(root_frame)

        ctk.CTkLabel(page, text="Bienvenido", font=("Arial", 20)).pack(pady=20)
        ctk.CTkLabel(page, text="Este programa es calificar examenes, compilar estas calificaciones y graficar los resultados.\n"
                     "Por favor ingresar la base de datos base con los datos de los alumnos de la generación.").pack(pady=10)

        # --- Botón para cargar archivo principal ---
        boton_cargar_excel = ctk.CTkButton(
            page,
            text="📂 Cargar base de datos",
            command=actualizar_columnas
        )
        boton_cargar_excel.pack(pady=10)

    elif page_name == "Formato":
        root_frame = ctk.CTkFrame(container)
        root_frame.pack(fill="both", expand=True)

        page = crear_pagina_scrollable(root_frame)
        
        ctk.CTkLabel(page, text="Esta pagina es para corregir el formato en los resultados de FormScanner o Zipgrade\n"
                     "dependiendo del programa que se uso. Todos los archivos resultantes tienen que ser formateados\n"
                     "para que funcione el programa.").pack(pady=10)
        
        def cargar_formateador(tipo_form):
            main.cargar_formateador(tipo_form)


        boton_cargar_formateador_csv = ctk.CTkButton(
            page,
            text="📂 FORMSCANNER",
            command=lambda: run_formscanner_workflow()   
        )
        boton_cargar_formateador_csv.pack(pady=10)

        # Botón para abrir formateador CSV a Excel
        boton_cargar_formateador_csv2 = ctk.CTkButton(
            page,
            text="📂 Abrir formateador de FormScanner",
            command=lambda: cargar_formateador(1)   
        )
        boton_cargar_formateador_csv2.pack(pady=10)

        # Botón para abrir formateador Zipgrade a Excel
        boton_cargar_formateador_zipgrade = ctk.CTkButton(
            page,
            text="📂 Abrir formateador de Zipgrade",
            command=lambda: cargar_formateador(2)   
        )
        boton_cargar_formateador_zipgrade.pack(pady=10)

    elif page_name == "Graficas":
        root_frame = ctk.CTkFrame(container)
        root_frame.pack(fill="both", expand=True)

        page = crear_pagina_scrollable(root_frame)
        
        # --- Descripcion de Graficas
        ctk.CTkLabel(page, text="En esta pagina se puede crear una grafica para cada columna de la base de datos base.\n" 
                    "Seleccione una columna.").pack(pady=10)

        # --- Combobox (vacío al inicio) ---
        combo = ttk.Combobox(page, values=[], state="readonly")
        combo.set("Seleccione una columna para graficar")
        combo.pack(pady=10)

        # Actualizar valores si hay base de datos cargada
        if main.base_datos is not None:
            combo["values"] = main.base_datos.columns.tolist()
            combo.set("Seleccione una columna para graficar")
        else:
            combo["values"] = []
            combo.set("Seleccione una columna para graficar")

        # --- Error Label ---
        error_label = ctk.CTkLabel(page, text="", text_color="red")
        error_label.pack()

        if main.base_datos is not None:
            error_label.configure(text="✅ Archivo cargado correctamente.", text_color="green")
        else:
            error_label.configure(text="⚠ No se cargó ningún archivo.", text_color="red")

        frame_grafico = ctk.CTkFrame(page)

        # --- Botones principales ---
        boton = ctk.CTkButton(page, text="Mostrar gráfico", command=lambda: mostrar_grafico(combo.get().strip(), error_label, frame_grafico))
        boton.pack(pady=10)

        boton_exportar = ctk.CTkButton(page, text="Exportar gráfico como imagen", command=lambda: exportar_grafico(error_label_exportar))
        boton_exportar.pack()
        
        frame_grafico.pack(pady=20, fill="both", expand=True)

        error_label_exportar = ctk.CTkLabel(page, text="", text_color="red")
        error_label_exportar.pack()
        
    elif page_name == "Calificar":
        root_frame = ctk.CTkFrame(container)
        root_frame.pack(fill="both", expand=True)

        page = crear_pagina_scrollable(root_frame)
    
        # --- Frame principal de calificación ---
        frame_calificacion = ctk.CTkFrame(page, corner_radius=10)
        frame_calificacion.pack(pady=30, fill="x", padx=20)

        titulo_calificacion = ctk.CTkLabel(
            frame_calificacion,
            text="📘 Sección de Calificación",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="lightblue"
        )
        titulo_calificacion.pack(pady=10)

        #Descripcion para mejor uso de la pagina
        ctk.CTkLabel(frame_calificacion, text="Aqui se califican y se combinan los diferentes archivos.\n"
                     "Se tiene que tener en cuenta que se califica por separado los examenes diagnostico y final.").pack(pady=10)
        ctk.CTkLabel(frame_calificacion, text="INGRESAR EXAMENES A CALIFICAR").pack(pady=10)

        # Sub-frame interno para organizar botones en grid
        frame_botones = ctk.CTkFrame(frame_calificacion, fg_color="transparent")
        frame_botones.pack(fill="x", padx=10, pady=10)

        # Configurar 3 columnas
        frame_botones.grid_columnconfigure(0, weight=1)
        frame_botones.grid_columnconfigure(1, weight=1)
        frame_botones.grid_columnconfigure(2, weight=1)

        # 📘 Izquierda: Cargar respuestas de alumnos
        btn_cargar_m = ctk.CTkButton(
            frame_botones, 
            text="📘 Cargar carpeta con examenes (mañana)", 
            command=main.cargar_excel_diag_mañana
        )
        btn_cargar_m.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        btn_cargar_t = ctk.CTkButton(
            frame_botones, 
            text="📗 Cargar carpeta con examenes (tarde)", 
            command=main.cargar_excel_diag_tarde
        )
        btn_cargar_t.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        # 📗 Derecha: Cargar respuestas correctas
        boton_respuestas_m = ctk.CTkButton(
            frame_botones, 
            text="📘 Cargar respuestas correctas (mañana)", 
            command=lambda: setattr(main, 'respuestas_m', main.cargar_excel_respuestas())
        )
        boton_respuestas_m.grid(row=0, column=2, padx=10, pady=10, sticky="e")

        boton_respuestas_t = ctk.CTkButton(
            frame_botones, 
            text="📗 Cargar respuestas correctas (tarde)", 
            command=lambda: setattr(main, 'respuestas_t', main.cargar_excel_respuestas())
        )
        boton_respuestas_t.grid(row=1, column=2, padx=10, pady=10, sticky="e")

        # ⚙️ Centro: Generar calificaciones
        boton_calificar = ctk.CTkButton(
            frame_botones,
            text="⚙️ Generar archivos calificados",
            height=80,
            fg_color="#C0421F",
            hover_color="#AD3A10",
            command=generar_archivos_calificados
        )
        boton_calificar.grid(row=0, column=1, rowspan=2, padx=10, pady=20)

        btn_combinar_diag_final = ctk.CTkButton(page, text="Combinar diagnosticos con final", command=main.combinar_diag_con_final)
        btn_combinar_diag_final.pack(pady=10)

        btn_combinar_combinado_base = ctk.CTkButton(page, text="Combinar diagnostico-finales con base de datos", command=main.combinar_combinado_completo_con_base_datos)
        btn_combinar_combinado_base.pack(pady=10)

    elif page_name == "Analisis":
        root_frame = ctk.CTkFrame(container)
        root_frame.pack(fill="both", expand=True)

        page = crear_pagina_scrollable(root_frame)
        
        frame_analisis = ctk.CTkFrame(page, corner_radius=10)
        frame_analisis.pack(pady=30, fill="x", padx=20)

        titulo_analisis = ctk.CTkLabel(
            frame_analisis,
            text="📊 Panel de Análisis y Comparativas",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="lightblue"
        )
        titulo_analisis.pack(pady=10)

        ctk.CTkLabel(frame_analisis, text="Esta pagina es para el analisis de datos, combinando los resultados de los examenes\n"
                     "con la información de la base de datos. Se pueden analizar patrones en el aprendizaje de los alumnos.").pack(pady=10)

        ctk.CTkButton(frame_analisis, text="📊 Cargar archivo combinado para análisis",
              command=main.cargar_excel_analitico).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📈 Gráficar porcentaje de aciertos por pregunta",
              command=lambda: main.analizar_datos("aciertos")).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📈 Gráficar porcentaje de errores por pregunta",
              command=lambda: main.analizar_datos("errores")).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="🔵🟠 Generar gráfica de dispersión de aciertos (2 archivos calificados)",
              command=lambda: main.analizar_datos2("aciertos")).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="🔵🟠 Generar gráfica de dispersión de errores (2 archivos calificados)",
              command=lambda: main.analizar_datos2("errores")).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📊 Comparar aprobados y reprobados (2 archivos)",
              command=main.comparar_reprobados).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📊 Comparar porcentajes aprobados y reprobados (2 archivos)",
              command=main.comparar_porcentajes).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📊 Ver Promedios por carrera",
              command=main.comparar_promedio_final_por_carrera).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📊 Comparar promedios",
              command=main.comparar_promedios_por_carrera_dos_archivos).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="🔵🟠 Comparar promedios totales de 2 archivos",
              command=main.comparar_promedio_total).pack(pady=8)
        
        ctk.CTkButton(frame_analisis, text="🔵🟠 ver prepas",
              command=main.comparar_promedio_por_modelo).pack(pady=8)



# Botones en la barra superior
btn1 = ctk.CTkButton(topbar, text="Inicio", command=lambda: show_page("Inicio"))
btn1.pack(side="left", padx=10, pady=5)

btn2 = ctk.CTkButton(topbar, text="Formato", command=lambda: show_page("Formato"))
btn2.pack(side="left", padx=10, pady=5)

btn3 = ctk.CTkButton(topbar, text="Graficas", command=lambda: show_page("Graficas"))
btn3.pack(side="left", padx=10, pady=5)

btn4 = ctk.CTkButton(topbar, text="Calificar", command=lambda: show_page("Calificar"))
btn4.pack(side="left", padx=10, pady=5)

btn5 = ctk.CTkButton(topbar, text="Analisis", command=lambda: show_page("Analisis"))
btn5.pack(side="left", padx=10, pady=5)

# Mostrar la primera página por defecto
show_page("Inicio")

def on_close():
    plt.close('all')
    ventana.destroy()

ventana.protocol("WM_DELETE_WINDOW", on_close)

ventana.mainloop()