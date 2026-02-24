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
ctk.set_default_color_theme("blue")

ventana = ctk.CTk()
ventana.title("Graficador de Calificaciones")
ventana.geometry("1280x720")

# Barra superior con botones (fila)
topbar = ctk.CTkFrame(ventana, height=50)
topbar.pack(side="top", fill="x")

# Frame contenedor donde se mostrarán las páginas
container = ctk.CTkFrame(ventana)
container.pack(fill="both", expand=True)

# -----------------------
# FUNCIONES
# -----------------------
frame_grafico = None

def mostrar_grafico(nombre_columna, error_label, frame_grafico):
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


def exportar_grafico(error_label_exportar):
    if 'fig_actual' not in globals():
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
    # Eliminar cualquier contenido previo
    for p in container.winfo_children():
        p.destroy()

    # Crear la página según el nombre
    if page_name == "Inicio":
        page = ctk.CTkFrame(container)
        ctk.CTkLabel(page, text="Página 1: Bienvenido", font=("Arial", 20)).pack(pady=20)
        ctk.CTkLabel(page, text="Este programa es calificar examenes, compilar estas calificaciones y graficar los resultados.\n"
                     "Por favor ingresar la base de datos base con los datos de los alumnos de la generación.").pack(pady=10)

        # --- Botón para cargar archivo principal ---
        boton_cargar_excel = ctk.CTkButton(
            page,
            text="📂 Cargar base de datos",
            command=actualizar_columnas
        )
        boton_cargar_excel.pack(pady=10)
        
        page.pack(fill="both", expand=True)

    elif page_name == "Formato":
        page = ctk.CTkFrame(container)
        
        ctk.CTkLabel(page, text="Esta pagina es para corregir el formato en los resultados de FormScanner o Zipgrade\n"
                     "dependiendo del programa que se uso. Todos los archivos resultantes tienen que ser formateados\n"
                     "para que funcione el programa.").pack(pady=10)
        
        def cargar_formateador(tipo_form):
            main.cargar_formateador(tipo_form)

        # Botón para abrir formateador CSV a Excel
        boton_cargar_formateador_csv = ctk.CTkButton(
            page,
            text="📂 Abrir formateador de FormScanner",
            command=lambda: cargar_formateador(1)   
        )
        boton_cargar_formateador_csv.pack(pady=10)

        # Botón para abrir formateador Zipgrade a Excel
        boton_cargar_formateador_zipgrade = ctk.CTkButton(
            page,
            text="📂 Abrir formateador de Zipgrade",
            command=lambda: cargar_formateador(2)   
        )
        boton_cargar_formateador_zipgrade.pack(pady=10)

        page.pack(fill="both", expand=True)

    elif page_name == "Graficas":
        page = ctk.CTkFrame(container)
        
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
        
        frame_grafico.pack(pady=20, fill="both", expand=True)

        error_label_exportar = ctk.CTkLabel(page, text="", text_color="red")

        boton_exportar = ctk.CTkButton(page, text="Exportar gráfico como imagen", command=lambda: exportar_grafico(error_label_exportar))
        boton_exportar.pack()

        error_label_exportar.pack()

        page.pack(fill="both", expand=True)
        
    elif page_name == "Calificar":
        page = ctk.CTkFrame(container)
    
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
            text="⚙️ Generar archivos calificados\n(1=Correcto, 0=Incorrecto)",
            height=80,
            fg_color="#3B82F6",
            hover_color="#2563EB",
            command=generar_archivos_calificados
        )
        boton_calificar.grid(row=0, column=1, rowspan=2, padx=10, pady=20)

        btn_combinar_diag_final = ctk.CTkButton(page, text="Combinar diagnosticos con final", command=main.combinar_diag_con_final)
        btn_combinar_diag_final.pack(pady=10)

        btn_combinar_combinado_base = ctk.CTkButton(page, text="Combinar diagnostico-finales con base de datos", command=main.combinar_combinado_completo_con_base_datos)
        btn_combinar_combinado_base.pack(pady=10)

        page.pack(fill="both", expand=True)
    elif page_name == "Analisis":
        page = ctk.CTkFrame(container)
        
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

        ctk.CTkButton(frame_analisis, text="📊 Comparar reprobados (2 archivos)",
              command=main.comparar_reprobados).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📊 Comparar porcentajes aprobados (2 archivos)",
              command=main.comparar_porcentajes).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📊 Ver Promedios por carrera",
              command=main.comparar_promedio_final_por_carrera).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="📊 Comparar promedios",
              command=main.comparar_promedios_por_carrera_dos_archivos).pack(pady=8)

        ctk.CTkButton(frame_analisis, text="🔵🟠 Comparar promedios totales de 2 archivos",
              command=main.comparar_promedio_total).pack(pady=8)
    
        page.pack(fill="both", expand=True)



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

ventana.mainloop()