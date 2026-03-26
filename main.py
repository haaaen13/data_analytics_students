from tkinter import simpledialog
import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog
import seaborn as sns
from tkinter import ttk
import os



#base_datos=pd.read_excel(r"C:\Users\Ramir\Downloads\Base_alumnos 3.xlsx")
#xlsx_cpa_mat_diag=pd.read_excel(r"C:\Users\Ramir\Downloads\2025 CPA MAT COMP diag.xlsx")
#xlsx_cpa_mat_final=pd.read_excel(r"C:\Users\Ramir\Downloads\2025 CPA mat final.xlsx")




base_datos = None
xlsx_cpa_mat_diag = None

def cargar_excel_base():
    """Permite al usuario seleccionar un Excel y lo carga en base_datos"""
    global base_datos
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel para graficar",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not ruta:
        messagebox.showinfo("Cancelado", "No se seleccionó ningún archivo.")
        return

    try:
        base_datos = pd.read_excel(ruta)
        messagebox.showinfo("Éxito", f"Archivo cargado correctamente:\n{ruta}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")


# -----------------------
# SCROLL HELPER para ventanas Toplevel
# -----------------------

def crear_frame_scrollable_toplevel(toplevel_window):
    """
    Crea un frame scrollable dentro de una ventana CTkToplevel/Toplevel.
    Devuelve el frame interno donde se agregan los widgets.
    Soporta scroll con rueda del mouse.
    """
    canvas = tk.Canvas(toplevel_window, highlightthickness=0)
    scrollbar_v = ctk.CTkScrollbar(toplevel_window, orientation="vertical", command=canvas.yview)
    scrollbar_h = ctk.CTkScrollbar(toplevel_window, orientation="horizontal", command=canvas.xview)

    canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

    scrollbar_v.pack(side="right", fill="y")
    scrollbar_h.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)

    inner_frame = ctk.CTkFrame(canvas)
    inner_frame_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_configure(event):
        # Solo expandir el ancho si el frame es más pequeño que el canvas
        if inner_frame.winfo_reqwidth() < event.width:
            canvas.itemconfig(inner_frame_id, width=event.width)

    inner_frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_configure)

    # Scroll con rueda del mouse
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


#Excel Diagnostico
def seleccionar_carpeta_diag(turno):
    """Abre el explorador de carpetas y devuelve la ruta seleccionada."""
    messagebox.showinfo(
        "Aviso",
        f"Abrir folder de tablas de respuestas del turno {turno}"
    )

    ruta_carpeta = filedialog.askdirectory(
        title=f"Selecciona la carpeta de respuestas del examen ({turno})"
    )

    if not ruta_carpeta:
        print(f"No se seleccionó ninguna carpeta para el turno {turno}.")
        return None
    
    return ruta_carpeta


def combinar_archivos_excel_diag(ruta_carpeta, turno):
    """Carga y combina todos los archivos Excel dentro de la carpeta dada."""
    global base_datos_diag_m, base_datos_diag_t

    if not ruta_carpeta:
        messagebox.showinfo("Cancelado", "No se seleccionó ninguna carpeta.")
        return None

    dataframesdiag = []

    for archivo in os.listdir(ruta_carpeta):
        if archivo.endswith(".xlsx") or archivo.endswith(".xls"):
            ruta_completa = os.path.join(ruta_carpeta, archivo)
            try:
                df = pd.read_excel(ruta_completa)
                df["archivo_origen"] = archivo
                df["turno"] = turno
                dataframesdiag.append(df)
                print(f"[{turno}] Cargado: {archivo}")
            except Exception as e:
                print(f"[{turno}] Error al cargar {archivo}: {e}")

    if dataframesdiag:
        combinado = pd.concat(dataframesdiag, ignore_index=True)
        print(f"[{turno}] Archivos combinados exitosamente.")
        print(combinado.head())

        if turno.lower() == "mañana":
            base_datos_diag_m = combinado
        else:
            base_datos_diag_t = combinado

        return combinado
    else:
        messagebox.showerror("Error", "No se seleccionó ubicación para guardar el archivo.")
        return None


def cargar_excel_diag_mañana():
    """Carga archivos de respuestas del turno mañana."""
    ruta = seleccionar_carpeta_diag("mañana")
    return combinar_archivos_excel_diag(ruta, "mañana")


def cargar_excel_diag_tarde():
    """Carga archivos de respuestas del turno tarde."""
    ruta = seleccionar_carpeta_diag("tarde")
    return combinar_archivos_excel_diag(ruta, "tarde")


respuestas_m = None
respuestas_t = None


def cargar_respuestas(turno):
    """Carga archivo de respuestas correctas del turno."""
    global respuestas_m, respuestas_t

    messagebox.showinfo(
        "Abrir archivo de respuestas correctas",
        f"Selecciona el archivo de respuestas correctas del turno {turno}"
    )

    ruta = filedialog.askopenfilename(
        title=f"Selecciona archivo de respuestas correctas ({turno})",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not ruta:
        print(f"No se seleccionó archivo de respuestas para el turno {turno}")
        return None

    df = pd.read_excel(ruta)
    print(f"✅ Archivo de respuestas cargado ({turno}): {ruta}")
    print(df.head())

    if turno.lower() == "mañana":
        respuestas_m = df
    else:
        respuestas_t = df

    return df


def cargar_excel_respuestas():
    global respuestas

    messagebox.showinfo(
        "Aviso",
        "Abrir archivo con las respuestas correctas"
    )

    ruta_archivo = filedialog.askopenfilename(
        title="Selecciona un archivo Excel base para alumnos",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    
    if ruta_archivo:
        respuestas = pd.read_excel(ruta_archivo)
        return respuestas
    else:
        messagebox.showerror("Error", f"No se seleccionó ningún archivo.")
        return None


def comparar_respuestas_y_calificar(df_diag, df_respuestas, turno):
    """
    Compara las respuestas del excel con las respuestas correctas.
    Devuelve un DataFrame con columnas originales y nuevas columnas de aciertos (1/0).
    """
    if df_diag is None or df_respuestas is None:
        messagebox.showerror("Error", f"Faltan archivos de respuestas o respuestas para {turno}.")
        return None

    preguntas = [f"P{i}" for i in range(1, 21)]
    for col in preguntas:
        if col not in df_diag.columns:
            df_diag[col] = None
        if col not in df_respuestas.columns:
            df_respuestas[col] = None

    respuestas_correctas = df_respuestas.iloc[0][preguntas]

    for col in preguntas:
        col_correcta = f"{col}_correcta"
        df_diag[col_correcta] = df_diag[col].apply(
            lambda x: 1 if pd.notna(x) and str(x).strip().upper() == str(respuestas_correctas[col]).strip().upper() else 0
        )

    print(f"✅ Respuestas comparadas para el turno {turno}")
    return df_diag


def generar_excels_calificados(df_m, df_t, df_resp_m, df_resp_t):
    """
    Genera 3 archivos Excel:
      - Calificados para turno mañana
      - Calificados para turno tarde
      - Combinado de ambos
    """
    if df_m is None and df_t is None:
        messagebox.showerror("Error", "No hay datos cargados para calificar.")
        return

    df_m_cal = comparar_respuestas_y_calificar(df_m, df_resp_m, "mañana") if df_m is not None else None
    df_t_cal = comparar_respuestas_y_calificar(df_t, df_resp_t, "tarde") if df_t is not None else None

    def calcular_calificacion(df):
        preguntas_correctas = [f"P{i}_correcta" for i in range(1, 21)]
        df["CALIF DIAG"] = df[preguntas_correctas].sum(axis=1) * 5
        return df

    if df_m_cal is not None:
        df_m_cal = calcular_calificacion(df_m_cal)
    if df_t_cal is not None:
        df_t_cal = calcular_calificacion(df_t_cal)

    ruta_m = None
    ruta_t = None
    ruta_c = None

    ruta_guardado = filedialog.askdirectory(title="Selecciona la carpeta donde guardar los archivos calificados")
    if not ruta_guardado:
        messagebox.showwarning("Aviso", "No se seleccionó carpeta para guardar los resultados.")
        return

    if df_m_cal is not None:
        ruta_m = filedialog.asksaveasfilename(
            title="Guardar Examen Matutino",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialfile="Examen_Matutino_Calificado.xlsx"
        )

    if ruta_m:
        try:
            df_m_cal.to_excel(ruta_m, index=False)
            messagebox.showinfo("Guardado exitoso", f"El archivo se guardó correctamente en:\n{ruta_m}")
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo.\n\n{e}")

    if df_t_cal is not None:
        ruta_t = filedialog.asksaveasfilename(
            title="Guardar Examen Vespertino",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialfile="Examen_Vespertino_Calificado.xlsx"
        )

    if ruta_t:
        try:
            df_t_cal.to_excel(ruta_t, index=False)
            messagebox.showinfo("Guardado exitoso", f"El archivo se guardó correctamente en:\n{ruta_t}")
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo.\n\n{e}")

    combinados = []
    if df_m_cal is not None:
        combinados.append(df_m_cal)
    if df_t_cal is not None:
        combinados.append(df_t_cal)

    if combinados:
        df_total = pd.concat(combinados, ignore_index=True)
        ruta_c = filedialog.asksaveasfilename(
            title="Guardar Examen Combinado",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialfile="Examen_Combinado.xlsx"
        )

    if ruta_c:
        try:
            df_total.to_excel(ruta_c, index=False)
            messagebox.showinfo("Guardado exitoso", f"El archivo se guardó correctamente en:\n{ruta_c}")
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo.\n\n{e}")

        return df_total

    return None


excel_combinado = None

def cargar_excel_analitico():
    """Permite seleccionar el archivo Excel combinado generado anteriormente."""
    global excel_combinado

    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo combinado para análisis",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not ruta:
        messagebox.showwarning("Aviso", "No se seleccionó ningún archivo para análisis.")
        return None

    excel_combinado = pd.read_excel(ruta)
    messagebox.showinfo("Éxito", f"Archivo cargado correctamente:\n{ruta}")
    print(f"📊 Archivo combinado cargado para análisis: {ruta}")
    return excel_combinado


# ---------- analizar_datos (un solo archivo, con filtros) ----------
def analizar_datos(estado='aciertos'):
    """Analiza un solo archivo con opción de filtrar por carrera y grupo."""
    global excel_combinado
    if excel_combinado is None:
        messagebox.showerror("Error", "Primero carga el archivo combinado para análisis.")
        return

    ventana_analisis = tk.Toplevel()
    ventana_analisis.title("Analíticas del Examen")
    ventana_analisis.geometry("1300x850")

    frame_scroll = ctk.CTkScrollableFrame(ventana_analisis, label_text="Resultados Analíticos")
    frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(frame_scroll, text="Filtrar por carrera (opcional):").pack(pady=5)
    if "carrera" in excel_combinado.columns:
        carreras = excel_combinado["carrera"].dropna().astype(str).str.strip().unique().tolist()
        carreras = sorted(carreras)
    else:
        carreras = []
    combo_carrera = ttk.Combobox(frame_scroll, values=["(Todas)"] + carreras, state="readonly")
    combo_carrera.set("(Todas)")
    combo_carrera.pack(pady=5)

    ctk.CTkLabel(frame_scroll, text="Filtrar por grupo (opcional):").pack(pady=5)
    combo_grupo = ttk.Combobox(frame_scroll, values=["(Todos)"], state="readonly")
    combo_grupo.set("(Todos)")
    combo_grupo.pack(pady=5)

    def actualizar_grupos(event=None):
        carrera_sel = combo_carrera.get()
        if carrera_sel == "(Todas)" or "ppgrupo" not in excel_combinado.columns:
            combo_grupo["values"] = ["(Todos)"]
            combo_grupo.set("(Todos)")
        else:
            df_filtrado = excel_combinado.copy()
            df_filtrado = df_filtrado[df_filtrado["carrera"].astype(str).str.strip() == str(carrera_sel).strip()]
            grupos_filtrados = df_filtrado["ppgrupo"].dropna().astype(str).str.strip().unique().tolist()
            grupos_filtrados = sorted(grupos_filtrados)
            combo_grupo["values"] = ["(Todos)"] + grupos_filtrados if grupos_filtrados else ["(Todos)"]
            combo_grupo.set("(Todos)")

    combo_carrera.bind("<<ComboboxSelected>>", actualizar_grupos)

    frame_graficas = ctk.CTkFrame(frame_scroll)
    frame_graficas.pack(fill="both", expand=True, pady=20)

    preguntas_esperadas = [f"P{i}_correcta" for i in range(1, 21)]
    x = range(1, 21)

    figs_generadas = {}

    def generar_graficas():
        nonlocal figs_generadas
        for fig in figs_generadas.values():
            plt.close(fig)
        figs_generadas = {}

        for widget in frame_graficas.winfo_children():
            widget.destroy()

        df = excel_combinado.copy()

        carrera_sel = combo_carrera.get()
        grupo_sel = combo_grupo.get()

        if carrera_sel != "(Todas)" and "carrera" in df.columns:
            df = df[df["carrera"].astype(str).str.strip() == str(carrera_sel).strip()]
        if grupo_sel != "(Todos)" and "ppgrupo" in df.columns:
            df = df[df["ppgrupo"].astype(str).str.strip() == str(grupo_sel).strip()]

        for col in preguntas_esperadas:
            if col not in df.columns:
                df[col] = float("nan")

        titulo_modo = "Aciertos" if estado == "aciertos" else "Errores"

        if estado == "aciertos":
            aciertos = df[preguntas_esperadas].mean().reindex(preguntas_esperadas)
        else:
            errores = 1 - df[preguntas_esperadas].mean()
            errores = errores.reindex(preguntas_esperadas)

        fig1 = plt.figure(figsize=(10, 6))

        if estado == 'aciertos':
            plt.scatter(x, aciertos.values * 100, color="skyblue")
        else:
            plt.scatter(x, errores.values * 100, color="skyblue")
        plt.title(f"Porcentaje de {titulo_modo} por pregunta")
        plt.xlabel("Pregunta")
        plt.ylabel(f"% de {titulo_modo}")
        plt.xticks(x, [f"P{i}" for i in x])
        plt.xlim(0.5, 20.5)
        plt.ylim(0, 100)
        plt.grid(True)
        plt.tight_layout()

        fig1 = plt.gcf()

        canvas1 = FigureCanvasTkAgg(fig1, master=frame_graficas)
        canvas1.draw()
        canvas1.get_tk_widget().pack(pady=20)

        figs_generadas[titulo_modo] = fig1

        if "carrera" in df.columns and "CALIFICACION DIAG" in df.columns:
            prom = df.groupby(df["carrera"].astype(str).str.strip())["CALIFICACION DIAG"].mean().sort_values(ascending=False)

            fig2 = plt.figure(figsize=(10, 6))
            prom.plot(kind="bar", color="skyblue")
            plt.title("Promedio de calificación diagnóstica por carrera")
            plt.xlabel("Carrera")
            plt.ylabel("Promedio")
            plt.xticks(rotation=45, ha="right")
            plt.ylim(0, 100)
            plt.tight_layout()

            canvas2 = FigureCanvasTkAgg(fig2, master=frame_graficas)
            canvas2.draw()
            canvas2.get_tk_widget().pack(pady=20)

            figs_generadas["promedios"] = fig2

    def guardar_graficas():
        if not figs_generadas:
            messagebox.showwarning("Aviso", "Primero genera las gráficas.")
            return

        for nombre, fig in figs_generadas.items():
            ruta = filedialog.asksaveasfilename(
                title=f"Guardar gráfica: {nombre}",
                defaultextension=".png",
                filetypes=[("Imagen PNG", "*.png")],
                initialfile=f"{nombre}.png"
            )

            if not ruta:
                continue

            fig.savefig(ruta, dpi=300, bbox_inches="tight")

    boton_filtrar = ctk.CTkButton(frame_scroll, text="🔍 Aplicar filtros y generar gráficas", command=generar_graficas)
    boton_filtrar.pack(pady=10)

    boton_guardar = ctk.CTkButton(frame_scroll, text="💾 Guardar gráficas en PNG", command=guardar_graficas)
    boton_guardar.pack(pady=10)

    def on_close():
        for fig in figs_generadas.values():
            plt.close(fig)
        ventana_analisis.destroy()

    ventana_analisis.protocol("WM_DELETE_WINDOW", on_close)

    generar_graficas()


# ---------- analizar_datos2 (comparativo entre 2 archivos) ----------
def analizar_datos2(modo="aciertos"):
    """Comparador de dos archivos Excel por grupo (ppgrupo)."""

    ruta1 = filedialog.askopenfilename(
        title="Selecciona el primer archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta1:
        messagebox.showinfo("Cancelado", "No se seleccionó el primer archivo.")
        return

    ruta2 = filedialog.askopenfilename(
        title="Selecciona el segundo archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta2:
        messagebox.showinfo("Cancelado", "No se seleccionó el segundo archivo.")
        return

    try:
        excel1 = pd.read_excel(ruta1)
        excel2 = pd.read_excel(ruta2)
    except Exception as e:
        messagebox.showerror("Error al leer Excel", str(e))
        return

    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    ventana_comp = tk.Toplevel()
    titulo_modo = "Aciertos" if modo == "aciertos" else "Errores"
    ventana_comp.title(f"Comparador de Resultados— {titulo_modo}")
    ventana_comp.geometry("1300x850")

    frame_scroll = ctk.CTkScrollableFrame(
        ventana_comp,
        label_text="Comparación de grupos o archivos"
    )
    frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(frame_scroll, text=f"Grupo en {nombre1}:").pack(pady=5)
    grupos1 = sorted(excel1["ppgrupo"].dropna().astype(str).str.strip().unique().tolist()) \
              if "ppgrupo" in excel1.columns else []
    combo_grupo1 = ttk.Combobox(frame_scroll, values=["(Todos)"] + grupos1)
    combo_grupo1.set("(Todos)")
    combo_grupo1.pack(pady=5)

    ctk.CTkLabel(frame_scroll, text=f"Grupo en {nombre2}:").pack(pady=5)
    grupos2 = sorted(excel2["ppgrupo"].dropna().astype(str).str.strip().unique().tolist()) \
              if "ppgrupo" in excel2.columns else []
    combo_grupo2 = ttk.Combobox(frame_scroll, values=["(Todos)"] + grupos2)
    combo_grupo2.set("(Todos)")
    combo_grupo2.pack(pady=5)

    frame_graficas = ctk.CTkFrame(frame_scroll)
    frame_graficas.pack(fill="both", expand=True, pady=20)

    preguntas = [f"P{i}_correcta" for i in range(1, 21)]
    x = range(1, 21)
    fig_comparacion_actual = None
    fig_aprob_actual = None

    def generar_comparacion():
        nonlocal fig_comparacion_actual, fig_aprob_actual

        if fig_comparacion_actual is not None:
            plt.close(fig_comparacion_actual)
            fig_comparacion_actual = None
        if fig_aprob_actual is not None:
            plt.close(fig_aprob_actual)
            fig_aprob_actual = None

        for w in frame_graficas.winfo_children():
            w.destroy()

        grupo1 = combo_grupo1.get()
        grupo2 = combo_grupo2.get()

        df1 = excel1.copy()
        df2 = excel2.copy()

        if grupo1 != "(Todos)" and "ppgrupo" in df1.columns:
            df1 = df1[df1["ppgrupo"].astype(str).str.strip() == grupo1]
        if grupo2 != "(Todos)" and "ppgrupo" in df2.columns:
            df2 = df2[df2["ppgrupo"].astype(str).str.strip() == grupo2]

        for df in [df1, df2]:
            for col in preguntas:
                if col not in df.columns:
                    df[col] = float("nan")

        if modo == "aciertos":
            e1 = df1[preguntas].mean() * 100
            e2 = df2[preguntas].mean() * 100
            titulo = "Comparación de aciertos por pregunta"
            ylabel = "% de aciertos"
        else:
            e1 = (1 - df1[preguntas].mean()) * 100
            e2 = (1 - df2[preguntas].mean()) * 100
            titulo = "Comparación de errores por pregunta"
            ylabel = "% de errores"

        plt.figure(figsize=(10, 6))
        plt.plot(x, e1.values, marker="o", linestyle="--", label=f"{nombre1} - {grupo1}")
        plt.plot(x, e2.values, marker="o", linestyle="--", label=f"{nombre2} - {grupo2}")

        plt.title(titulo)
        plt.xlabel("Pregunta")
        plt.ylabel(ylabel)
        plt.xticks(x, [f"P{i}" for i in x])
        plt.ylim(0, 100)
        plt.grid(True)
        plt.legend()
        plt.tight_layout()

        fig_comparacion = plt.gcf()
        fig_comparacion_actual = fig_comparacion

        canvas1 = FigureCanvasTkAgg(fig_comparacion, master=frame_graficas)
        canvas1.draw()
        canvas1.get_tk_widget().pack(pady=20)

        if "CALIFICACION DIAG" in df1.columns and "CALIFICACION DIAG" in df2.columns:
            aprob1 = (df1["CALIFICACION DIAG"] >= 60).sum()
            reprob1 = (df1["CALIFICACION DIAG"] < 60).sum()
            aprob2 = (df2["CALIFICACION DIAG"] >= 60).sum()
            reprob2 = (df2["CALIFICACION DIAG"] < 60).sum()

            df_aprob = pd.DataFrame({
                "Archivo": [nombre1, nombre1, nombre2, nombre2],
                "Resultado": ["Aprobado", "Reprobado", "Aprobado", "Reprobado"],
                "Cantidad": [aprob1, reprob1, aprob2, reprob2]
            })

            plt.figure(figsize=(7, 6))
            sns.barplot(data=df_aprob, x="Archivo", y="Cantidad", hue="Resultado", palette="pastel")
            plt.title("Comparación de aprobados y reprobados")
            plt.ylim(0, 100)
            plt.tight_layout()

            fig_aprob = plt.gcf()
            fig_aprob_actual = fig_aprob

            canvas2 = FigureCanvasTkAgg(fig_aprob, master=frame_graficas)
            canvas2.draw()
            canvas2.get_tk_widget().pack(pady=20)

        def guardar_png():
            archivo = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("Imagen PNG", "*.png")],
                title="Guardar gráfica como PNG"
            )
            if archivo:
                if fig_comparacion_actual is not None:
                    fig_comparacion_actual.savefig(archivo, dpi=300)
                messagebox.showinfo("Guardado", f"Gráfica guardada como:\n{archivo}")

        ctk.CTkButton(
            frame_graficas,
            text="💾 Guardar gráfica como PNG",
            command=guardar_png
        ).pack(pady=10)

    def resetear_comparacion():
        combo_grupo1.set("(Todos)")
        combo_grupo2.set("(Todos)")
        generar_comparacion()

    botones_frame = ctk.CTkFrame(frame_scroll)
    botones_frame.pack(pady=10)

    ctk.CTkButton(
        botones_frame, text="📊 Generar comparación", command=generar_comparacion
    ).grid(row=0, column=0, padx=10)

    ctk.CTkButton(
        botones_frame, text="🔄 Restablecer comparación general",
        fg_color="gray", command=resetear_comparacion
    ).grid(row=0, column=1, padx=10)

    def on_close():
        if fig_comparacion_actual is not None:
            plt.close(fig_comparacion_actual)
        if fig_aprob_actual is not None:
            plt.close(fig_aprob_actual)
        ventana_comp.destroy()

    ventana_comp.protocol("WM_DELETE_WINDOW", on_close)
    generar_comparacion()


# ---------- comparar_reprobados ----------
def comparar_reprobados():
    """Compara aprobados y reprobados entre dos archivos Excel con scroll."""

    ruta1 = filedialog.askopenfilename(
        title="Selecciona el primer archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta1:
        messagebox.showinfo("Cancelado", "No se seleccionó el primer archivo.")
        return

    ruta2 = filedialog.askopenfilename(
        title="Selecciona el segundo archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta2:
        messagebox.showinfo("Cancelado", "No se seleccionó el segundo archivo.")
        return

    try:
        df1 = pd.read_excel(ruta1)
        df2 = pd.read_excel(ruta2)
    except Exception as e:
        messagebox.showerror("Error al leer Excel", str(e))
        return

    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    preguntas = [f"P{i}_correcta" for i in range(1, 21)]
    for df in [df1, df2]:
        for col in preguntas:
            if col not in df.columns:
                df[col] = float("nan")

    df1["total_correctas"] = df1[preguntas].sum(axis=1)
    df2["total_correctas"] = df2[preguntas].sum(axis=1)

    reprobados1 = (df1["total_correctas"] < 14).sum()
    reprobados2 = (df2["total_correctas"] < 14).sum()

    total1 = len(df1)
    total2 = len(df2)

    aprobados1 = total1 - reprobados1
    aprobados2 = total2 - reprobados2

    resumen = pd.DataFrame({
        "Archivo": [nombre1, nombre2],
        "Aprobados": [aprobados1, aprobados2],
        "Reprobados": [reprobados1, reprobados2],
        "Total alumnos": [total1, total2]
    })

    # --- Ventana con scroll ---
    ventana_comp = ctk.CTkToplevel()
    ventana_comp.title("Comparación de Aprobados y Reprobados")
    ventana_comp.geometry("950x750")

    frame = crear_frame_scrollable_toplevel(ventana_comp)

    # --- Gráfica ---
    fig, ax = plt.subplots(figsize=(8, 6))
    x = range(len(resumen))
    ancho = 0.35

    ax.bar([i - ancho/2 for i in x], resumen["Aprobados"], width=ancho, label="Aprobados", color="royalblue")
    ax.bar([i + ancho/2 for i in x], resumen["Reprobados"], width=ancho, label="Reprobados", color="salmon")

    for i, (ap, rp) in enumerate(zip(resumen["Aprobados"], resumen["Reprobados"])):
        ax.text(i - ancho/2, ap + 0.5, str(ap), ha='center', va='bottom', fontsize=9, color='blue')
        ax.text(i + ancho/2, rp + 0.5, str(rp), ha='center', va='bottom', fontsize=9, color='red')

    ax.set_xticks(x)
    ax.set_xticklabels(resumen["Archivo"], rotation=15)
    ax.set_ylabel("Número de alumnos")
    ax.set_title("Comparación de Aprobados (azul) y Reprobados (rojo)")
    ax.set_ylim(0, max(resumen[["Aprobados", "Reprobados"]].max().max() * 1.2, 10))
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.6)
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10)

    # --- Tabla ---
    tabla_frame = ctk.CTkFrame(frame)
    tabla_frame.pack(pady=10)

    tabla = ttk.Treeview(tabla_frame, columns=("Aprobados", "Reprobados", "Total"), show="headings", height=3)
    tabla.heading("Aprobados", text="Aprobados")
    tabla.heading("Reprobados", text="Reprobados")
    tabla.heading("Total", text="Total alumnos")
    tabla.column("Aprobados", anchor="center", width=150)
    tabla.column("Reprobados", anchor="center", width=150)
    tabla.column("Total", anchor="center", width=150)

    for i in range(len(resumen)):
        tabla.insert("", "end",
                     values=(resumen["Aprobados"][i], resumen["Reprobados"][i], resumen["Total alumnos"][i]),
                     text=resumen["Archivo"][i])
    tabla.pack(pady=10)

    def guardar_grafica():
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("Imagen PNG", "*.png")],
            title="Guardar gráfica como imagen"
        )
        if ruta_guardado:
            fig.savefig(ruta_guardado, dpi=300)
            messagebox.showinfo("Guardado", f"✅ Gráfica guardada como:\n{ruta_guardado}")

    botones_frame = ctk.CTkFrame(frame)
    botones_frame.pack(pady=15)
    ctk.CTkButton(botones_frame, text="💾 Guardar gráfica", command=guardar_grafica).grid(row=0, column=0, padx=10)

    def on_close():
        plt.close(fig)
        ventana_comp.destroy()

    ventana_comp.protocol("WM_DELETE_WINDOW", on_close)


# ---------- comparar_porcentajes ----------
def comparar_porcentajes():
    """Compara porcentajes de aprobados y reprobados entre dos archivos Excel con scroll."""

    ruta1 = filedialog.askopenfilename(
        title="Selecciona el primer archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta1:
        messagebox.showinfo("Cancelado", "No se seleccionó el primer archivo.")
        return

    ruta2 = filedialog.askopenfilename(
        title="Selecciona el segundo archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta2:
        messagebox.showinfo("Cancelado", "No se seleccionó el segundo archivo.")
        return

    try:
        df1 = pd.read_excel(ruta1)
        df2 = pd.read_excel(ruta2)
    except Exception as e:
        messagebox.showerror("Error al leer Excel", str(e))
        return

    preguntas = [f"P{i}_correcta" for i in range(1, 21)]
    for df in [df1, df2]:
        for col in preguntas:
            if col not in df.columns:
                df[col] = float("nan")

    df1["total_correctas"] = df1[preguntas].sum(axis=1)
    df2["total_correctas"] = df2[preguntas].sum(axis=1)

    reprobados1 = (df1["total_correctas"] < 14).sum()
    reprobados2 = (df2["total_correctas"] < 14).sum()

    total1 = len(df1)
    total2 = len(df2)

    aprobados1 = total1 - reprobados1
    aprobados2 = total2 - reprobados2

    porc_ap1 = (aprobados1 / total1) * 100
    porc_ap2 = (aprobados2 / total2) * 100
    porc_rp1 = (reprobados1 / total1) * 100
    porc_rp2 = (reprobados2 / total2) * 100

    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    resumen = pd.DataFrame({
        "Archivo": [nombre1, nombre2],
        "Aprobados (%)": [porc_ap1, porc_ap2],
        "Reprobados (%)": [porc_rp1, porc_rp2],
        "Total alumnos": [total1, total2]
    })

    # --- Ventana con scroll ---
    ventana = ctk.CTkToplevel()
    ventana.title("Comparación de Porcentajes")
    ventana.geometry("950x750")

    frame = crear_frame_scrollable_toplevel(ventana)

    # --- Gráfica ---
    fig, ax = plt.subplots(figsize=(8, 6))
    x = range(len(resumen))
    ancho = 0.5

    ax.bar([i - ancho/2 for i in x], resumen["Aprobados (%)"], width=ancho, label="Aprobados (%)")
    ax.bar([i + ancho/2 for i in x], resumen["Reprobados (%)"], width=ancho, label="Reprobados (%)")

    for i, (ap, rp) in enumerate(zip(resumen["Aprobados (%)"], resumen["Reprobados (%)"])):
        ax.text(i - ancho/2, ap + 1, f"{ap:.1f}%", ha='center', fontsize=9)
        ax.text(i + ancho/2, rp + 1, f"{rp:.1f}%", ha='center', fontsize=9)

    ax.set_xticks(list(x))
    ax.set_xticklabels(resumen["Archivo"], rotation=15)
    ax.set_ylabel("Porcentaje (%)")
    ax.set_title("Porcentajes de Aprobados y Reprobados")
    ax.set_ylim(0, 110)
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1), borderaxespad=0.)
    ax.grid(axis="y", linestyle="--", alpha=0.6)
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10)

    # --- Tabla ---
    tabla_frame = ctk.CTkFrame(frame)
    tabla_frame.pack(pady=10)

    tabla = ttk.Treeview(
        tabla_frame,
        columns=("Aprobados (%)", "Reprobados (%)", "Total"),
        show="headings",
        height=3
    )
    tabla.heading("Aprobados (%)", text="Aprobados (%)")
    tabla.heading("Reprobados (%)", text="Reprobados (%)")
    tabla.heading("Total", text="Total alumnos")
    tabla.column("Aprobados (%)", anchor="center", width=150)
    tabla.column("Reprobados (%)", anchor="center", width=150)
    tabla.column("Total", anchor="center", width=150)

    for i in range(len(resumen)):
        tabla.insert("", "end",
                     values=(f"{resumen['Aprobados (%)'][i]:.1f}%",
                             f"{resumen['Reprobados (%)'][i]:.1f}%",
                             resumen["Total alumnos"][i]))
    tabla.pack(pady=10)

    def guardar_grafica():
        ruta = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png")],
            title="Guardar gráfica"
        )
        if ruta:
            fig.savefig(ruta, dpi=300, bbox_inches="tight")
            messagebox.showinfo("Guardado", f"Gráfica guardada en:\n{ruta}")

    botones = ctk.CTkFrame(frame)
    botones.pack(pady=15)
    ctk.CTkButton(botones, text="💾 Guardar gráfica", command=guardar_grafica).grid(row=0, column=0, padx=10)

    def on_close():
        plt.close(fig)
        ventana.destroy()

    ventana.protocol("WM_DELETE_WINDOW", on_close)


# ---------- comparar_promedio_total ----------
def comparar_promedio_total():
    """Compara el promedio total (CALIF DIAG) entre dos archivos Excel con scroll."""

    ruta1 = filedialog.askopenfilename(
        title="Selecciona el primer archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta1:
        messagebox.showinfo("Cancelado", "No se seleccionó el primer archivo.")
        return

    ruta2 = filedialog.askopenfilename(
        title="Selecciona el segundo archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta2:
        messagebox.showinfo("Cancelado", "No se seleccionó el segundo archivo.")
        return

    try:
        df1 = pd.read_excel(ruta1)
        df2 = pd.read_excel(ruta2)
    except Exception as e:
        messagebox.showerror("Error al leer Excel", str(e))
        return

    for df in [df1, df2]:
        if "CALIF DIAG" not in df.columns:
            messagebox.showerror("Error", "Uno de los archivos no contiene la columna 'CALIF DIAG'.")
            return

    prom1 = df1["CALIF DIAG"].mean()
    prom2 = df2["CALIF DIAG"].mean()

    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    resumen = pd.DataFrame({
        "Archivo": [nombre1, nombre2],
        "Promedio total": [prom1, prom2],
        "Total alumnos": [len(df1), len(df2)]
    })

    # --- Ventana con scroll ---
    ventana = ctk.CTkToplevel()
    ventana.title("Comparación de Promedio Total")
    ventana.geometry("900x700")

    frame = crear_frame_scrollable_toplevel(ventana)

    # --- Gráfica ---
    fig, ax = plt.subplots(figsize=(8, 6))
    x = range(len(resumen))

    ax.bar(x, resumen["Promedio total"], width=0.5, label="Promedio total")

    for i, v in enumerate(resumen["Promedio total"]):
        ax.text(i, v + 1, f"{v:.2f}", ha='center', fontsize=10)

    ax.set_xticks(x)
    ax.set_xticklabels(resumen["Archivo"], rotation=15)
    ax.set_ylabel("Promedio")
    ax.set_title("Comparación del Promedio Total")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", linestyle="--", alpha=0.6)
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10)

    # --- Tabla ---
    tabla_frame = ctk.CTkFrame(frame)
    tabla_frame.pack(pady=10)

    tabla = ttk.Treeview(tabla_frame, columns=("Promedio", "Total"), show="headings", height=3)
    tabla.heading("Promedio", text="Promedio total")
    tabla.heading("Total", text="Total alumnos")
    tabla.column("Promedio", anchor="center", width=150)
    tabla.column("Total", anchor="center", width=150)

    for i in range(len(resumen)):
        tabla.insert("", "end",
                     values=(f"{resumen['Promedio total'][i]:.2f}",
                             resumen["Total alumnos"][i]))
    tabla.pack(pady=10)

    def guardar_grafica():
        ruta = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png")],
            title="Guardar gráfica"
        )
        if ruta:
            fig.savefig(ruta, dpi=300)
            messagebox.showinfo("Guardado", f"Gráfica guardada en:\n{ruta}")

    botones = ctk.CTkFrame(frame)
    botones.pack(pady=15)
    ctk.CTkButton(botones, text="💾 Guardar gráfica", command=guardar_grafica).grid(row=0, column=0, padx=10)

    def on_close():
        plt.close(fig)
        ventana.destroy()

    ventana.protocol("WM_DELETE_WINDOW", on_close)


# ---------- comparar_promedios_por_carrera_dos_archivos ----------
def comparar_promedios_por_carrera_dos_archivos():
    """Compara el promedio final por carrera entre dos archivos Excel con scroll."""

    ruta1 = filedialog.askopenfilename(
        title="Selecciona el primer archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta1:
        messagebox.showinfo("Cancelado", "No se seleccionó el primer archivo.")
        return

    ruta2 = filedialog.askopenfilename(
        title="Selecciona el segundo archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta2:
        messagebox.showinfo("Cancelado", "No se seleccionó el segundo archivo.")
        return

    try:
        df1 = pd.read_excel(ruta1)
        df2 = pd.read_excel(ruta2)
    except Exception as e:
        messagebox.showerror("Error al leer Excel", str(e))
        return

    for df in [df1, df2]:
        if "carrera" not in df.columns:
            messagebox.showerror("Error", "Falta la columna 'carrera'.")
            return
        if "CALIF DIAG" not in df.columns:
            messagebox.showerror("Error", "Falta la columna 'CALIF DIAG'.")
            return

    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    prom1 = df1.groupby("carrera")["CALIF DIAG"].mean().reset_index()
    prom2 = df2.groupby("carrera")["CALIF DIAG"].mean().reset_index()

    comparacion = pd.merge(prom1, prom2, on="carrera", how="inner", suffixes=("_" + nombre1, "_" + nombre2))
    comparacion = comparacion.sort_values(by="carrera")

    # --- Ventana con scroll ---
    ventana = ctk.CTkToplevel()
    ventana.title("Comparación de promedios por carrera")
    ventana.geometry("1100x800")

    frame = crear_frame_scrollable_toplevel(ventana)

    # --- Gráfica ---
    fig, ax = plt.subplots(figsize=(12, 6))
    x = range(len(comparacion))

    ax.scatter(x, comparacion[f"CALIF DIAG_{nombre1}"], label=f"{nombre1}", s=80)
    ax.scatter(x, comparacion[f"CALIF DIAG_{nombre2}"], label=f"{nombre2}", s=80)

    for i, row in comparacion.iterrows():
        idx = list(range(len(comparacion)))[i]
        ax.text(idx, row[f"CALIF DIAG_{nombre1}"] + 1,
                f"{row[f'CALIF DIAG_{nombre1}']:.1f}", ha="center", fontsize=8)
        ax.text(idx, row[f"CALIF DIAG_{nombre2}"] + 1,
                f"{row[f'CALIF DIAG_{nombre2}']:.1f}", ha="center", fontsize=8, color="red")

    ax.set_xticks(x)
    ax.set_xticklabels(comparacion["carrera"], rotation=25, ha="right")
    ax.set_ylim(0, 100)
    ax.set_ylabel("Promedio final (CALIF DIAG)")
    ax.set_title("Comparación de promedios por carrera")
    ax.grid(axis="y", linestyle="--", alpha=0.6)
    ax.legend()
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=20)

    def guardar_grafica():
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png")],
            title="Guardar gráfica"
        )
        if ruta_guardado:
            fig.savefig(ruta_guardado, dpi=300)
            messagebox.showinfo("Guardado", f"Gráfica guardada en:\n{ruta_guardado}")

    botones = ctk.CTkFrame(frame)
    botones.pack(pady=10)
    ctk.CTkButton(botones, text="💾 Guardar gráfica", command=guardar_grafica).grid(row=0, column=0, padx=10)

    def on_close():
        plt.close(fig)
        ventana.destroy()

    ventana.protocol("WM_DELETE_WINDOW", on_close)


# ---------- comparar_promedio_final_por_carrera ----------
def comparar_promedio_final_por_carrera():
    """Lee un archivo Excel y grafica el promedio final por carrera con scroll."""

    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta:
        messagebox.showinfo("Cancelado", "No se seleccionó ningún archivo.")
        return

    try:
        df = pd.read_excel(ruta)
    except Exception as e:
        messagebox.showerror("Error al leer Excel", str(e))
        return

    if "carrera" not in df.columns:
        messagebox.showerror("Error", "El archivo no contiene la columna 'carrera'.")
        return

    if "CALIF DIAG" not in df.columns:
        messagebox.showerror("Error", "El archivo no contiene la columna 'CALIF DIAG'.")
        return

    nombre_archivo = os.path.basename(ruta).replace(".xlsx", "").replace(".xls", "")

    promedios = df.groupby("carrera")["CALIF DIAG"].mean().reset_index()
    promedios = promedios.sort_values(by="carrera", ascending=True)

    # --- Ventana con scroll ---
    ventana = ctk.CTkToplevel()
    ventana.title(f"Promedio final por carrera - {nombre_archivo}")
    ventana.geometry("1100x800")

    frame = crear_frame_scrollable_toplevel(ventana)

    # --- Gráfica ---
    fig, ax = plt.subplots(figsize=(10, 6))

    barras = ax.bar(promedios["carrera"], promedios["CALIF DIAG"], color="mediumslateblue")

    for barra, (_, row) in zip(barras, promedios.iterrows()):
        altura = barra.get_height()
        ax.text(
            barra.get_x() + barra.get_width() / 2,
            altura + 1,
            f"{altura:.2f}",
            ha="center", va="bottom", fontsize=9
        )

    ax.set_ylim(0, 100)
    ax.set_xticklabels(promedios["carrera"], rotation=25, ha="right")
    ax.set_ylabel("Promedio final")
    ax.set_title(f"Promedio final por carrera ({nombre_archivo})")
    ax.grid(axis="y", linestyle="--", alpha=0.6)
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=20)

    def guardar_grafica():
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("Imagen PNG", "*.png")],
            title="Guardar gráfica como imagen"
        )
        if ruta_guardado:
            fig.savefig(ruta_guardado, dpi=300)
            messagebox.showinfo("Guardado", f"Gráfica guardada en:\n{ruta_guardado}")

    botones = ctk.CTkFrame(frame)
    botones.pack(pady=10)
    ctk.CTkButton(botones, text="💾 Guardar gráfica", command=guardar_grafica).grid(row=0, column=0, padx=10)

    def on_close():
        plt.close(fig)
        ventana.destroy()

    ventana.protocol("WM_DELETE_WINDOW", on_close)


def combinar_diag_con_final():
    ruta_1 = filedialog.askopenfilename(title="Selecciona el Excel diagnóstico combinado calificado", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not ruta_1:
        messagebox.showwarning("Archivo faltante", "No se seleccionó el primer archivo.")
        return

    ruta_2 = filedialog.askopenfilename(title="Selecciona el Excel final combinado calificado", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not ruta_2:
        messagebox.showwarning("Archivo faltante", "No se seleccionó el segundo archivo.")
        return

    try:
        df1 = pd.read_excel(ruta_1)
        df2 = pd.read_excel(ruta_2)

        df1["NUMERO DE CONTROL"] = pd.to_numeric(df1["NUMERO DE CONTROL"], errors="coerce")
        df2["NUMERO DE CONTROL"] = pd.to_numeric(df2["NUMERO DE CONTROL"], errors="coerce")

        columnas_clave = ["NUMERO DE CONTROL", "carrera", "ppgrupo", "NOMBRE COMPLETO", "turno"]
        for col in columnas_clave:
            if col not in df1.columns or col not in df2.columns:
                messagebox.showerror("Error", f"Falta la columna '{col}' en uno de los archivos.")
                return

        df1_extra = df1.drop(columns=columnas_clave).add_prefix("DIAG_")
        df2_extra = df2.drop(columns=columnas_clave).add_prefix("FINAL_")

        df1_base = df1[columnas_clave].copy()
        df2_base = df2[columnas_clave].copy()

        df1_renombrado = pd.concat([df1_base, df1_extra], axis=1)
        df2_renombrado = pd.concat([df2_base, df2_extra], axis=1)

        df_combinado = pd.merge(df1_renombrado, df2_renombrado, on=columnas_clave, how="outer")
        df_combinado.sort_values(by=columnas_clave, inplace=True)
        df_final = df_combinado.groupby("NUMERO DE CONTROL", as_index=False).first()

        ruta_guardado = filedialog.asksaveasfilename(title="Guardar archivo combinado", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if ruta_guardado:
            df_final.to_excel(ruta_guardado, index=False)
            messagebox.showinfo("Éxito", f"Archivo combinado guardado:\n{ruta_guardado}")
        else:
            messagebox.showwarning("Cancelado", "No se guardó el archivo combinado.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema al combinar los archivos:\n{str(e)}")


def combinar_combinado_completo_con_base_datos():
    ruta_2 = filedialog.askopenfilename(
        title="Selecciona el Excel final combinado calificado",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not ruta_2:
        messagebox.showwarning("Archivo faltante", "No se seleccionó el segundo archivo.")
        return

    try:
        df1 = base_datos
        df2 = pd.read_excel(ruta_2)

        if "ficha" not in df1.columns:
            messagebox.showerror("Error", "El primer archivo no contiene la columna 'ficha'.")
            return
        df1 = df1.rename(columns={"ficha": "NUMERO DE CONTROL"})

        df1['entidad_procedencia'] = pd.to_numeric(df1['entidad_procedencia'], errors='coerce')
        df1["NUMERO DE CONTROL"] = pd.to_numeric(df1["NUMERO DE CONTROL"], errors="coerce")
        df2["NUMERO DE CONTROL"] = pd.to_numeric(df2["NUMERO DE CONTROL"], errors="coerce")

        columnas_clave = ["NUMERO DE CONTROL", "carrera", "ppgrupo", "NOMBRE COMPLETO", "turno"]
        for col in columnas_clave:
            if col not in df2.columns:
                messagebox.showerror("Error", f"Falta la columna '{col}' en el segundo archivo.")
                return

        columnas_df2 = [col for col in df2.columns if col != "NUMERO DE CONTROL"]

        df_combinado = pd.merge(df1, df2[columnas_clave + [col for col in columnas_df2 if col not in columnas_clave]],
                                on="NUMERO DE CONTROL", how="left")

        ruta_guardado = filedialog.asksaveasfilename(
            title="Guardar archivo combinado",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if ruta_guardado:
            df_combinado.to_excel(ruta_guardado, index=False)
            messagebox.showinfo("Éxito", f"Archivo combinado guardado:\n{ruta_guardado}")
        else:
            messagebox.showwarning("Cancelado", "No se guardó el archivo combinado.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema al combinar los archivos:\n{str(e)}")


def cargar_formateador(tipo_form):
    def procesar_csv():
        csv_path = filedialog.askopenfilename(
            title="Selecciona el archivo CSV",
            filetypes=[("CSV files", "*.csv")]
        )

        if not csv_path:
            messagebox.showerror("Error", "No se seleccionó ningún archivo CSV.")
            return False

        df = pd.read_csv(csv_path, sep=';')

        print("Columnas disponibles en el CSV:")
        print(df.columns.tolist())

        df = df.rename(columns={
            df.columns[0]: "File name",
            df.columns[1]: "GRUPO.1",
            df.columns[2]: "GRUPO.2",
            df.columns[3]: "GRUPO.3",
            df.columns[4]: "NUMERO DE CONTROL.1",
            df.columns[5]: "NUMERO DE CONTROL.2",
            df.columns[6]: "NUMERO DE CONTROL.3",
            df.columns[7]: "NUMERO DE CONTROL.4",
            df.columns[8]: "P1",
            df.columns[9]: "P2",
            df.columns[10]: "P3",
            df.columns[11]: "P4",
            df.columns[12]: "P5",
            df.columns[13]: "P6",
            df.columns[14]: "P7",
            df.columns[15]: "P8",
            df.columns[16]: "P9",
            df.columns[17]: "P10",
            df.columns[18]: "P11",
            df.columns[19]: "P12",
            df.columns[20]: "P13",
            df.columns[21]: "P14",
            df.columns[22]: "P15",
            df.columns[23]: "P16",
            df.columns[24]: "P17",
            df.columns[25]: "P18",
            df.columns[26]: "P19",
            df.columns[27]: "P20",
            df.columns[28]: "Tipo de Examen"
        })

        df = df.drop(columns=['File name'], errors='ignore')

        df['NUMERO DE CONTROL'] = (
            df['NUMERO DE CONTROL.1'].fillna('').astype(str) +
            df['NUMERO DE CONTROL.2'].fillna('').astype(str) +
            df['NUMERO DE CONTROL.3'].fillna('').astype(str) +
            df['NUMERO DE CONTROL.4'].fillna('').astype(str)
        )

        df['NUMERO DE CONTROL'] = df['NUMERO DE CONTROL'].str.replace('|', '', regex=False)
        df['NUMERO DE CONTROL'] = df['NUMERO DE CONTROL'].astype(str).str.lower().apply(
            lambda s: ''.join(str(ord(ch) - 97) for ch in s if ch.isalpha())
        )
        df['NUMERO DE CONTROL'] = pd.to_numeric(df['NUMERO DE CONTROL'], errors='coerce')

        mapa_letras_grupos = {
            'A': 'A', 'B': 'B', 'C': 'C', 'D': 'D', 'E': 'E',
            'F': 'G', 'G': 'I', 'H': 'L', 'I': 'M', 'J': 'N',
            'K': 'O', 'L': 'P', 'M': 'S'
        }

        df['GRUPO.2'] = 'P'

        df['GRUPO'] = (
            df['GRUPO.1'].fillna('').astype(str) +
            df['GRUPO.2'].fillna('').astype(str) +
            df['GRUPO.3'].fillna('').astype(str)
        )

        df['GRUPO'] = df['GRUPO'].str.replace('|', '', regex=False)
        df['GRUPO'] = df['GRUPO'].map(mapa_letras_grupos).fillna(df['GRUPO'])

        df['carrera'] = ''
        df['ppgrupo'] = ''
        df['NOMBRE COMPLETO'] = ''
        df['CALIF DIAG'] = ''

        if base_datos is None:
            messagebox.showerror("Error", "No se cargo una base de datos.")
            return False

        carrera = seleccionar_carrera()

        df['carrera'] = carrera
        df['ppgrupo'] = df['GRUPO']

        columnas_finales = [
            'carrera', 'ppgrupo', 'NUMERO DE CONTROL', 'NOMBRE COMPLETO',
            'Tipo de Examen', 'CALIF DIAG'
        ] + [f'P{i}' for i in range(1, 21)]
        columnas_existentes = [col for col in columnas_finales if col in df.columns]
        df = df[columnas_existentes]

        excel_path = filedialog.asksaveasfilename(
            title="Guardar archivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not excel_path:
            messagebox.showerror("Error", "No se seleccionó ubicación para guardar el archivo.")
            return False

        df.to_excel(excel_path, index=False, engine='openpyxl')

        continuar = messagebox.askyesno("Proceso completado", f"Archivo guardado como:\n{excel_path}\n\n¿Deseas procesar otro archivo?")
        return continuar

    def procesar_excel():
        excel_path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if not excel_path:
            messagebox.showerror("Error", "No se seleccionó ningún archivo Excel.")
            return False

        df = pd.read_excel(excel_path, engine='openpyxl')

        print("Columnas disponibles en el Excel:")
        print(df.columns.tolist())

        df = df.rename(columns={
            'StudentID': 'NUMERO DE CONTROL',
            'PercentCorrect': 'CALIF DIAG'
        })

        rename_dict = {f'Stu{i}': f'P{i}' for i in range(1, 21)}
        df = df.rename(columns=rename_dict)

        for i in range(1, 21):
            col = f'P{i}'
            if col in df.columns:
                df[col] = df[col].apply(lambda x: str(x)[0] if pd.notna(x) and str(x).strip() != "" else "")

        df['carrera'] = ''
        df['ppgrupo'] = ''
        df['NOMBRE COMPLETO'] = ''

        if base_datos is None:
            messagebox.showerror("Error", "No se cargo una base de datos.")
            return False

        carrera = seleccionar_carrera()
        ppgrupo = tk.simpledialog.askstring("Entrada", "Ingresa el grupo:")

        df['carrera'] = carrera
        df['ppgrupo'] = ppgrupo

        columnas_finales = ['carrera', 'ppgrupo', 'NUMERO DE CONTROL', 'NOMBRE COMPLETO', 'CALIF DIAG'] + [f'P{i}' for i in range(1, 21)]
        columnas_existentes = [col for col in columnas_finales if col in df.columns]
        df = df[columnas_existentes]

        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            messagebox.showerror("Error", "No se seleccionó ubicación para guardar el archivo.")
            return False

        df.to_excel(save_path, index=False, engine='openpyxl')

        continuar = messagebox.askyesno(
            "Proceso completado",
            f"Archivo guardado como:\n{save_path}\n\n¿Deseas procesar otro archivo?"
        )
        return continuar

    def seleccionar_carrera():
        if base_datos is not None:
            carrera = base_datos["carrera"].dropna().astype(str).str.strip().unique().tolist()
            carrera = sorted(carrera)
        else:
            carrera = ["Sin base cargada"]

        win = tk.Toplevel()
        win.title("Seleccionar carrera")
        win.geometry("400x150")
        win.resizable(False, False)
        tk.Label(win, text="Selecciona la carrera:", font=("Arial", 12)).pack(pady=10)

        carrera_var = tk.StringVar()
        combo = ttk.Combobox(win, textvariable=carrera_var, values=carrera, state="readonly", width=40)
        combo.pack(pady=5)
        combo.current(0)

        def confirmar():
            win.destroy()

        tk.Button(win, text="Aceptar", command=confirmar).pack(pady=10)
        win.wait_window()
        return carrera_var.get()

    if tipo_form == 1:
        while True:
            repetir = procesar_csv()
            if not repetir:
                break
    else:
        while True:
            repetir2 = procesar_excel()
            if not repetir2:
                break