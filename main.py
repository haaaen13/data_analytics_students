import pyodbc
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
        print(f"Ruta no válida para el turno {turno}.")
        return None

    dataframesdiag = []

    for archivo in os.listdir(ruta_carpeta):
        if archivo.endswith(".xlsx") or archivo.endswith(".xls"):
            ruta_completa = os.path.join(ruta_carpeta, archivo)
            try:
                df = pd.read_excel(ruta_completa)
                df["archivo_origen"] = archivo
                df["turno"] = turno  # se agrega columna de turno
                dataframesdiag.append(df)
                print(f"[{turno}] Cargado: {archivo}")
            except Exception as e:
                print(f"[{turno}] Error al cargar {archivo}: {e}")

    if dataframesdiag:
        combinado = pd.concat(dataframesdiag, ignore_index=True)
        print(f"[{turno}] Archivos combinados exitosamente.")
        print(combinado.head())

        # Guardar según turno
        if turno.lower() == "mañana":
            base_datos_diag_m = combinado
        else:
            base_datos_diag_t = combinado

        return combinado
    else:
        print(f"[{turno}] No se encontraron archivos Excel en la carpeta.")
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
    "Abrir archivo de concentrado de alumnos"
    )


    # Abre el explorador de archivos
    ruta_archivo = filedialog.askopenfilename(
        title="Selecciona un archivo Excel base para alumnos",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    
    if ruta_archivo:
        # Carga el archivo como DataFrame
        respuestas = pd.read_excel(ruta_archivo)
        print(f"Archivo cargado: {ruta_archivo}")
        print(respuestas.head())  # Muestra las primeras filas
        return respuestas
    else:
        print("No se seleccionó ningún archivo.")
        return None


def comparar_respuestas_y_calificar(df_diag, df_respuestas, turno):
    """
    Compara las respuestas del excel con las respuestas correctas.
    Devuelve un DataFrame con columnas originales y nuevas columnas de aciertos (1/0).
    """
    if df_diag is None or df_respuestas is None:
        messagebox.showerror("Error", f"Faltan archivos de respuestas o respuestas para {turno}.")
        return None

    # Asegurarnos de que las columnas P1...P20 existan en ambos
    preguntas = [f"P{i}" for i in range(1, 21)]
    for col in preguntas:
        if col not in df_diag.columns:
            df_diag[col] = None
        if col not in df_respuestas.columns:
            df_respuestas[col] = None

    # Tomar solo la primera fila del archivo de respuestas (se asume que es una sola clave)
    respuestas_correctas = df_respuestas.iloc[0][preguntas]

    # Crear columnas de acierto (1 si es igual, 0 si no)
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

    # Calificar cada turno
    df_m_cal = comparar_respuestas_y_calificar(df_m, df_resp_m, "mañana") if df_m is not None else None
    df_t_cal = comparar_respuestas_y_calificar(df_t, df_resp_t, "tarde") if df_t is not None else None

    # Guardar los archivos individualmente
    ruta_guardado = filedialog.askdirectory(title="Selecciona la carpeta donde guardar los archivos calificados")
    if not ruta_guardado:
        messagebox.showwarning("Aviso", "No se seleccionó carpeta para guardar los resultados.")
        return

    if df_m_cal is not None:
        ruta_m = os.path.join(ruta_guardado, "Exámen_Matutino_Calificado.xlsx")
        df_m_cal.to_excel(ruta_m, index=False)
        print(f"💾 Archivo guardado: {ruta_m}")

    if df_t_cal is not None:
        ruta_t = os.path.join(ruta_guardado, "Exámen_Vespertino_Calificado.xlsx")
        df_t_cal.to_excel(ruta_t, index=False)
        print(f"💾 Archivo guardado: {ruta_t}")

    # Generar combinado
    combinados = []
    if df_m_cal is not None:
        combinados.append(df_m_cal)
    if df_t_cal is not None:
        combinados.append(df_t_cal)

    if combinados:
        df_total = pd.concat(combinados, ignore_index=True)
        ruta_c = os.path.join(ruta_guardado, "Exámen_Combinado.xlsx")
        df_total.to_excel(ruta_c, index=False)
        print(f"💾 Archivo combinado guardado: {ruta_c}")
        messagebox.showinfo("Éxito", "Archivos calificados generados correctamente.")
        return df_total

    return None
















excel_combinado = None  # Aquí se guardará el archivo combinado cargado

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
def analizar_datos():
    """Analiza un solo archivo con opción de filtrar por carrera y grupo.
    Eje X fijo: P1..P20."""
    global excel_combinado
    if excel_combinado is None:
        messagebox.showerror("Error", "Primero carga el archivo combinado para análisis.")
        return

    # Crear ventana de análisis
    ventana_analisis = tk.Toplevel()
    ventana_analisis.title("Analíticas del Examen")
    ventana_analisis.geometry("1300x850")

    frame_scroll = ctk.CTkScrollableFrame(ventana_analisis, label_text="Resultados Analíticos")
    frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    # --- Combobox de filtros dinámicos ---
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

    # --- Actualiza los grupos según la carrera seleccionada ---
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

            if grupos_filtrados:
                combo_grupo["values"] = ["(Todos)"] + grupos_filtrados
                combo_grupo.set("(Todos)")
            else:
                combo_grupo["values"] = ["(Todos)"]
                combo_grupo.set("(Todos)")

    combo_carrera.bind("<<ComboboxSelected>>", actualizar_grupos)

    # Frame donde se colocarán las gráficas
    frame_graficas = ctk.CTkFrame(frame_scroll)
    frame_graficas.pack(fill="both", expand=True, pady=20)

    preguntas_esperadas = [f"P{i}_correcta" for i in range(1, 21)]
    x = range(1, 21)

    def generar_graficas():
        for widget in frame_graficas.winfo_children():
            widget.destroy()  # limpiar gráficas anteriores

        df = excel_combinado.copy()

        # --- Aplicar filtros ---
        carrera_sel = combo_carrera.get()
        grupo_sel = combo_grupo.get()
        if carrera_sel != "(Todas)" and "carrera" in df.columns:
            df = df[df["carrera"].astype(str).str.strip() == str(carrera_sel).strip()]
        if grupo_sel != "(Todos)" and "ppgrupo" in df.columns:
            df = df[df["ppgrupo"].astype(str).str.strip() == str(grupo_sel).strip()]

        # --- 1️⃣ Errores por pregunta (siempre P1..P20) ---
        for col in preguntas_esperadas:
            if col not in df.columns:
                df[col] = float("nan")

        aciertos = df[preguntas_esperadas].mean()
        aciertos = aciertos.reindex(preguntas_esperadas)

        plt.figure(figsize=(10, 6))
        plt.scatter(x, aciertos.values * 100, color="skyblue")
        plt.title("Porcentaje de aciertos por pregunta")
        plt.xlabel("Pregunta")
        plt.ylabel("% de aciertos")
        plt.xticks(x, [f"P{i}" for i in x])
        plt.xlim(0.5, 20.5)
        plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100
        plt.grid(True)
        plt.tight_layout()

        canvas1 = FigureCanvasTkAgg(plt.gcf(), master=frame_graficas)
        canvas1.draw()
        canvas1.get_tk_widget().pack(pady=20)

        # --- 2️⃣ Promedios por carrera ---
        if "carrera" in df.columns and "CALIFICACION DIAG" in df.columns:
            prom = df.groupby(df["carrera"].astype(str).str.strip())["CALIFICACION DIAG"].mean().sort_values(ascending=False)
            plt.figure(figsize=(10, 6))
            prom.plot(kind="bar", color="skyblue")
            plt.title("Promedio de calificación diagnóstica por carrera")
            plt.xlabel("Carrera")
            plt.ylabel("Promedio")
            plt.xticks(rotation=45, ha="right")
            plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100
            plt.tight_layout()

            canvas2 = FigureCanvasTkAgg(plt.gcf(), master=frame_graficas)
            canvas2.draw()
            canvas2.get_tk_widget().pack(pady=20)

        # --- 3️⃣ Aprobados vs reprobados ---
        if "CALIFICACION DIAG" in df.columns:
            aprobados = (df["CALIFICACION DIAG"] >= 60).sum()
            reprobados = (df["CALIFICACION DIAG"] < 60).sum()
            df_aprob = pd.DataFrame({"Resultado": ["Aprobado", "Reprobado"], "Cantidad": [aprobados, reprobados]})

            plt.figure(figsize=(6, 6))
            sns.barplot(data=df_aprob, x="Resultado", y="Cantidad", palette="pastel")
            plt.title("Total de exámenes aprobados y reprobados")
            plt.xlabel("")
            plt.ylabel("Número de alumnos")
            plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100 (si deseas mantener comparabilidad)
            plt.tight_layout()

            canvas3 = FigureCanvasTkAgg(plt.gcf(), master=frame_graficas)
            canvas3.draw()
            canvas3.get_tk_widget().pack(pady=20)

    # Botón para generar gráficas filtradas
    boton_filtrar = ctk.CTkButton(frame_scroll, text="🔍 Aplicar filtros y generar gráficas", command=generar_graficas)
    boton_filtrar.pack(pady=10)

    # Auto-generar sin filtro al abrir
    generar_graficas()

def analizar_datos_errores():
    """Analiza un solo archivo con opción de filtrar por carrera y grupo.
    Eje X fijo: P1..P20."""
    global excel_combinado
    if excel_combinado is None:
        messagebox.showerror("Error", "Primero carga el archivo combinado para análisis.")
        return

    # Crear ventana de análisis
    ventana_analisis = tk.Toplevel()
    ventana_analisis.title("Analíticas del Examen")
    ventana_analisis.geometry("1300x850")

    frame_scroll = ctk.CTkScrollableFrame(ventana_analisis, label_text="Resultados Analíticos")
    frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    # --- Combobox de filtros dinámicos ---
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

    # --- Actualiza los grupos según la carrera seleccionada ---
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

            if grupos_filtrados:
                combo_grupo["values"] = ["(Todos)"] + grupos_filtrados
                combo_grupo.set("(Todos)")
            else:
                combo_grupo["values"] = ["(Todos)"]
                combo_grupo.set("(Todos)")

    combo_carrera.bind("<<ComboboxSelected>>", actualizar_grupos)

    # Frame donde se colocarán las gráficas
    frame_graficas = ctk.CTkFrame(frame_scroll)
    frame_graficas.pack(fill="both", expand=True, pady=20)

    preguntas_esperadas = [f"P{i}_correcta" for i in range(1, 21)]
    x = range(1, 21)

    def generar_graficas():
        for widget in frame_graficas.winfo_children():
            widget.destroy()  # limpiar gráficas anteriores

        df = excel_combinado.copy()

        # --- Aplicar filtros ---
        carrera_sel = combo_carrera.get()
        grupo_sel = combo_grupo.get()
        if carrera_sel != "(Todas)" and "carrera" in df.columns:
            df = df[df["carrera"].astype(str).str.strip() == str(carrera_sel).strip()]
        if grupo_sel != "(Todos)" and "ppgrupo" in df.columns:
            df = df[df["ppgrupo"].astype(str).str.strip() == str(grupo_sel).strip()]

        # --- 1️⃣ Errores por pregunta (siempre P1..P20) ---
        for col in preguntas_esperadas:
            if col not in df.columns:
                df[col] = float("nan")

        errores = 1- df[preguntas_esperadas].mean()
        errores = errores.reindex(preguntas_esperadas)

        plt.figure(figsize=(10, 6))
        plt.scatter(x, errores.values * 100, color="skyblue")
        plt.title("Porcentaje de errores por pregunta")
        plt.xlabel("Pregunta")
        plt.ylabel("% de errores")
        plt.xticks(x, [f"P{i}" for i in x])
        plt.xlim(0.5, 20.5)
        plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100
        plt.grid(True)
        plt.tight_layout()

        canvas1 = FigureCanvasTkAgg(plt.gcf(), master=frame_graficas)
        canvas1.draw()
        canvas1.get_tk_widget().pack(pady=20)

        # --- 2️⃣ Promedios por carrera ---
        if "carrera" in df.columns and "CALIFICACION DIAG" in df.columns:
            prom = df.groupby(df["carrera"].astype(str).str.strip())["CALIFICACION DIAG"].mean().sort_values(ascending=False)
            plt.figure(figsize=(10, 6))
            prom.plot(kind="bar", color="skyblue")
            plt.title("Promedio de calificación diagnóstica por carrera")
            plt.xlabel("Carrera")
            plt.ylabel("Promedio")
            plt.xticks(rotation=45, ha="right")
            plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100
            plt.tight_layout()

            canvas2 = FigureCanvasTkAgg(plt.gcf(), master=frame_graficas)
            canvas2.draw()
            canvas2.get_tk_widget().pack(pady=20)

        # --- 3️⃣ Aprobados vs reprobados ---
        if "CALIFICACION DIAG" in df.columns:
            aprobados = (df["CALIFICACION DIAG"] >= 60).sum()
            reprobados = (df["CALIFICACION DIAG"] < 60).sum()
            df_aprob = pd.DataFrame({"Resultado": ["Aprobado", "Reprobado"], "Cantidad": [aprobados, reprobados]})

            plt.figure(figsize=(6, 6))
            sns.barplot(data=df_aprob, x="Resultado", y="Cantidad", palette="pastel")
            plt.title("Total de exámenes aprobados y reprobados")
            plt.xlabel("")
            plt.ylabel("Número de alumnos")
            plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100 (si deseas mantener comparabilidad)
            plt.tight_layout()

            canvas3 = FigureCanvasTkAgg(plt.gcf(), master=frame_graficas)
            canvas3.draw()
            canvas3.get_tk_widget().pack(pady=20)

    # Botón para generar gráficas filtradas
    boton_filtrar = ctk.CTkButton(frame_scroll, text="🔍 Aplicar filtros y generar gráficas", command=generar_graficas)
    boton_filtrar.pack(pady=10)

    # Auto-generar sin filtro al abrir
    generar_graficas()

# ---------- analizar_datos2 (comparativo entre 2 archivos) ----------
def analizar_datos2():
    """Permite seleccionar y comparar dos archivos Excel con opción de filtrar por grupo (ppgrupo),
    además de restablecer la vista general sin filtros y guardar la gráfica como PNG."""

    # --- Seleccionar los archivos ---
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

    # Obtener nombres base de los archivos (sin extensión)
    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    # --- Crear ventana principal ---
    ventana_comp = tk.Toplevel()
    ventana_comp.title("Comparador de Resultados (dos archivos)")
    ventana_comp.geometry("1300x850")

    frame_scroll = ctk.CTkScrollableFrame(ventana_comp, label_text="Comparación de grupos o archivos")
    frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    # --- Combobox de grupos para cada archivo ---
    ctk.CTkLabel(frame_scroll, text=f"Grupo en {nombre1}:").pack(pady=5)
    if "ppgrupo" in excel1.columns:
        grupos1 = sorted(excel1["ppgrupo"].dropna().astype(str).str.strip().unique().tolist())
    else:
        grupos1 = []
    combo_grupo1 = ttk.Combobox(frame_scroll, values=["(Todos)"] + grupos1)
    combo_grupo1.set("(Todos)")
    combo_grupo1.pack(pady=5)

    ctk.CTkLabel(frame_scroll, text=f"Grupo en {nombre2}:").pack(pady=5)
    if "ppgrupo" in excel2.columns:
        grupos2 = sorted(excel2["ppgrupo"].dropna().astype(str).str.strip().unique().tolist())
    else:
        grupos2 = []
    combo_grupo2 = ttk.Combobox(frame_scroll, values=["(Todos)"] + grupos2)
    combo_grupo2.set("(Todos)")
    combo_grupo2.pack(pady=5)

    # --- Frame donde se dibujarán las gráficas ---
    frame_graficas = ctk.CTkFrame(frame_scroll)
    frame_graficas.pack(fill="both", expand=True, pady=20)

    preguntas = [f"P{i}_correcta" for i in range(1, 21)]
    x = range(1, 21)

    # --- Función para generar las gráficas ---
    def generar_comparacion():
        for w in frame_graficas.winfo_children():
            w.destroy()

        grupo1 = combo_grupo1.get()
        grupo2 = combo_grupo2.get()

        df1 = excel1.copy()
        df2 = excel2.copy()

        # Filtrado independiente por grupo
        if grupo1 != "(Todos)" and "ppgrupo" in df1.columns:
            df1 = df1[df1["ppgrupo"].astype(str).str.strip() == grupo1]
        if grupo2 != "(Todos)" and "ppgrupo" in df2.columns:
            df2 = df2[df2["ppgrupo"].astype(str).str.strip() == grupo2]

        # Asegurar columnas P1..P20_correcta
        for df in [df1, df2]:
            for col in preguntas:
                if col not in df.columns:
                    df[col] = float("nan")

        # Calcular % de error
        e1 = ( df1[preguntas].mean()) * 100
        e2 = ( df2[preguntas].mean()) * 100

        # --- Gráfica 1: comparación de errores ---
        plt.figure(figsize=(10, 6))
        plt.plot(x, e1.values, marker="o", linestyle="--", label=f"{nombre1} - {grupo1}")
        plt.plot(x, e2.values, marker="o", linestyle="--", label=f"{nombre2} - {grupo2}")
        plt.title("Comparación de aciertos por pregunta")
        plt.xlabel("Pregunta")
        plt.ylabel("% de aciertos")
        plt.xticks(x, [f"P{i}" for i in x])
        plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100
        plt.grid(True)
        plt.legend()
        plt.tight_layout()

        # Guardar referencia de la figura para el botón de guardado
        fig_comparacion = plt.gcf()

        canvas1 = FigureCanvasTkAgg(fig_comparacion, master=frame_graficas)
        canvas1.draw()
        canvas1.get_tk_widget().pack(pady=20)

        # --- Gráfica 2: Aprobados vs Reprobados ---
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
            plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100 para uniformidad visual
            plt.tight_layout()

            fig_aprob = plt.gcf()
            canvas2 = FigureCanvasTkAgg(fig_aprob, master=frame_graficas)
            canvas2.draw()
            canvas2.get_tk_widget().pack(pady=20)

        # --- Botón para guardar la gráfica ---
        def guardar_png():
            archivo = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("Imagen PNG", "*.png")],
                title="Guardar gráfica como PNG"
            )
            if archivo:
                fig_comparacion.savefig(archivo, dpi=300)
                messagebox.showinfo("Guardado", f"Gráfica guardada como:\n{archivo}")

        boton_guardar = ctk.CTkButton(frame_graficas, text="💾 Guardar gráfica como PNG", command=guardar_png)
        boton_guardar.pack(pady=10)

    # --- Función para restablecer la comparación general ---
    def resetear_comparacion():
        combo_grupo1.set("(Todos)")
        combo_grupo2.set("(Todos)")
        generar_comparacion()

    # --- Botones de acción ---
    botones_frame = ctk.CTkFrame(frame_scroll)
    botones_frame.pack(pady=10)

    boton_generar = ctk.CTkButton(botones_frame, text="📊 Generar comparación", command=generar_comparacion)
    boton_generar.grid(row=0, column=0, padx=10)

    boton_reset = ctk.CTkButton(botones_frame, text="🔄 Restablecer comparación general", fg_color="gray", command=resetear_comparacion)
    boton_reset.grid(row=0, column=1, padx=10)

    # --- Genera la comparación inicial (sin filtros) ---
    generar_comparacion()


def analizar_datos2_errores():
    """Permite seleccionar y comparar dos archivos Excel con opción de filtrar por grupo (ppgrupo),
    además de restablecer la vista general sin filtros y guardar la gráfica como PNG."""

    # --- Seleccionar los archivos ---
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

    # Obtener nombres base de los archivos (sin extensión)
    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    # --- Crear ventana principal ---
    ventana_comp = tk.Toplevel()
    ventana_comp.title("Comparador de Resultados (dos archivos)")
    ventana_comp.geometry("1300x850")

    frame_scroll = ctk.CTkScrollableFrame(ventana_comp, label_text="Comparación de grupos o archivos")
    frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    # --- Combobox de grupos para cada archivo ---
    ctk.CTkLabel(frame_scroll, text=f"Grupo en {nombre1}:").pack(pady=5)
    if "ppgrupo" in excel1.columns:
        grupos1 = sorted(excel1["ppgrupo"].dropna().astype(str).str.strip().unique().tolist())
    else:
        grupos1 = []
    combo_grupo1 = ttk.Combobox(frame_scroll, values=["(Todos)"] + grupos1)
    combo_grupo1.set("(Todos)")
    combo_grupo1.pack(pady=5)

    ctk.CTkLabel(frame_scroll, text=f"Grupo en {nombre2}:").pack(pady=5)
    if "ppgrupo" in excel2.columns:
        grupos2 = sorted(excel2["ppgrupo"].dropna().astype(str).str.strip().unique().tolist())
    else:
        grupos2 = []
    combo_grupo2 = ttk.Combobox(frame_scroll, values=["(Todos)"] + grupos2)
    combo_grupo2.set("(Todos)")
    combo_grupo2.pack(pady=5)

    # --- Frame donde se dibujarán las gráficas ---
    frame_graficas = ctk.CTkFrame(frame_scroll)
    frame_graficas.pack(fill="both", expand=True, pady=20)

    preguntas = [f"P{i}_correcta" for i in range(1, 21)]
    x = range(1, 21)

    # --- Función para generar las gráficas ---
    def generar_comparacion():
        for w in frame_graficas.winfo_children():
            w.destroy()

        grupo1 = combo_grupo1.get()
        grupo2 = combo_grupo2.get()

        df1 = excel1.copy()
        df2 = excel2.copy()

        # Filtrado independiente por grupo
        if grupo1 != "(Todos)" and "ppgrupo" in df1.columns:
            df1 = df1[df1["ppgrupo"].astype(str).str.strip() == grupo1]
        if grupo2 != "(Todos)" and "ppgrupo" in df2.columns:
            df2 = df2[df2["ppgrupo"].astype(str).str.strip() == grupo2]

        # Asegurar columnas P1..P20_correcta
        for df in [df1, df2]:
            for col in preguntas:
                if col not in df.columns:
                    df[col] = float("nan")

        # Calcular % de error
        e1 = ( 1- df1[preguntas].mean()) * 100
        e2 = ( 1- df2[preguntas].mean()) * 100

        # --- Gráfica 1: comparación de errores ---
        plt.figure(figsize=(10, 6))
        plt.plot(x, e1.values, marker="o", linestyle="--", label=f"{nombre1} - {grupo1}")
        plt.plot(x, e2.values, marker="o", linestyle="--", label=f"{nombre2} - {grupo2}")
        plt.title("Comparación de errores por pregunta")
        plt.xlabel("Pregunta")
        plt.ylabel("% de errores")
        plt.xticks(x, [f"P{i}" for i in x])
        plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100
        plt.grid(True)
        plt.legend()
        plt.tight_layout()
        # Guardar referencia de la figura para el botón de guardado
        fig_comparacion = plt.gcf()

        canvas1 = FigureCanvasTkAgg(fig_comparacion, master=frame_graficas)
        canvas1.draw()
        canvas1.get_tk_widget().pack(pady=20)

        # --- Gráfica 2: Aprobados vs Reprobados ---
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
            plt.ylim(0, 100)  # 🔹 Escala fija 0 a 100 para uniformidad visual
            plt.tight_layout()

            fig_aprob = plt.gcf()
            canvas2 = FigureCanvasTkAgg(fig_aprob, master=frame_graficas)
            canvas2.draw()
            canvas2.get_tk_widget().pack(pady=20)

        # --- Botón para guardar la gráfica ---
        def guardar_png():
            archivo = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("Imagen PNG", "*.png")],
                title="Guardar gráfica como PNG"
            )
            if archivo:
                fig_comparacion.savefig(archivo, dpi=300)
                messagebox.showinfo("Guardado", f"Gráfica guardada como:\n{archivo}")



        boton_guardar = ctk.CTkButton(frame_graficas, text="💾 Guardar gráfica como PNG", command=guardar_png)
        boton_guardar.pack(pady=10)

    # --- Función para restablecer la comparación general ---
    def resetear_comparacion():
        combo_grupo1.set("(Todos)")
        combo_grupo2.set("(Todos)")
        generar_comparacion()




    # --- Botones de acción ---
    botones_frame = ctk.CTkFrame(frame_scroll)
    botones_frame.pack(pady=10)

    boton_generar = ctk.CTkButton(botones_frame, text="📊 Generar comparación", command=generar_comparacion)
    boton_generar.grid(row=0, column=0, padx=10)

    boton_reset = ctk.CTkButton(botones_frame, text="🔄 Restablecer comparación general", fg_color="gray", command=resetear_comparacion)
    boton_reset.grid(row=0, column=1, padx=10)

    # --- Genera la comparación inicial (sin filtros) ---
    generar_comparacion()


def comparar_reprobados():
    """Compara aprobados y reprobados (según menos de 14 correctas) entre dos archivos Excel,
    mostrando la gráfica en Tkinter, la tabla resumen y permitiendo exportar resultados e imagen."""
    
    # --- Seleccionar los archivos ---
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

    # --- Nombres base de los archivos ---
    nombre1 = os.path.basename(ruta1).replace(".xlsx", "").replace(".xls", "")
    nombre2 = os.path.basename(ruta2).replace(".xlsx", "").replace(".xls", "")

    # --- Columnas esperadas ---
    preguntas = [f"P{i}_correcta" for i in range(1, 21)]

    for df in [df1, df2]:
        for col in preguntas:
            if col not in df.columns:
                df[col] = float("nan")

    # --- Calcular total de respuestas correctas ---
    df1["total_correctas"] = df1[preguntas].sum(axis=1)
    df2["total_correctas"] = df2[preguntas].sum(axis=1)

    # --- Determinar reprobados (<14 correctas) y aprobados (>=14) ---
    reprobados1 = (df1["total_correctas"] < 14).sum()
    reprobados2 = (df2["total_correctas"] < 14).sum()

    total1 = len(df1)
    total2 = len(df2)

    aprobados1 = total1 - reprobados1
    aprobados2 = total2 - reprobados2

    # --- Crear DataFrame resumen ---
    resumen = pd.DataFrame({
        "Archivo": [nombre1, nombre2],
        "Aprobados": [aprobados1, aprobados2],
        "Reprobados": [reprobados1, reprobados2],
        "Total alumnos": [total1, total2]
    })

    # --- Crear ventana ---
    ventana_comp = ctk.CTkToplevel()
    ventana_comp.title("Comparación de Aprobados y Reprobados")
    ventana_comp.geometry("950x750")

    frame = ctk.CTkFrame(ventana_comp)
    frame.pack(fill="both", expand=True, padx=20, pady=20)

    # --- Crear gráfica ---
    fig, ax = plt.subplots(figsize=(8, 6))
    x = range(len(resumen))
    ancho = 0.35

    ax.bar([i - ancho/2 for i in x], resumen["Aprobados"], width=ancho, label="Aprobados", color="royalblue")
    ax.bar([i + ancho/2 for i in x], resumen["Reprobados"], width=ancho, label="Reprobados", color="salmon")


    # Mostrar valores sobre las barras
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

    # --- Mostrar gráfica en Tkinter ---
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10)

    # --- Tabla con datos ---
    import tkinter.ttk as ttk

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
        tabla.insert("", "end", values=(resumen["Aprobados"][i], resumen["Reprobados"][i], resumen["Total alumnos"][i]), text=resumen["Archivo"][i])

    tabla.pack(pady=10)

    # --- Botones de acción ---
    def guardar_grafica():
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("Imagen PNG", "*.png")],
            title="Guardar gráfica como imagen"
        )
        if ruta_guardado:
            fig.savefig(ruta_guardado, dpi=300)
            messagebox.showinfo("Guardado", f"✅ Gráfica guardada como:\n{ruta_guardado}")

    def exportar_excel():
        ruta_excel = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
            title="Exportar resultados a Excel"
        )
        if ruta_excel:
            try:
                resumen.to_excel(ruta_excel, index=False)
                messagebox.showinfo("Exportado", f"✅ Resultados exportados a:\n{ruta_excel}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo exportar el archivo:\n{e}")

    botones_frame = ctk.CTkFrame(frame)
    botones_frame.pack(pady=15)

    boton_guardar = ctk.CTkButton(botones_frame, text="💾 Guardar gráfica", command=guardar_grafica)
    boton_guardar.grid(row=0, column=0, padx=10)

    boton_exportar = ctk.CTkButton(botones_frame, text="📤 Exportar resultados a Excel", fg_color="green", command=exportar_excel)
    boton_exportar.grid(row=0, column=1, padx=10)





def comparar_por_carrera_unico():
    """Lee un solo archivo Excel y grafica aprobados y reprobados por carrera,
    mostrando totales y permitiendo exportar resultados y la imagen."""
    
    # --- Seleccionar el archivo ---
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

    # --- Validar columna 'carrera' ---
    if "carrera" not in df.columns:
        messagebox.showerror("Error", "El archivo no contiene la columna 'carrera'.")
        return

    # --- Nombres base ---
    nombre_archivo = os.path.basename(ruta).replace(".xlsx", "").replace(".xls", "")

    # --- Columnas esperadas ---
    preguntas = [f"P{i}_correcta" for i in range(1, 21)]
    for col in preguntas:
        if col not in df.columns:
            df[col] = float("nan")

    # --- Calcular total de correctas ---
    df["total_correctas"] = df[preguntas].sum(axis=1)

    # --- Determinar aprobados/reprobados ---
    df["resultado"] = df["total_correctas"].apply(lambda x: "Aprobado" if x >= 14 else "Reprobado")

    # --- Agrupar por carrera ---
    resumen = df.groupby(["carrera", "resultado"]).size().unstack(fill_value=0)

    # Asegurar columnas para evitar errores
    if "Aprobado" not in resumen.columns:
        resumen["Aprobado"] = 0
    if "Reprobado" not in resumen.columns:
        resumen["Reprobado"] = 0

    resumen = resumen.reset_index()

    # --- Crear ventana ---
    ventana = ctk.CTkToplevel()
    ventana.title(f"Resultados por carrera - {nombre_archivo}")
    ventana.geometry("1100x800")

    frame = ctk.CTkFrame(ventana)
    frame.pack(fill="both", expand=True, padx=20, pady=20)

    # --- Crear gráfica ---
    fig, ax = plt.subplots(figsize=(10, 6))
    x = range(len(resumen))
    ancho = 0.35

    ax.bar([i - ancho/2 for i in x], resumen["Aprobado"], width=ancho, color="royalblue", label="Aprobados ")
    ax.bar([i + ancho/2 for i in x], resumen["Reprobado"], width=ancho, color="salmon", label="Reprobados ")

    # --- Etiquetas con números ---
    for i, row in resumen.iterrows():
        ax.text(i - ancho/2, row["Aprobado"] + 0.5, str(row["Aprobado"]), ha="center", va="bottom", color="blue", fontsize=9)
        ax.text(i + ancho/2, row["Reprobado"] + 0.5, str(row["Reprobado"]), ha="center", va="bottom", color="red", fontsize=9)

    ax.set_xticks(range(len(resumen)))
    ax.set_xticklabels(resumen["carrera"], rotation=25, ha="right")
    ax.set_ylabel("Número de alumnos")
    ax.set_title(f"Aprobados y Reprobados por Carrera ({nombre_archivo})")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.6)
    ax.set_ylim(0, max(resumen[["Aprobado", "Reprobado"]].max().max() * 1.2, 10))
    plt.tight_layout()

    # --- Mostrar en Tkinter ---
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=20)

    # --- Botones para exportar ---
    def guardar_grafica():
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("Imagen PNG", "*.png")],
            title="Guardar gráfica como imagen"
        )
        if ruta_guardado:
            fig.savefig(ruta_guardado, dpi=300)
            messagebox.showinfo("Guardado", f"✅ Gráfica guardada como:\n{ruta_guardado}")

    def exportar_excel():
        ruta_excel = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
            title="Exportar resultados a Excel"
        )
        if ruta_excel:
            try:
                resumen.to_excel(ruta_excel, index=False)
                messagebox.showinfo("Exportado", f"✅ Resultados exportados a:\n{ruta_excel}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo exportar el archivo:\n{e}")

    # --- Frame de botones ---
    botones = ctk.CTkFrame(frame)
    botones.pack(pady=10)

    ctk.CTkButton(botones, text="💾 Guardar gráfica", command=guardar_grafica).grid(row=0, column=0, padx=10)
    ctk.CTkButton(botones, text="📤 Exportar a Excel", fg_color="green", command=exportar_excel).grid(row=0, column=1, padx=10)




def combinar_diag_con_final():
    import pandas as pd
    from tkinter import filedialog, messagebox

    # Seleccionar el primer archivo
    ruta_1 = filedialog.askopenfilename(title="Selecciona el Excel diagnóstico combinado calificado", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not ruta_1:
        messagebox.showwarning("Archivo faltante", "No se seleccionó el primer archivo.")
        return

    # Seleccionar el segundo archivo
    ruta_2 = filedialog.askopenfilename(title="Selecciona el Excel final combinado calificado", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not ruta_2:
        messagebox.showwarning("Archivo faltante", "No se seleccionó el segundo archivo.")
        return

    try:
        # Cargar ambos archivos
        df1 = pd.read_excel(ruta_1)
        df2 = pd.read_excel(ruta_2)

        # Convertir NUMERO DE CONTROL a formato numérico
        df1["NUMERO DE CONTROL"] = pd.to_numeric(df1["NUMERO DE CONTROL"], errors="coerce")
        df2["NUMERO DE CONTROL"] = pd.to_numeric(df2["NUMERO DE CONTROL"], errors="coerce")

        # Columnas clave para hacer el merge
        columnas_clave = ["NUMERO DE CONTROL", "carrera", "ppgrupo", "NOMBRE COMPLETO", "turno"]
        for col in columnas_clave:
            if col not in df1.columns or col not in df2.columns:
                messagebox.showerror("Error", f"Falta la columna '{col}' en uno de los archivos.")
                return

        # Separar columnas no clave y renombrarlas para evitar colisiones
        df1_extra = df1.drop(columns=columnas_clave).add_prefix("DIAG_")
        df2_extra = df2.drop(columns=columnas_clave).add_prefix("FINAL_")

        # Unir columnas clave con sus respectivas columnas extra
        df1_base = df1[columnas_clave].copy()
        df2_base = df2[columnas_clave].copy()

        df1_renombrado = pd.concat([df1_base, df1_extra], axis=1)
        df2_renombrado = pd.concat([df2_base, df2_extra], axis=1)

        # Combinar ambos DataFrames por columnas clave
        df_combinado = pd.merge(df1_renombrado, df2_renombrado, on=columnas_clave, how="outer")

        # Agrupar por NUMERO DE CONTROL y conservar la primera fila con datos no nulos
        df_combinado.sort_values(by=columnas_clave, inplace=True)
        df_final = df_combinado.groupby("NUMERO DE CONTROL", as_index=False).first()

        # Guardar el archivo combinado
        ruta_guardado = filedialog.asksaveasfilename(title="Guardar archivo combinado", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if ruta_guardado:
            df_final.to_excel(ruta_guardado, index=False)
            messagebox.showinfo("Éxito", f"Archivo combinado guardado:\n{ruta_guardado}")
        else:
            messagebox.showwarning("Cancelado", "No se guardó el archivo combinado.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema al combinar los archivos:\n{str(e)}")




def combinar_combinado_completo_con_base_datos():
    import pandas as pd
    from tkinter import filedialog, messagebox

    # Seleccionar el archivo final (df2)
    ruta_2 = filedialog.askopenfilename(
        title="Selecciona el Excel final combinado calificado",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not ruta_2:
        messagebox.showwarning("Archivo faltante", "No se seleccionó el segundo archivo.")
        return

    try:
        # Cargar archivos
        df1 = base_datos
        df2 = pd.read_excel(ruta_2)

        # Renombrar 'ficha' a 'NUM_CONTROL' en df1
        if "ficha" not in df1.columns:
            messagebox.showerror("Error", "El primer archivo no contiene la columna 'ficha'.")
            return
        df1 = df1.rename(columns={"ficha": "NUMERO DE CONTROL"})

        # Convertir NUM_CONTROL a numérico
        
        df1["NUMERO DE CONTROL"] = pd.to_numeric(df1["NUMERO DE CONTROL"], errors="coerce")
        df2["NUMERO DE CONTROL"] = pd.to_numeric(df2["NUMERO DE CONTROL"], errors="coerce")

        

        # Verificar que df2 tenga las columnas clave
        columnas_clave = ["NUMERO DE CONTROL", "carrera", "ppgrupo", "NOMBRE COMPLETO", "turno"]
        for col in columnas_clave:
            if col not in df2.columns:
                messagebox.showerror("Error", f"Falta la columna '{col}' en el segundo archivo.")
                return

        # Evitar duplicación de columnas clave en el merge
        columnas_df2 = [col for col in df2.columns if col != "NUMERO DE CONTROL"]

        # Realizar merge con df1 como base y df2 al final
        df_combinado = pd.merge(df1, df2[columnas_clave + [col for col in columnas_df2 if col not in columnas_clave]],
                                on="NUMERO DE CONTROL", how="left")

        # Guardar resultado
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