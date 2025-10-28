import pyodbc
import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog
import seaborn as sns
import os


#base_datos=pd.read_excel(r"C:\Users\Ramir\Downloads\Base_alumnos 3.xlsx")
#xlsx_cpa_mat_diag=pd.read_excel(r"C:\Users\Ramir\Downloads\2025 CPA MAT COMP diag.xlsx")
#xlsx_cpa_mat_final=pd.read_excel(r"C:\Users\Ramir\Downloads\2025 CPA mat final.xlsx")




base_datos = None
xlsx_cpa_mat_diag = None

def cargar_excel():
    global base_datos

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
        base_datos = pd.read_excel(ruta_archivo)
        print(f"Archivo cargado: {ruta_archivo}")
        print(base_datos.head())  # Muestra las primeras filas
        return base_datos
    else:
        print("No se seleccionó ningún archivo.")
        return None

cargar_excel()

#Excel Diagnostico
def seleccionar_carpeta_diag(turno):
    """Abre el explorador de carpetas y devuelve la ruta seleccionada."""
    messagebox.showinfo(
        "Aviso",
        f"Abrir folder de tablas de respuestas del turno {turno}"
    )

    ruta_carpeta = filedialog.askdirectory(
        title=f"Selecciona la carpeta de respuestas diagnóstico ({turno})"
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
    """Carga archivos de diagnóstico del turno mañana."""
    ruta = seleccionar_carpeta_diag("mañana")
    return combinar_archivos_excel_diag(ruta, "mañana")


def cargar_excel_diag_tarde():
    """Carga archivos de diagnóstico del turno tarde."""
    ruta = seleccionar_carpeta_diag("tarde")
    return combinar_archivos_excel_diag(ruta, "tarde")


def exportar_excel_diag(turno="mañana"):
    """Exporta el DataFrame combinado del turno especificado."""
    if turno == "mañana":
        df = globals().get("base_datos_diag_m", None)
    else:
        df = globals().get("base_datos_diag_t", None)

    if df is not None:
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
            title=f"Guardar archivo combinado ({turno}) como"
        )
        if ruta_guardado:
            df.to_excel(ruta_guardado, index=False)
            print(f"✅ Archivo del turno {turno} guardado en: {ruta_guardado}")
    else:
        messagebox.showerror("Error", f"No hay datos cargados para el turno {turno}.")


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


def comparar_respuestas_y_calificar(df_diag, df_respuestas, turno):
    """
    Compara las respuestas del diagnóstico con las respuestas correctas.
    Devuelve un DataFrame con columnas originales y nuevas columnas de aciertos (1/0).
    """
    if df_diag is None or df_respuestas is None:
        messagebox.showerror("Error", f"Faltan archivos de diagnóstico o respuestas para {turno}.")
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
        messagebox.showerror("Error", "No hay datos cargados de diagnóstico para calificar.")
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
        ruta_m = os.path.join(ruta_guardado, "Diagnostico_Matutino_Calificado.xlsx")
        df_m_cal.to_excel(ruta_m, index=False)
        print(f"💾 Archivo guardado: {ruta_m}")

    if df_t_cal is not None:
        ruta_t = os.path.join(ruta_guardado, "Diagnostico_Vespertino_Calificado.xlsx")
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
        ruta_c = os.path.join(ruta_guardado, "Diagnostico_Combinado.xlsx")
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


def analizar_datos():
    """Crea una ventana nueva con las gráficas de análisis."""
    if excel_combinado is None:
        messagebox.showerror("Error", "Primero carga el archivo combinado para análisis.")
        return

    # Crear ventana de análisis
    ventana_analisis = tk.Toplevel()
    ventana_analisis.title("Analíticas del Examen Diagnóstico")
    ventana_analisis.geometry("1280x800")

    frame_scroll = ctk.CTkScrollableFrame(ventana_analisis, label_text="Resultados Analíticos")
    frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    # --- 1️⃣ Gráfica de errores por pregunta ---
    preguntas_correctas = [col for col in excel_combinado.columns if col.endswith("_correcta")]
    if not preguntas_correctas:
        messagebox.showerror("Error", "No se encontraron columnas de respuestas correctas en el archivo.")
        return

    errores = 1 - excel_combinado[preguntas_correctas].mean()
    plt.figure(figsize=(10, 6))
    plt.scatter(range(1, len(errores) + 1), errores.values * 100)
    plt.title("Porcentaje de errores por pregunta")
    plt.xlabel("Pregunta")
    plt.ylabel("% de errores")
    plt.xticks(range(1, len(errores) + 1), [f"P{i}" for i in range(1, len(errores) + 1)])
    plt.grid(True)
    plt.tight_layout()

    canvas1 = FigureCanvasTkAgg(plt.gcf(), master=frame_scroll)
    canvas1.draw()
    canvas1.get_tk_widget().pack(pady=20)

    # --- 2️⃣ Promedios por carrera ---
    if "carrera" in excel_combinado.columns and "CALIFICACION DIAG" in excel_combinado.columns:
        promedio_carrera = excel_combinado.groupby("carrera")["CALIFICACION DIAG"].mean().sort_values(ascending=False)
        plt.figure(figsize=(10, 6))
        promedio_carrera.plot(kind="bar", color="skyblue")
        plt.title("Promedio de calificación diagnóstica por carrera")
        plt.xlabel("Carrera")
        plt.ylabel("Promedio")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()

        canvas2 = FigureCanvasTkAgg(plt.gcf(), master=frame_scroll)
        canvas2.draw()
        canvas2.get_tk_widget().pack(pady=20)

    # --- 3️⃣ Aprobados vs reprobados ---
    if "CALIFICACION DIAG" in excel_combinado.columns:
        aprobados = (excel_combinado["CALIFICACION DIAG"] >= 60).sum()
        reprobados = (excel_combinado["CALIFICACION DIAG"] < 60).sum()
        df_aprob = pd.DataFrame({
            "Resultado": ["Aprobado", "Reprobado"],
            "Cantidad": [aprobados, reprobados]
        })

        plt.figure(figsize=(6, 6))
        sns.barplot(data=df_aprob, x="Resultado", y="Cantidad", palette="pastel")
        plt.title("Total de exámenes aprobados y reprobados")
        plt.xlabel("")
        plt.ylabel("Número de alumnos")
        plt.tight_layout()

        canvas3 = FigureCanvasTkAgg(plt.gcf(), master=frame_scroll)
        canvas3.draw()
        canvas3.get_tk_widget().pack(pady=20)

    # Mostrar mensaje final
    label_info = ctk.CTkLabel(frame_scroll, text="Analítica generada correctamente ✅", text_color="green")
    label_info.pack(pady=10)

#num_ficha = xlsx_datos_alumnos["ficha"].drop_duplicates().reset_index(drop=True)
#preguntas_diag_df = xlsx_cpa_mat_diag[["P1", "P2", "P3", "P4", "P5" 
#                                       , "P6", "P7", "P8", "P9", "P10" 
#                                       , "P11", "P12", "P13", "P14", "P15" 
#                                       , "P16", "P17", "P18", "P19", "P20" , "CALIF DIAG"]].reset_index(drop=True)
#examen_diag_df = pd.concat([num_ficha, preguntas_diag_df], axis=1)
#print(examen_diag_df)

"""
cols_a_combinar = ["NUMERO DE CONTROL.1", "Unnamed: 26", "Unnamed: 27", "Unnamed: 28"]  # o cualquier lista de columnas
xlsx_cpa_mat_final["NUMERO DE CONTROL"] = xlsx_cpa_mat_final[cols_a_combinar].astype(str).agg("".join, axis=1)

# Renombrar la columna en xlsx_cpa_mat_diag para que coincida con 'ficha'
xlsx_cpa_mat_final_renamed = xlsx_cpa_mat_final.rename(columns={"NUMERO DE CONTROL": "ficha"})

# Realizar el merge entre los dos DataFrames usando 'ficha' como clave

xlsx_datos_alumnos["ficha"] = xlsx_datos_alumnos["ficha"].astype(str)
xlsx_cpa_mat_final_renamed["ficha"] = xlsx_cpa_mat_final_renamed["ficha"].astype(str)
merged_final_df = pd.merge(xlsx_datos_alumnos, xlsx_cpa_mat_final_renamed, on="ficha", how="inner")

# Seleccionar solo las columnas de preguntas y calificación
preguntas_cols_f = [f"P_{i}" for i in range(1, 21)] + ["CALIFICACION"]
preguntas_final_df = merged_final_df[preguntas_cols_f]

# Combinar con los datos de ficha para formar el examen completo
examen_final_df = pd.concat([merged_final_df["ficha"].reset_index(drop=True), preguntas_final_df.reset_index(drop=True)], axis=1)

# Mostrar el resultado
print(examen_final_df)
"""
"""
COLUMNA = "carrera"
COLUMNA = "nombre_preparatoria"
COLUMNA = "escuela_procedencia"


carrera = base_datos["carrera"].drop_duplicates().reset_index(drop=True)

preparatoria_df = base_datos[["escuela_procedencia", "nombre_preparatoria"]].drop_duplicates().reset_index(drop=True)

carrera_df = pd.DataFrame({
    "id_carrera": range(1, len(carrera) + 1),
    "carrera": carrera
})


# Conteo completo
conteo_escuelas = base_datos['nombre_preparatoria'].value_counts()

# Eliminar una escuela específica
conteo_escuelas = conteo_escuelas.drop('OTRA', errors='ignore')

# Tomar las top 10 después de excluir
conteo_escuelas = conteo_escuelas.head(10)

#Truncar la longitud de los nombres
conteo_escuelas.index = conteo_escuelas.index.str.slice(0, 40)  # o str[:30]

print(conteo_escuelas)


plt.figure(figsize=(10, 6))
conteo_escuelas.plot(kind='bar', color='skyblue')

# Configuración del gráfico
plt.title('Número de Alumnos por Escuela de Procedencia')
plt.xlabel('Escuela de Procedencia')
plt.ylabel('Número de Alumnos')
plt.xticks(rotation=45, ha='right') # Rotar etiquetas para mejor lectura
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()

# Mostrar o guardar el gráfico
# plt.show() # Si estás en un entorno interactivo
plt.savefig('conteo_alumnos_por_escuela.png')
plt.show()
"""




