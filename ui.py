import pyodbc
import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Cargar base de datos
base_datos = pd.read_excel(r"C:\Users\hankr\Downloads\Base_alumnos 3.xlsx")

# Configurar apariencia
ctk.set_appearance_mode("dark")  # "light", "dark", "system"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

# Crear ventana principal
ventana = ctk.CTk()
ventana.title("Graficador de Calificaciones")
ventana.geometry("1280x720")

# --- Función para mostrar gráfico dinámico ---
def mostrar_grafico(nombre_columna):
    # Validar si la columna existe
    if nombre_columna not in base_datos.columns:
        error_label.configure(text=f"⚠️ La columna '{nombre_columna}' no existe.", text_color="red")
        return
    else:
        error_label.configure(text="")  # limpiar mensaje de error
    
    # Obtener el conteo de valores únicos
    conteo = base_datos[nombre_columna].value_counts().head(20)
    
    # Crear figura
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar(conteo.index.astype(str), conteo.values, color='skyblue')
    
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

# Widgets de la interfaz
label_instruccion = ctk.CTkLabel(ventana, text="Ingrese el nombre de la columna a graficar:")
label_instruccion.pack(pady=10)

entry_columna = ctk.CTkEntry(ventana, width=300, placeholder_text="Ejemplo: nombre_preparatoria")
entry_columna.pack(pady=5)

boton = ctk.CTkButton(ventana, text="Mostrar gráfico", command=generar_grafico)
boton.pack(pady=10)

error_label = ctk.CTkLabel(ventana, text="", text_color="red")
error_label.pack()

frame_grafico = ctk.CTkFrame(ventana)
frame_grafico.pack(pady=20, fill="both", expand=True)

# Ejecutar ventana
ventana.mainloop()
