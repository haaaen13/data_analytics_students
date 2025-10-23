import pyodbc
import customtkinter
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np 

base_datos=pd.read_excel(r"C:\Users\hankr\Downloads\Base_alumnos 3.xlsx")








carrera = base_datos["carrera"].drop_duplicates().reset_index(drop=True)
carrera_df = pd.DataFrame({"id_carrera": range(1, len(carrera) + 1),"carrera": carrera
})


preparatoria_df = base_datos[["escuela_procedencia", "nombre_preparatoria"]].drop_duplicates().reset_index(drop=True)





#print(preparatoria_df)



#Se crea la tabla alumnos con indice no_control
alumnos_df=base_datos[["no_de_control","nombre_alumno", "sexo"]]
#alumnos_df = alumnos_df.set_index("no_de_control")

alumnos_df['id_carrera_fk'] = np.nan


#print(alumnos_df)

xlsx_cpa_mat_diag=pd.read_excel(r"C:\Users\hankr\Downloads\2025 CPA MAT COMP diag.xlsx")



# Renombrar la columna en xlsx_cpa_mat_diag para que coincida con 'ficha'
xlsx_cpa_mat_diag_renamed = xlsx_cpa_mat_diag.rename(columns={"NUMERO DE CONTROL": "ficha"})

# Realizar el merge entre los dos DataFrames usando 'ficha' como clave
merged_diag_df = pd.merge(base_datos, xlsx_cpa_mat_diag_renamed, on="ficha", how="inner")

# Seleccionar solo las columnas de preguntas y calificación
preguntas_cols = [f"P{i}" for i in range(1, 21)] + ["CALIF DIAG"]
preguntas_diag_df = merged_diag_df[preguntas_cols]

# Combinar con los datos de ficha para formar el examen completo
examen_diag_df = pd.concat([merged_diag_df["ficha"].reset_index(drop=True), preguntas_diag_df.reset_index(drop=True)], axis=1)

# Mostrar el resultado
#print(examen_diag_df)




xlsx_cpa_mat_final=pd.read_excel(r"C:\Users\hankr\Downloads\2025 CPA mat final.xlsx")





cols_a_combinar = ["NUMERO DE CONTROL.1", "Unnamed: 26", "Unnamed: 27", "Unnamed: 28"]  # o cualquier lista de columnas
xlsx_cpa_mat_final["NUMERO DE CONTROL"] = xlsx_cpa_mat_final[cols_a_combinar].astype(str).agg("".join, axis=1)

# Renombrar la columna en xlsx_cpa_mat_diag para que coincida con 'ficha'
xlsx_cpa_mat_final_renamed = xlsx_cpa_mat_final.rename(columns={"NUMERO DE CONTROL": "ficha"})

# Realizar el merge entre los dos DataFrames usando 'ficha' como clave

base_datos["ficha"] = base_datos["ficha"].astype(str)
xlsx_cpa_mat_final_renamed["ficha"] = xlsx_cpa_mat_final_renamed["ficha"].astype(str)
merged_final_df = pd.merge(base_datos, xlsx_cpa_mat_final_renamed, on="ficha", how="inner")

# Seleccionar solo las columnas de preguntas y calificación
preguntas_cols_f = [f"P_{i}" for i in range(1, 21)] + ["CALIFICACION"]
preguntas_final_df = merged_final_df[preguntas_cols_f]

# Combinar con los datos de ficha para formar el examen completo
examen_final_df = pd.concat([merged_final_df["ficha"].reset_index(drop=True), preguntas_final_df.reset_index(drop=True)], axis=1)

# Mostrar el resultado
print("final")
print(examen_final_df)



conteo_escuelas=base_datos['nombre_preparatoria'].value_counts()

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









