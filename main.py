import pyodbc
import customtkinter
import pandas as pd


xlsx_datos_alumnos=pd.read_excel(r"C:\Users\hankr\Downloads\Copia de BASE DE DATOS GENERACIÓN 2025.xlsx")

print(xlsx_datos_alumnos.head())