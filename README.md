# Documentación del Proyecto

## Resumen
Este proyecto es una aplicación de escritorio en Python (Tkinter + CustomTkinter) para:
- Cargar bases de datos de alumnos.
- Formatear resultados de exámenes (FormScanner/ZipGrade).
- Calificar exámenes diagnósticos y finales en lote.
- Combinar resultados con datos académicos.
- Generar análisis y gráficas comparativas.

Los dos archivos principales son:
- `main.py`: lógica de negocio, lectura/escritura de Excel y análisis.
- `ui.py`: interfaz gráfica principal, navegación entre páginas y conexión con `main.py`.

## Requisitos
- Python 3.10+
- Librerías:
  - `pandas`
  - `matplotlib`
  - `seaborn`
  - `customtkinter`
  - `openpyxl`
  - `pyodbc` (importado en `main.py`)

Instalación sugerida:

```bash
pip install pandas matplotlib seaborn customtkinter openpyxl pyodbc
```

## Ejecución
Desde la raíz del proyecto:

```bash
python ui.py
```

## Arquitectura General
- `ui.py` crea la ventana principal, la barra de navegación superior y las páginas:
  - Inicio
  - Formato
  - Gráficas
  - Calificar
  - Análisis
- Cada botón de interfaz llama funciones de `main.py` para procesar archivos y generar resultados.
- Los datos se mantienen en variables globales dentro de `main.py` (por ejemplo `base_datos`, `respuestas_m`, `respuestas_t`, `excel_combinado`).

## Documentación de `main.py`

### Rol del módulo
Centraliza el procesamiento de datos: carga de archivos, calificación por respuesta correcta, combinación de datasets, análisis estadístico y exportación de resultados.

### Variables globales importantes
- `base_datos`: base general de alumnos.
- `base_datos_diag_m`, `base_datos_diag_t`: respuestas del diagnóstico por turno.
- `respuestas_m`, `respuestas_t`: plantillas de respuestas correctas por turno.
- `excel_combinado`: archivo consolidado para análisis.

### Funciones principales (por bloques)

#### 1) Carga de datos
- `cargar_excel_base()`
  - Carga la base general de alumnos desde Excel.
- `cargar_excel_diag_mañana()` / `cargar_excel_diag_tarde()`
  - Seleccionan carpeta y consolidan archivos de respuestas por turno.
- `combinar_archivos_excel_diag(ruta_carpeta, turno)`
  - Une múltiples Excel de una carpeta en un solo DataFrame y agrega columnas de trazabilidad (`archivo_origen`, `turno`).
- `cargar_excel_respuestas()`
  - Carga archivo de respuestas correctas desde Excel.

#### 2) Calificación
- `comparar_respuestas_y_calificar(df_diag, df_respuestas, turno)`
  - Compara columnas `P1..P20` contra la fila de respuestas correctas.
  - Crea columnas `P1_correcta..P20_correcta` con valores binarios (`1` acierto, `0` error).
- `generar_excels_calificados(df_m, df_t, df_resp_m, df_resp_t)`
  - Califica por turno.
  - Calcula `CALIF DIAG` como suma de aciertos por 5.
  - Exporta:
    - Examen matutino calificado.
    - Examen vespertino calificado.
    - Examen combinado.

#### 3) Análisis y visualización
- `cargar_excel_analitico()`
  - Carga un Excel combinado para análisis.
- `analizar_datos(estado='aciertos')`
  - Gráfica porcentaje por pregunta (`P1_correcta..P20_correcta`), con filtros por `carrera` y `ppgrupo`.
- `analizar_datos2(modo='aciertos')`
  - Compara dos archivos calificados por grupo.
- `comparar_reprobados()`
  - Contrasta alumnos reprobados entre dos archivos.
- `comparar_porcentajes()`
  - Compara porcentaje de aprobados/reprobados entre archivos.
- `comparar_promedio_total()`
  - Compara promedio global entre dos archivos.
- `comparar_promedios_por_carrera_dos_archivos()`
  - Compara promedio por carrera en dos datasets.
- `comparar_promedio_final_por_carrera()`
  - Muestra promedio final por carrera para un archivo.

#### 4) Combinación de datasets
- `combinar_diag_con_final()`
  - Une diagnóstico + final por columnas clave:
    - `NUMERO DE CONTROL`, `carrera`, `ppgrupo`, `NOMBRE COMPLETO`, `turno`.
  - Renombra columnas no clave con prefijos `DIAG_` y `FINAL_` para evitar colisiones.
- `combinar_combinado_completo_con_base_datos()`
  - Une el combinado (diag+final) con la base general de alumnos usando `NUMERO DE CONTROL`.

#### 5) Formateadores
- `cargar_formateador(tipo_form)`
  - Ejecuta flujo repetible para transformar archivos de origen (CSV/Excel) a un formato estándar.
  - Normaliza columnas, crea/ajusta campos (`carrera`, `ppgrupo`, `NUMERO DE CONTROL`, `P1..P20`) y exporta a Excel.

## Documentación de `ui.py`

### Rol del módulo
Construye la interfaz principal, administra navegación y ejecuta acciones de usuario invocando funciones en `main.py`.

### Flujo de UI
- Inicializa tema y ventana con `customtkinter`.
- Crea topbar con pestañas:
  - `Inicio`: carga base de datos principal.
  - `Formato`: acceso al flujo de formateo (FormScanner/ZipGrade) y ejecución batch de FormScanner.
  - `Graficas`: genera gráfica de conteo para una columna de la base cargada y permite exportar PNG.
  - `Calificar`: carga exámenes por turno, carga respuestas correctas y genera archivos calificados/combinados.
  - `Analisis`: acceso a todas las comparativas estadísticas y gráficas avanzadas.

### Funciones clave
- `show_page(page_name)`
  - Renderiza dinámicamente cada página del sistema.
- `mostrar_grafico(nombre_columna, error_label, frame_grafico)`
  - Genera gráfico de barras por columna seleccionada.
- `exportar_grafico(error_label_exportar)`
  - Guarda la gráfica actual en PNG.
- `actualizar_columnas()`
  - Lanza carga de Excel base.
- `generar_archivos_calificados()`
  - Wrapper de seguridad para llamar `main.generar_excels_calificados(...)`.

## Formato de Datos Esperado
Para evitar errores, los archivos deben incluir al menos:
- Preguntas: `P1` a `P20`.
- Identificadores frecuentes:
  - `NUMERO DE CONTROL`
  - `carrera`
  - `ppgrupo`
  - `NOMBRE COMPLETO`
  - `turno`
- En archivos calificados:
  - `CALIF DIAG`
  - `P1_correcta` ... `P20_correcta`

## Flujo recomendado de uso
1. Abrir aplicación (`python ui.py`).
2. En `Inicio`, cargar base de alumnos.
3. En `Formato`, convertir archivos crudos al formato estándar si aplica.
4. En `Calificar`:
   - Cargar carpetas de respuestas (mañana/tarde).
   - Cargar plantillas de respuestas correctas.
   - Generar archivos calificados y combinado.
5. En `Calificar`, combinar diagnóstico y final si se requiere.
6. En `Analisis`, cargar combinado y generar gráficas/comparativas.

## Notas técnicas
- La aplicación depende de selección manual de rutas con diálogos de archivos.
- La mayoría de procesos operan en memoria con `pandas` y se exportan a Excel.
- El proyecto usa estado global en `main.py`; si se reinicia la app, se pierde el estado cargado.

## Estructura mínima actual
- `main.py`
- `ui.py`
- `formscanner_batch.py`
- `benchmark.py`
