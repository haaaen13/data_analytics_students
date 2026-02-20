import main
import time



def prueba_completa():
    print("Cargando base...")
    main.cargar_excel_base()

    print("Seleccionando la carpeta de diagnostico")
    main.cargar_excel_diag_mañana() #problema: hay que llamar otra funcion para que esta funcione


    print("Analizando datos...")
    main.analizar_datos()

    print("Analizando errores...")
    main.analizar_datos_errores()

    print("Comparando...")
    main.comparar_porcentajes()

if __name__ == "__main__":
    inicio = time.perf_counter()

    prueba_completa()

    fin = time.perf_counter()
    print(f"\nTiempo total: {fin - inicio:.3f} segundos")