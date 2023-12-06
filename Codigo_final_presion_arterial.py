import time
import pandas as pd
from scipy.signal import find_peaks
import numpy as np
import serial

# Configuración del puerto serial y la velocidad de baudios
PUERTO = 'COM4'  # Ajusta este valor según el puerto en el que está conectado tu Arduino
BAUDIOS = 9600
ser = serial.Serial(PUERTO, BAUDIOS, timeout=1)

# DataFrame para almacenar los resultados
resultados = pd.DataFrame(columns=['Presión Sistólica', 'Presión Diastólica', 'Diagnóstico'])

# Funciones
def enviar_a_arduino(mensaje):
    ser.write(f"{mensaje}\n".encode())

def diagnostico_oms(sistolica, diastolica):
    if sistolica < 90 or diastolica < 60:
        return "Estado: hipotensión"
    elif 90 <= sistolica <= 120 and 60 <= diastolica <= 80:
        return "Estado: normal"
    else:
        return "Estado: hipertensión"

def agregar_resultado(sistolica, diastolica, diagnostico):
    global resultados
    index = len(resultados)
    resultados.loc[index] = [sistolica, diastolica, diagnostico]

def guardar_en_excel():
    ruta_escritorio = "D:/Semestre 8/VISION DE MAQUINA/BASE_DATOS_PRESION/resultados_presion_arterial.xlsx"
    resultados.to_excel(ruta_escritorio, index=False)

# Bucle Principal
try:
    while True:
        datos_presion = []
        tiempos = []
        estabilizado = False
        tiempo_inicio_estabilizacion = None
        inicio_medicion = time.time()

        enviar_a_arduino("BIENVENIDO :)")

        # Proceso de inflado
        while True:
            if ser.in_waiting:
                dato = ser.readline().decode('utf-8').strip()
                try:
                    valor = float(dato)
                    tiempo_actual = time.time() - inicio_medicion
                    datos_presion.append(valor)
                    tiempos.append(tiempo_actual)

                    if not estabilizado and valor >= 122.35 - 0.5:
                        if tiempo_inicio_estabilizacion is None:
                            tiempo_inicio_estabilizacion = tiempo_actual
                            enviar_a_arduino("Comenzando a inflar...")
                        elif tiempo_actual - tiempo_inicio_estabilizacion >= 2:
                            estabilizado = True
                            enviar_a_arduino("Valor máximo estabilizado.")

                    if estabilizado and valor <= 122.35 - 0.5:
                        break

                except ValueError as e:
                    print(f"Error al convertir el dato: {e}")

        # Proceso de desinflado
        enviar_a_arduino("Desinflando...")
        while True:
            if ser.in_waiting:
                dato = ser.readline().decode('utf-8').strip()
                try:
                    valor = float(dato)
                    tiempo_actual = time.time() - inicio_medicion
                    datos_presion.append(valor)
                    tiempos.append(tiempo_actual)

                    if valor <= 14:
                        enviar_a_arduino("Desinflado completado.")
                        break

                except ValueError as e:
                    print(f"Error al convertir el dato: {e}")

        # Procesamiento de los datos para la detección de picos
        derivada_presion = np.diff(datos_presion) / np.diff(tiempos)
        derivada_presion = np.insert(derivada_presion, 0, 0)
        picos_derivada, _ = find_peaks(abs(derivada_presion), height=np.std(derivada_presion))

        if tiempo_inicio_estabilizacion is not None:
            picos_derivada = [p for p in picos_derivada if tiempos[p] >= tiempo_inicio_estabilizacion + 2]

            if len(picos_derivada) > 0:
                # Procesamiento de los datos para la detección de picos
                # Calcula la derivada de la señal de presión dividiendo las diferencias entre los valores de presión por las diferencias de tiempo
                derivada_presion = np.diff(datos_presion) / np.diff(tiempos)

                # Inserta un valor cero al principio de la derivada para asegurarse de que tenga la misma longitud que los datos de presión
                derivada_presion = np.insert(derivada_presion, 0, 0)

                # Utiliza la función find_peaks de SciPy para detectar picos en el valor absoluto de la derivada de presión
                # Establece un umbral de altura basado en la desviación estándar de la derivada
                picos_derivada, _ = find_peaks(abs(derivada_presion), height=np.std(derivada_presion))

                # Verifica si ha ocurrido una estabilización de la señal antes de intentar detectar los picos
                if tiempo_inicio_estabilizacion is not None:
                    # Filtra los picos detectados para incluir solo aquellos que ocurrieron después de la estabilización y al menos 2 segundos después
                    picos_derivada = [p for p in picos_derivada if tiempos[p] >= tiempo_inicio_estabilizacion + 2]

                    # Comprueba si se detectaron picos válidos
                    if len(picos_derivada) > 0:
                        # Toma el primer pico detectado como el pico sistólico
                        pico_sistolico = picos_derivada[0]

                        # Calcula la presión sistólica tomando el valor de presión en el pico
                        presion_sistolica = round(datos_presion[pico_sistolico], 2)

                        # Calcula la presión diastólica asumiendo una relación fija con la sistólica
                        presion_diastolica = round(presion_sistolica * 0.67, 2)

                        # Determina el diagnóstico utilizando la función diagnostico_oms
                        diagnostico = diagnostico_oms(presion_sistolica, presion_diastolica)

                        # Crea un mensaje con los valores de presión y el diagnóstico
                        mensaje_presiones = f"{presion_sistolica} | {presion_diastolica} | {diagnostico}"

                        # Envía los resultados al Arduino
                        enviar_a_arduino(mensaje_presiones)

                        # Agrega los resultados al DataFrame de resultados
                        agregar_resultado(presion_sistolica, presion_diastolica, diagnostico)

                        # Guarda los resultados en un archivo Excel
                        guardar_en_excel()
            else:
                mensaje_error = "No se detectó presión sistólica"
                print(mensaje_error)
                enviar_a_arduino(mensaje_error)

except Exception as e:
    print(f"Se produjo un error: {e}")
finally:
    # Cierre del puerto serial al final
    ser.close()
