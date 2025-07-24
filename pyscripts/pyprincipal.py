
from datetime import datetime
import serial
import time
import re
import manejoexcel # crear un archivo manejoexcel.py con la función agregar_datos_a_excel. Debe estar en el mismo directorio que este script.

#si no tiene las librerias instaladas:
# pip install openpyxl pyserial


nombre_archivo = "seminario1.xlsx"


puerto_serial = 'COM4'
velocidad_baudios = 115200

def leer_y_procesar_datos_serial(puerto, baudios, archivo_excel):
    
    """
    Args:
        puerto (str): El nombre del puerto serial.
        baudios (int): La velocidad de comunicación (baudios).
        archivo_excel (str): El nombre del archivo Excel para guardar los datos.
    """
    ser = None
    try:

        ser = serial.Serial(puerto, baudios, timeout=1)
        print(f"Abriendo puerto serial {puerto} a {baudios} baudios...")
        time.sleep(2) 

        print("Esperando datos... Presiona Ctrl+C para detener.")
        leer_op = input("¿Desea leer datos del puerto serial? (s/n): ")

        while leer_op.lower() == 's':
            if ser.in_waiting > 0:
                linea = ser.readline().decode('utf-8').strip() # Lee, decodifica y limpia la línea
                if linea: # Verifica que la línea no esté vacía
                    print(f"Dato recibido: {linea}")

                    # Intenta extraer los valores usando la expresión regular
                    # regex: 'kilos:([\d.]+),material:([\w.]+)'
                    # ([\d.]+) captura números (enteros o flotantes)
                    # ([\w.]+) captura palabras (letras, números, guiones bajos y puntos)
                    match = re.match(r'peso:([\d.]+)', linea)

                    if match:
                        try:
                            peso_val = float(match.group(1)) 

                            fecha_hora_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                            # Prepara el arreglo de datos para el Excel
                            datos_para_excel = [fecha_hora_actual, peso_val]

                            # Llama a la función del módulo manejoexcel para guardar
                            manejoexcel.agregar_datos_a_excel(archivo_excel, datos_para_excel)

                        except ValueError as ve:
                            print(f"  Error al parsear valores numéricos o de formato: {ve}. Línea: '{linea}'")
                        except Exception as excel_err:
                            print(f"  Error al escribir en Excel: {excel_err}")
                    else:
                        print(f"  Formato de datos no reconocido. Se esperaba 'peso:X.XX'. Línea: '{linea}'")

    except serial.SerialException as e:
        print(f"Error al abrir o comunicar con el puerto serial: {e}")
        print("Asegúrate de que la ESP32 esté conectada y el puerto sea correcto.")
        print("Verifica que el puerto serial no esté siendo utilizado por otro programa (ej. Monitor Serial de Arduino IDE).")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
    except KeyboardInterrupt:
        print("\nLectura de datos seriales detenida por el usuario.")
    finally:
        if ser is not None and ser.is_open:
            ser.close()
            print("Puerto serial cerrado.")

if __name__ == "__main__":
    leer_y_procesar_datos_serial(puerto_serial, velocidad_baudios, nombre_archivo)
