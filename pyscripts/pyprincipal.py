import serial
import time
import re
import manejo_excel 
#si no tiene las librerias instaladas:
# pip install openpyxl pyserial


nombre_archivo = "seminario1.xlsx"


puerto_serial = 'COM7'
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

       
        while True:

            if ser.in_waiting > 0:
                linea = ser.readline().decode('utf-8').strip() # Lee, decodifica y limpia la línea
                if linea: # Verifica que la línea no esté vacía
                    print(f"Dato recibido: {linea}")

                    # Intenta extraer los valores usando la expresión regular
                    # Formato: fecha y hora:DD/MM/AAAA HH:MM:SS,peso:X.XX
                    # ([\d.]+) captura números (enteros o flotantes)
                    # ([\d/]+) captura la fecha (formato DD/MM/AAAA)
                    # ([\d:]+) captura la hora (formato HH:MM:SS)
                    # ([\w.]+) captura palabras (letras, números, guiones bajos y puntos)
                    match = re.match(r'Fecha y hora:([\d/]+) ([\d:]+), Peso:([\d.]+)', linea)

                    if match:
                        try:
                            peso_val = float(match.group(3)) 

                            fecha_val = match.group(1)  # Captura la fecha
                            hora_val = match.group(2)    # Captura la hora
                            # Prepara el arreglo de datos para el Excel
                            datos_para_excel = [fecha_val, hora_val, peso_val]

                            # Llama a la función del módulo manejo_excel para guardar
                            manejo_excel.agregar_datos_a_excel(archivo_excel, datos_para_excel)


                        except ValueError as ve:
                            print(f"  Error al parsear valores numéricos o de formato: {ve}. Línea: '{linea}'")
                        except Exception as excel_err:
                            print(f"  Error al escribir en Excel: {excel_err}")
                    else:
                        print(f"  Formato de datos no reconocido. Se esperaba 'fecha y hora: DD/MM/AAAA HH:MM:SS, peso:X.XX'. Línea: '{linea}'")


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
