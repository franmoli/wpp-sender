import sys
import requests
import time
import webbrowser
import subprocess
import pygetwindow as gw
import pyautogui
import pyperclip
import pandas as pd
from datetime import datetime
import os
import keyboard

FILES_DIRECTORY = "\\archivos"

def send_wpp(numero, mensaje, archivos = None):

    count = 0
    numeroIncorrecto = False
    pudoCopiar = False

    while count <= 5 and not numeroIncorrecto and not pudoCopiar:
        # abrir el chat que corresponde y pegar el numero
        pudoAbrir = abrirTelefono(numero)
        if not pudoAbrir:
            print("No lo pude abrir")
            break

        # copiar mensaje y comprobar si esta correcto
        pudoCopiar = copiar_mensaje(mensaje)

        if not pudoCopiar:
            numerocorrecto = testear_numero_correcto()
            if numerocorrecto:
                continue
        
        if archivos is not None and pudoCopiar:
            copiar_archivo(archivos)
        
        count += 1

    if pudoCopiar:
        print("Envio correcto a :" + numero)
    else:
        loguearFallo(numero)
        
    enter()
    borrarTodo()

    cerrarVentanaWpp()


def copiar_archivo(archivos):
    set_clipboard_files(archivos)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.sleep(3)

def copiar_mensaje(mensaje):
    
    # seteo el numero de reintentos para copiar el mensaje
    count = 0
    comprobado = False

    while count <= 3 and not comprobado:
        # focusear_chat()
        pyperclip.copy(mensaje)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(.5)
        count += 1
        comprobado = testear_mensaje_copiado(mensaje)

    # lo pudo copiar o no
    if count <= 3:
        return True
    else:
        return False

def enter():
    pyautogui.press('enter')

def testear_numero_correcto():
    print("Testeando si el numero es correcto")
    pyautogui.hotkey('ctrl', '3')
    return copiar_mensaje("Test")

def testear_mensaje_copiado(mensaje):
    pyperclip.copy("Reset")
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    texto_pegado = pyperclip.paste().replace('\r\n', '\n').replace('\r', '\n')
    
    if(mensaje != texto_pegado):
        print("Error en el copiado del mensaje")
        pyautogui.hotkey('ctrl', 'x')
        return False
    
    return True

def set_clipboard_files(files):
    for archivo in files:
         fullpath = os.path.join(files_directory(), archivo)
         os.system(f"""powershell Set-Clipboard -Path '{fullpath}' """)

def files_directory():
    # Obtener la ruta completa del ejecutable
    ruta_ejecutable = sys.executable
    # Obtener el directorio del ejecutable
    if "Python" in ruta_ejecutable:
        return "." + FILES_DIRECTORY
    return os.path.dirname(ruta_ejecutable) + FILES_DIRECTORY

def cerrarVentanaWpp():
    
    # cerrar navegador y seguir 
    windows = gw.getWindowsWithTitle('chrome')  

    if windows:
        window = windows[0]  # Tomar la primera ventana que coincida
        window.activate()    # Enfocar la ventana
        time.sleep(1)        # Esperar un poco antes de enviar la combinación de teclas

        pyautogui.hotkey('ctrl', 'w') 
    else:
        print("No se encontró la ventana")

def loguearFallo(numero):
    print("No le pude mandar a: " + numero)
    # Obtener la fecha actual
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    # Nombre del archivo de log basado en la fecha actual
    nombre_archivo = f"./errores/log_{fecha_actual}.txt"
    
    # Formatear el mensaje de error con la hora actual
    hora_actual = datetime.now().strftime("%H:%M:%S")
    mensaje_formateado = f"{hora_actual} - ERROR: No se pudo enviar al numero: {numero}\n"
    
    # Escribir el mensaje en el archivo (crear o añadir)
    with open(nombre_archivo, 'a') as archivo_log:
        archivo_log.write(mensaje_formateado)

def abrirWpp():
    abrirTelefono("")

    count = 0
    comprobado = False

    while count <= 5 and not comprobado:
        comprobado = testear_numero_correcto()
        count += 1
    
    if comprobado:
        borrarTodo()
        print(f"Whatsapp comprobado correctamente {count}")
    else:
        print("Fallo al abrir whatsapp")
        return False

    cerrarVentanaWpp()
    return True

def abrirTelefono(telefono):
    url = "http://wa.me/" + telefono
    webbrowser.open(url, new=1)
    time.sleep(3)
    # checkear que no estoy en pagina de 404
    borrarTodo()
    pyperclip.copy("Reset")
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    return pyperclip.paste().find("404") == -1
    

def borrarTodo():
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')