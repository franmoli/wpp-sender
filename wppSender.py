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
from sendMessage import send_wpp, abrirWpp

REPO = "franmoli/wpp-sender"
VERSION = "0.2.6-alpha"
COLUMN_USER = "Usuario"
COLUMN_TELEPHONE = "Telefono"
COLUMN_MESSAGE = "Mensaje"
COLUMN_FILES = "Archivos"
EXCEL_DIRECTORY = "\\excel"

debug = False
archivos_excel = []
interrupcion = False

def excel_directory():
    # Obtener la ruta completa del ejecutable
    ruta_ejecutable = sys.executable
    # Obtener el directorio del ejecutable
    if "Python" in ruta_ejecutable:
        return "." + EXCEL_DIRECTORY
    return os.path.dirname(ruta_ejecutable) + EXCEL_DIRECTORY

def interrumpir():
    global interrupcion
    interrupcion = True

def traer_datos(id):
    # URL de la API
    url = 'http://develop.amaip.com.ar/api/get-wpp-msj-list.php?mensaje='

    print(id)

    # Datos que quieras enviar en el POST, usualmente como un diccionario
    

    # Envío de la solicitud POST
    response = requests.get(url + id)

    # Verifica que la solicitud fue exitosa
    if response.status_code == 200:
        # Procesa la respuesta
        return response.json()
        
    else:
        print("Error en la solicitud:", response.status_code)

def parse_url(url):
    url = ''.join(url)
    # Quitamos el prefijo 'wppsender://' si está presente
    if url.startswith('wppsender://'):
        url = url[len('wppsender://'):]
    
    # Dividimos el resto de la cadena en cada '/'
    arguments = url.split('/')
    
    return arguments

def send_list(list_id):
    print("Tengo que mandar a la lista numero " + list_id)
    datos = traer_datos(list_id)
    user_list = datos['user_list']
    mensaje = datos['mensaje']
    print("El mensaje recibido es: " + mensaje)
    # formato usuario [user_id, telefono, nombre, apellido]
    for usuario in user_list:
        print(usuario)
        # formateo el mensaje
        mensaje_formateado = mensaje.replace("{{NOMBRE}}", usuario[2]).replace("{{APELLIDO}}", usuario[3]) 
        # envio un mensaje
        send_wpp(usuario[1], mensaje_formateado)
        
def get_latest_release(repo):
    url = f'https://api.github.com/repos/{repo}/releases/latest'
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        version = data['tag_name']
        download_url = data['assets'][0]['browser_download_url']
        return version, download_url
    else:
        raise Exception(f"Error fetching release information: {response.status_code}")
    
def check_for_updates(current_version, repo):
    try:
        latest_version, download_url = get_latest_release(repo)
        if latest_version > current_version:
            return download_url
    except Exception as e:
        print(e)
    return None

def send_excel():
    global archivos_excel
    global debug
    print("Enviando mensajes...")
    pudoAbrir = False
    count = 0;
    while count <= 5 and not pudoAbrir:
        print("Iniciando wpp")
        pudoAbrir = abrirWpp()
        count += 1

        
    for archivo in archivos_excel:
        try:
            # Leer el archivo Excel
            df = pd.read_excel(excel_directory() + "/" + archivo, dtype=str)
            
            # Verificar si las columnas existen
            if f'{COLUMN_USER}' in df.columns and f'{COLUMN_TELEPHONE}' in df.columns and f'{COLUMN_MESSAGE}' in df.columns and f'{COLUMN_FILES}' in df.columns:
                usuarios = df[f'{COLUMN_USER}'].tolist()
                telefonos = df[f'{COLUMN_TELEPHONE}'].tolist()
                mensajes = df[f'{COLUMN_MESSAGE}'].tolist()
                archivos = df[f'{COLUMN_FILES}'].tolist()

                # Mostrar los nombres de usuario y sus teléfonos
                for usuario, telefono, mensaje, archivo in zip(usuarios, telefonos, mensajes, archivos):
                    
                    if interrupcion:
                        return
                    if pd.isna(telefono):
                        continue
                    if pd.isna(usuario):
                        usuario = ''
                    if pd.isna(mensaje):
                        mensaje = ''
                    if pd.isna(archivo):
                        archivo = ''
                    print(f"Enviando a Usuario: {usuario}, Telefono: {telefono}")
                    mensaje = mensaje.replace('{usuario}', usuario)
                    send_wpp(f"{telefono}", f"{mensaje}", parseFiles(archivo), debug)
            else:
                print(f"El archivo excel esta en un formato incorrecto.")
        except FileNotFoundError:
            print(f"El archivo '{archivo}' no se encontró.")
        except Exception as e:
            print(f"Ocurrió un error: {e}")

def autoupdate():
    # traer ultima version
    update_url = check_for_updates(VERSION,REPO)

    if update_url:
        print("Nueva versión disponible. Descargando e instalando actualización...")
        # Llama al script de actualización
        subprocess.Popen([ './utils/wppSenderUpdater.exe', update_url, sys.argv[0]])
        sys.exit()
        
    else:
        print("No hay actualizaciones disponibles.") 

def art():
    print("""
                                _____                __         
            _      ______  ____ / ___/___  ____  ____/ /__  _____
            | | /| / / __ \/ __ \ __ \/ _ \/ __ \/ __  / _ \/ ___/
            | |/ |/ / /_/ / /_/ /__/ /  __/ / / / /_/ /  __/ /    
            |__/|__/ .___/ .___/____/\___/_/ /_/\__,_/\___/_/     
                  /_/   /_/                                       

    """)
    print(f'version v{VERSION}')

def parseFiles(archivo):
    if archivo == "" or archivo == None or pd.isna(archivo):
        return None
    return [archivo]

def listar_archivos(carpeta):
    lista_archivos = []
    for nombre_archivo in os.listdir(carpeta):
        ruta_archivo = os.path.join(carpeta, nombre_archivo)
        if os.path.isfile(ruta_archivo):
            lista_archivos.append(nombre_archivo)
    return lista_archivos

def mostrarMenu():
    global debug
    global archivos_excel
    print(f""" 
        Listas cargadas:
            {archivos_excel}
        1-Comenzar envío 
        2-Refrescar listas
        3-Comprobar whatsapp instalado
        4-Salir
        5-Modo de prueba = {debug}
            Tip: para cortar el envio podes presionar 'q'
        """)

def goToMenuPrincipal():
    # os.system('cls')
    art()
    mostrarMenu()

def refrescarLista():
    global archivos_excel
    archivos_excel = listar_archivos(excel_directory())

def listarNumeros():
    global archivos_excel
    print("Encontraste una funcion oculta :)")
    print(archivos_excel)
    for archivo in archivos_excel:
        print(archivo)
        try:
            # Leer el archivo Excel
            df = pd.read_excel(excel_directory() + "/" + archivo, dtype=str)
            
            # Verificar si las columnas existen
            if f'{COLUMN_USER}' in df.columns and f'{COLUMN_TELEPHONE}' in df.columns and f'{COLUMN_MESSAGE}' in df.columns and f'{COLUMN_FILES}' in df.columns:
                usuarios = df[f'{COLUMN_USER}'].tolist()
                telefonos = df[f'{COLUMN_TELEPHONE}'].tolist()
                mensajes = df[f'{COLUMN_MESSAGE}'].tolist()
                archivos = df[f'{COLUMN_FILES}'].tolist()

                # Mostrar los nombres de usuario y sus teléfonos
                for usuario, telefono, mensaje, archivo in zip(usuarios, telefonos, mensajes, archivos):
                    
                    if interrupcion:
                        return
                    if pd.isna(telefono):
                        continue
                    print(f"Enviando a Usuario: {usuario}, Telefono: {telefono}")
            else:
                print(f"Las columnas '{COLUMN_USER}' y/o '{COLUMN_TELEPHONE}' no se encontraron en el archivo Excel.")
        except FileNotFoundError:
            print(f"El archivo '{archivo}' no se encontró.")
        except Exception as e:
            print(f"Ocurrió un error: {e}")

def main():

    global archivos_excel
    global debug
    # Configurar la hotkey
    keyboard.add_hotkey('q', interrumpir)

    art()
    autoupdate()


    args = parse_url(sys.argv[1:])

    if(args[0] == 'send_list'):
        send_list(args[1])
    
    archivos_excel = listar_archivos(excel_directory())

    mostrarMenu()

    while True:
        seleccion = input()
        print(f"presionaste {seleccion}")
        if seleccion == "1":
            send_excel()
        if seleccion == "2":
            refrescarLista()
        if seleccion == "3":
            abrirWpp()
        if seleccion == "4":
            break
        if seleccion == "5":
            debug = not debug
        if seleccion == "6":
            listarNumeros()
        
        goToMenuPrincipal()

    print("Adios.")


if __name__ == "__main__":
    main()