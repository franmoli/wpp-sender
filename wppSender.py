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

REPO = "franmoli/wpp-sender"
VERSION = "0.2.1-alpha"
COLUMN_USER = "Usuario"
COLUMN_TELEPHONE = "Telefono"
COLUMN_MESSAGE = "Mensaje"
EXCEL_DIRECTORY = "\\excel"
archivos_excel = []

def excel_directory():
    # Obtener la ruta completa del ejecutable
    ruta_ejecutable = sys.executable
    # Obtener el directorio del ejecutable
    if "Python" in ruta_ejecutable:
        return "." + EXCEL_DIRECTORY
    return os.path.dirname(ruta_ejecutable) + EXCEL_DIRECTORY

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
        
def send_wpp(numero, mensaje):
    url = "http://wa.me/" + numero
    print(str(url))
   # Abrir la URL en un navegador web
    webbrowser.open(url, new=1)

    # Esperar unos segundos
    time.sleep(4)  
    
    # copiar mensaje y comprobar si esta correcto
    response = copiar_mensaje(mensaje)

    if not response:
        loguearFallo(numero)

    cerrarVentanaWpp()

def loguearFallo(numero):
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

def copiar_mensaje(mensaje):
    
    # seteo el numero de reintentos para copiar el mensaje
    count = 0

    comprobado = False

    while count <= 5 and not comprobado:
        # focusear_chat()
        pyperclip.copy(mensaje)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        count += 1
        comprobado = testear_mensaje_copiado(mensaje)

    # envio mensaje
    if count <= 5:
        pyautogui.press('enter')
        return True
    else:
        pyautogui.press('enter')
        return False

def testear_mensaje_copiado(mensaje):
    pyperclip.copy("Reset")
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    texto_pegado = pyperclip.paste().replace('\r\n', '\n').replace('\r', '\n')
    
    if(mensaje != texto_pegado):
        print("Original: " + mensaje)
        print("Copia: " + texto_pegado)
        pyautogui.hotkey('ctrl', 'x')
        return False
    
    return True

def focusear_chat():
    pyautogui.hotkey('ctrl','shift', 'f')
    pyautogui.press('escape')
    
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
    print("Enviando mensajes...")
    abrirWpp()
    for archivo in archivos_excel:
        try:
            # Leer el archivo Excel
            df = pd.read_excel("./excel/" + archivo)
            
            # Verificar si las columnas existen
            if f'{COLUMN_USER}' in df.columns and f'{COLUMN_TELEPHONE}' in df.columns and f'{COLUMN_MESSAGE}' in df.columns:
                usuarios = df[f'{COLUMN_USER}'].tolist()
                telefonos = df[f'{COLUMN_TELEPHONE}'].tolist()
                mensajes = df[f'{COLUMN_MESSAGE}'].tolist()

                # Mostrar los nombres de usuario y sus teléfonos
                for usuario, telefono, mensaje in zip(usuarios, telefonos, mensajes):
                    print(f"Enviando a Usuario: {usuario}, Telefono: {telefono}")
                    send_wpp(f"{telefono}", f"{mensaje}")
            else:
                print(f"Las columnas '{COLUMN_USER}' y/o '{COLUMN_TELEPHONE}' no se encontraron en el archivo Excel.")
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

def listar_archivos(carpeta):
    lista_archivos = []
    for nombre_archivo in os.listdir(carpeta):
        ruta_archivo = os.path.join(carpeta, nombre_archivo)
        if os.path.isfile(ruta_archivo):
            lista_archivos.append(nombre_archivo)
    return lista_archivos

def mostrarMenu():
    print(f""" 
        Listas cargadas:
            {archivos_excel}
        1-Comenzar envío
        2-Refrescar listas
        3-Comprobar whatsapp instalado
        4-Salir

        """)

def goToMenuPrincipal():
    # os.system('cls')
    art()
    mostrarMenu()

def refrescarLista():
    global archivos_excel
    archivos_excel = listar_archivos(excel_directory())
    
def testfunct():
        send_wpp("5491164624620", """Hola Brenda! Cómo estás? Sabías que la Nutrición en el 
                 
                  deporte es tan importante como el entrenamiento y el descanso? 
                 
                 Por eso queremos invitarte a participar de nuestro curso de Nutrición Deportiva. Solo por este mes te ofrecemos un dto del 15con el siguiente cupon. CUPON: DESCUENTONUTRIAMAIPJULIOInscribite en el siguiente link: 👇https://amaip.com.ar/curso.php?id=40""")

def abrirWpp():
    url = "http://wa.me/"
   # Abrir la URL en un navegador web
    webbrowser.open(url, new=1)

    time.sleep(6)

    count = 0

    comprobado = False

    while count <= 5 and not comprobado:
        pyautogui.hotkey('ctrl', '1')
        pyperclip.copy("Test")
        time.sleep(.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        comprobado = testear_mensaje_copiado("Test")
        count+=1
    
    if comprobado:
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'x')
        print(f"Whatsapp comprobado correctamente {count}")
    else:
        print("Fallo al abrir whatsapp")

    cerrarVentanaWpp()

    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')



def main():

    global archivos_excel

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
            testfunct()
        
        goToMenuPrincipal()

    print("Adios.")


if __name__ == "__main__":
    main()