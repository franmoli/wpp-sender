import sys
import requests
import time
import webbrowser
import subprocess
import pygetwindow as gw
import pyautogui
import pyperclip

REPO = "franmoli/wpp-sender"
VERSION = "0.1.0-alpha"

 
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
    time.sleep(3)  
    
    # copiar mensaje y comprobar si esta correcto
    copiar_mensaje(mensaje)

    # cerrar navegador y seguir 
    windows = gw.getWindowsWithTitle('chrome')  

    if windows:
        window = windows[0]  # Tomar la primera ventana que coincida
        window.activate()    # Enfocar la ventana
        time.sleep(1)        # Esperar un poco antes de enviar la combinación de teclas

        pyautogui.hotkey('ctrl', 'w')  # Ejemplo: cerrar la pestaña activa
    else:
        print("No se encontró la ventana")

def copiar_mensaje(mensaje):
    
    # seteo el numero de reintentos para copiar el mensaje
    count = 0

    while not testear_mensaje_copiado(mensaje) or count >= 5:
        focusear_chat()
        # Copiar el mensaje al portapapeles
        pyperclip.copy(mensaje)
        # pegar mensaje
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        count = count + 1

    # envio mensaje
    pyautogui.press('enter')

def testear_mensaje_copiado(mensaje):
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    texto_pegado = pyperclip.paste()

    if(mensaje != texto_pegado):
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

def main():
    art()
    autoupdate()

    # comprobar instalacion si estan todas las carpetas

    args = parse_url(sys.argv[1:])

    if(args[0] == 'send_list'):
        send_list(args[1])
    
    input("Toque enter para cerrar el programa")



if __name__ == "__main__":
    main()