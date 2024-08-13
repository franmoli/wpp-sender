import winreg as reg
import os
import sys
import ctypes
import requests
import subprocess
import shutil

REPO = "franmoli/wpp-sender"
APP = 0
UPDATER = 2


def register_protocol(protocol, command):
    key = reg.HKEY_CLASSES_ROOT
    try:
        reg.CreateKey(key, protocol)
        reg.CreateKey(key, protocol + r"\shell\open\command")
        sub_key = reg.OpenKey(key, protocol, 0, reg.KEY_WRITE)
        reg.SetValueEx(sub_key, "", 0, reg.REG_SZ, protocol)
        reg.SetValueEx(sub_key, "URL Protocol",0, reg.REG_SZ, "")
        reg.CloseKey(sub_key)
        sub_key = reg.OpenKey(key, protocol + r"\shell\open\command", 0, reg.KEY_WRITE)
        reg.SetValueEx(sub_key, "", 0, reg.REG_SZ, command)
        reg.CloseKey(sub_key)
    except WindowsError as e:
        print("Error al registrar el protocolo", str(e))
        sys.exit()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if is_admin():
        print("El script ya se está ejecutando como administrador.")
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    
def download_file(url, local_filename):
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

def get_latest_release(repo, content):
    url = f'https://api.github.com/repos/{repo}/releases/latest'
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        download_url = data['assets'][content]['browser_download_url']
        return  download_url
    else:
        raise Exception(f"Error fetching release information: {response.status_code}")

def art():
    print("""
                                _____                __         
            _      ______  ____ / ___/___  ____  ____/ /__  _____
            | | /| / / __ \/ __ \ __ \/ _ \/ __ \/ __  / _ \/ ___/
            | |/ |/ / /_/ / /_/ /__/ /  __/ / / / /_/ /  __/ /    
            |__/|__/ .___/ .___/____/\___/_/ /_/\__,_/\___/_/     
                  /_/   /_/                                       

    """)
    print("Instalando....")

def main():

    art()

    # eliminar si existe instalacion previa
    try:
        shutil.rmtree('wppSender')
        print("Carpeta 'wppSender' eliminada con éxito")
    except FileNotFoundError:
        a = 1
    except OSError as e:
        print(f"Error al eliminar la carpeta 'wppSender': {e}")
    

    #crear carpetas necesarias 
    os.makedirs("./wppSender/utils")
    os.mkdir("./wppSender/excel")
    os.mkdir("./wppSender/errores")
    os.mkdir("./wppSender/archivos")

    # registrar protocolo
    run_as_admin()
    ruta_actual = os.getcwd()
    register_protocol("wppsender", ruta_actual + "\\wppSender\\wppSender.exe %1 %2 %3 %4 %5")

    
    # descargar aplicacion
    download_url = get_latest_release(REPO, APP)
    app_executable = ruta_actual + "\\wppSender\\wppSender.exe"

    print(f"Descargando contenido 1 de 2 desde {download_url}...")
    download_file(download_url, app_executable)

    # descargar updater
    download_url = get_latest_release(REPO, UPDATER)
    app_updater = ruta_actual + "\\wppSender\\utils\\wppSenderUpdater.exe"

    print(f"Descargando contenido 2 de 2 desde {download_url}...")
    download_file(download_url, app_updater)

     # Reiniciar la aplicación principal
    print("Instalacion completa saliendo...")
    # subprocess.Popen([app_executable])
    sys.exit()

        


if __name__ == "__main__":
    main()

