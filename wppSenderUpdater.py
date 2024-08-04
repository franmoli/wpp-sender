import requests
import os
import subprocess
import sys
import time

def download_file(url, local_filename):
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

def main():
    if len(sys.argv) != 3:
        print("Uso: python updater.py <download_url> <app_executable>")
        sys.exit(1)

    download_url = sys.argv[1]
    app_executable = sys.argv[2]
    temp_filename = '../update.exe'

    # Descargar el nuevo ejecutable
    print(f"Descargando actualización desde {download_url}...")
    download_file(download_url, temp_filename)

    # Esperar un momento para asegurarse de que la aplicación principal se haya cerrado
    time.sleep(2)

    # Reemplazar el ejecutable antiguo
    os.remove(app_executable)
    os.rename(temp_filename, app_executable)

    # Reiniciar la aplicación principal
    print("Reiniciando aplicación...")
    subprocess.Popen([app_executable])
    sys.exit()

if __name__ == "__main__":
    main()
