#!/bin/bash

# Definir el nombre del icono
ICON_NAME="wppSenderIcon.ico"

# Definir los scripts Python
SCRIPT_1="wppSender.py"
SCRIPT_2="wppSenderInstaller.py"
SCRIPT_3="wppSenderUpdater.py"

# Ejecutar PyInstaller para cada script
pyinstaller --onefile --icon=$ICON_NAME --name=wppSender $SCRIPT_1
pyinstaller --onefile --icon=$ICON_NAME --name=wppSenderInstaller $SCRIPT_2
pyinstaller --onefile --icon=$ICON_NAME --name=wppSenderUpdater $SCRIPT_3

echo "Los ejecutables se han generado exitosamente."
