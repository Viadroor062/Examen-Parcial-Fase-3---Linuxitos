# core.py
import subprocess
import pyautogui
import time
import logging
from datetime import datetime
from pathlib import Path

#TODO mover funciones a un archivo core.py y dejar solo ejecución en runner.py
path_forms ="https://forms.office.com/r/8RQ3Qxtxvv"
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1 #0.3

def run_powershell(cmd):
    try:
        result = subprocess.run(["powershell", "-Command", cmd],
                                capture_output=True, text=True, timeout=10)
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def take_screenshot(name):
    out = Path("out")
    out.mkdir(exist_ok=True)
    ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    path = out / f"{name}_{ts}.png"
    img = pyautogui.screenshot()
    img.save(path)
    return path

def fill_form(data):
    # TODO: usar coordenadas manuales para posicionar cursor
    # Ejemplo: start_coords = (450, 320)
    # Debes documentar resolución usada en README y aquí
    x = 1041
    y = 443
    # Paso 1: Abrir menú de inicio
    pyautogui.press('win')
    time.sleep(1)

    # Paso 2: Buscar Edge
    pyautogui.write('Brave', interval=0.1)
    time.sleep(3)
    pyautogui.press('enter')

    # Paso 3: Esperar a que Edge abra
    time.sleep(5)  # Ajusta según tu máquina

    # Paso 4: Escribir URL del formulario
    pyautogui.write(path_forms, interval=0.1)
    pyautogui.press('enter')

    # Paso 5: Esperar a que cargue el formulario
    time.sleep(3)
    #Tomar la primera captura
    take_screenshot("before")
    #Selecciona el primer campo
    pyautogui.moveTo(x, y)
    pyautogui.click()
    time.sleep(1)
    pyautogui.typewrite(data["fecha"])
    pyautogui.press("enter")
    pyautogui.press("tab")
    time.sleep(1)
    #Llena el primer campo
    pyautogui.typewrite(data["fecha"])
    time.sleep(2)
    #Avanza al segundo campo y lo llena
    pyautogui.press("tab")
    pyautogui.typewrite(data["nombre1"])
    time.sleep(3)
    pyautogui.press("enter")
    pyautogui.typewrite(data["nombre2"])
    time.sleep(3)
    pyautogui.press("enter")
    pyautogui.typewrite(data["nombre3"])
    time.sleep(3)
    #Toma la segunda captura
    take_screenshot("during") 
    time.sleep(1)
    #Avanza al tercer campo y lo llena
    pyautogui.press("tab") 
    pyautogui.typewrite(data["matriculas"])
    time.sleep(3) 
    #Avanza al 4to campo y escoge alguna opción
    pyautogui.press("tab") 
    pyautogui.press("space") 
    #Avanza al boton enviár y le da clic
    pyautogui.press("tab")
    pyautogui.press("enter")
    #Toma la 3ra y última captura
    take_screenshot("after")
    
def solicitar_coordenada(eje):
    while True:
        try:
            valor = int(input(f"Introduzca el valor de {eje}: ").strip())
            return valor
        except ValueError:
            print(f" Error: el valor de {eje} debe ser un número entero. Intente de nuevo.")

def main():
    logging.basicConfig(filename="run.log", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")
    logging.info("Inicio del examen")
    
    # TODO: validar que data tenga los campos requeridos
    tiempo = datetime.now()
    fecha = tiempo.strftime("%d/%m/%Y")
    m1 = 1909518
    m2 = 2225433
    m3 = 2115684
    mt = m1+m2+m3
    data = {
        "fecha" : str(fecha),
        "matriculas" : str(mt), 
        "nombre1": "Victor Adrian Rodriguez Ortiz",
        "nombre2": "Jose Rodrigo Perez Gonzalez",
        "nombre3": "Rodolfo Uriel Hernandez de Leon"
    }
    
    # TODO: permitir que el usuario defina las coordenadas manualmente
    #print("Introduzca las coordenadas deseadas.")
    #x = int(input("Introduzca el valor de x: ").strip())
    #y = int(input("Introduzca el valor de y: ").strip())
    #start_coords = (x, y)
    
    code, out, err = run_powershell("Get-Date")
    logging.info(f"PS code: {code}")
    logging.info(f"PS output: {out}")
    logging.info(f"PS error: {err}")
    
    fill_form(data)
    
    logging.info("Fin del examen")