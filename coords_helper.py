# coords_helper.py
import pyautogui
import time

print("Mueve el cursor a la posición deseada...")
time.sleep(10)
pos = pyautogui.position()
print(f"Coordenadas actuales: {pos}")

