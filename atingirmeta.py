import pyautogui
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import keyboard


wait_img1 = r'imgs\wait1.jpg'
wait_img2 = r'imgs\wait2.jpg'

def wait(mensagem_imagem):
    try:
        # Verifica se a imagem está na tela
        localizacao = pyautogui.locateOnScreen(mensagem_imagem, confidence=0.8)
        
        if localizacao:
            print("Meta atingida")
            return True
    except:
        print("Aguarde.")
        return False
# Função para verificar se a tecla Esc foi pressionada
def check_esc():
    if keyboard.is_pressed('esc'):
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Informação", "Execução interrompida")
        root.destroy()
        exit()

df = pd.read_excel(r'planilha.xlsx', sheet_name='name', engine='openpyxl')
num_linhas = len(df)

time.sleep(5)

for i in range(0, num_linhas, 2):
    check_esc()  # Verifica se a tecla Esc foi pressionada

    if pd.isna(df.iloc[i, 5]):
        continue

    if pd.isna(df.iloc[i, 52]):
        continue

    defcel = '$AJ$' + str(i + 2)
    pval = str(df.iloc[i, 52]).replace('.', ',')
    altercel = '$AG$' + str(i + 2)

    pyautogui.keyDown('alt')
    pyautogui.press('s')
    pyautogui.press('k')
    pyautogui.press('a')
    pyautogui.keyUp('alt')

    while not wait(wait_img2):
        time.sleep(0.1)

    pyautogui.typewrite(defcel)
    pyautogui.press('tab')
    pyautogui.typewrite(pval)
    pyautogui.press('tab')
    pyautogui.typewrite(altercel)
    pyautogui.press('enter')

    while not wait(wait_img1):
        time.sleep(0.1)
    
    pyautogui.press('enter')

root = tk.Tk()
root.withdraw()
messagebox.showinfo("Informação", "Execução finalizada")
root.destroy()


