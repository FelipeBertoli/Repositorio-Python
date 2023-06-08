import pyautogui
import pyperclip
import time
import pandas as pd
#print(pyautogui.position())
pyautogui.PAUSE = 0.5

# Entrar no sistema
pyautogui.hotkey("win", "1")
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(2)

# Navegar para local
pyautogui.click(x=546, y=319, clicks=2)
time.sleep(2)

# Download da base de dados
pyautogui.click(x=546, y=319, button="right")
pyautogui.click(x=581, y=644)
time.sleep(3)

# Calcular os indicadores (faturamento e quantidade de produtos)
table = pd.read_excel(r"C:\Users\user\Downloads\Vendas - Dez.xlsx")
quantidade = table["Quantidade"].sum()
faturamento = table["Valor Final"].sum()
print(quantidade)
print(faturamento)

# Entrar no email
pyautogui.hotkey("ctrl", "t")
time.sleep(1)
pyautogui.hotkey("alt", "d")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Mandar o email
pyautogui.click(x=124, y=182)
time.sleep(2)

pyautogui.write("destinatario@mail.com")
pyautogui.press("tab")
pyautogui.press("tab")

pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

text = f"""
Bom dia, 

O faturamento foi de: R$ {faturamento:,.2f}
A quantidade de produtos vendidos foi de: {quantidade}

Abraços,
Felipe
"""
pyperclip.copy(text)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")