# Entrar no sistema da empresa
# Navegar no sistema e encontrar a base de vendas
# Fazer dowload da base de vendas
# Importar a base de vendas
# Calcular os indicadores
# Enviar um email para a gestão
import openpyxl
import pandas as pd
from typing import Final
import pyperclip
#      Pyperclip -> copiar texto
import time
import pyautogui
pyautogui.PAUSE = 1
# pyautogui.click ->  Clicar
# pyautogui.press ->  Pressionar 1 tecla
# pyautogui.hotkey -> Conjunto de teclas
# pyautogui.write -> Escrever texto
# pyautogui.PAUSE = "x" -> Tempo entre comandos pyautogui


pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyperclip.copy(
    "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
time.sleep(0.75)

pyautogui.hotkey("ctrl", "v")
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(x=326, y=304, clicks=2)
pyautogui.click(x=370, y=490, clicks=1)
time.sleep(1)
pyautogui.click(x=1717, y=200, clicks=1)
pyautogui.click(x=1501, y=629, clicks=1)
time.sleep(4)

tabela = pd.read_excel(r'C:\Users\jcesa\Downloads\Vendas - Dez.xlsx')
# print(tabela)
faturamento = tabela['Valor Final'].sum()
print(faturamento)
quantidade = tabela['Quantidade'].sum()
print(quantidade)

time.sleep(1)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(x=153, y=206, clicks=1)
time.sleep(1)
pyautogui.write('julio c')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
pyperclip.copy('Relatório Diário')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')


texto = f'''Prezados, bom dia!

Segue relatório de vendas:
Faturamento: R$ {faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,} unidades.


Att.,Júlio C.
'''

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
