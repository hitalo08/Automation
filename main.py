import pyautogui #automatiza mouse,teclado,etc
import pyperclip #extensão do pyautogui
import time #biblioteca para manipular tempo
import pandas as pd#analise de dados
import openpyxl #abrir arquivo xl
import numpy

pyautogui.PAUSE = 2
pyautogui.press('win') #apertando tecla windows
pyautogui.write('edge') #digitando chrome
pyautogui.press('enter') #entrando
pyautogui.hotkey('ctrl','t') #utilizando conjuntos de teclas
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga') #copiando link
pyautogui.hotkey('ctrl','v') #colando
pyautogui.press('enter') #pressionando enter no teclado
time.sleep(3)
pyautogui.click(x=288, y=259, clicks=2) #< está identificando a posição do mouse e clicando duas vezes para acessar
pyautogui.click(x=357, y=260)
pyautogui.click(x=1178, y=156)
pyautogui.click(x=917, y=560)
#print(pyautogui.position()) # mostra a posição do mouser
#analise de dados.
df = pd.read_excel(r'C:\Users\Hitalo\Downloads\Vendas - Dez.xlsx') #O r no inicio serve para remover todos os caracteres especiais
#colocando numero ou nome da aba no sheets ele le a aba especifica
print(df)
faturamento = df['Valor Final'].sum()
quantidade = df['Quantidade'].sum()
print(faturamento)
print(quantidade)
#entraremail
pyautogui.hotkey('ctrl','t')
pyautogui.click(x=900, y=49)
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(3)
#escreveremail
pyautogui.click(x=85, y=173)
pyperclip.copy('hitalofernandes08+PESSOAL@gmail.com')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab') #assunto do email
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')#pularcorpo

#texto corpo email
texto =f""" 
Prezados, bom dia!

O faturamento de ontem foi de: {faturamento}
A quantidade de produtos foi de: {quantidade}

Abs
Hitalo"""

pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
pyautogui.click(x=769, y=863)
