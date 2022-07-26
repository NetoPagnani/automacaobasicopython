import pyautogui as pyg
import pyperclip as pyc
import time
import pandas as pd
import win32com.client as win32

#passo 1 entrar no sistema (link do google)

pyg.press("win")
pyg.write("chrome")
pyg.press("enter")
time.sleep(5)
pyc.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyg.hotkey("ctrl","v")
pyg.press("enter")
time.sleep(5)
#passo 2 navegar no sistema e encontrar a base de dados (entrar na pasta exportar)
pyg.click(x=392, y=308, clicks=2)
time.sleep(5)
#passo 3 download da base de dados
pyg.click(x=355, y=467)
time.sleep(3)
pyg.click(x=1160, y=192)
time.sleep(3)
pyg.click(x=959, y=597)
time.sleep(7)
#passo 4 calcular indicadores
tabela = pd.read_excel(r"C:\Users\orlan\Downloads\Vendas - Dez.xlsx")
print(tabela)
quantidade = tabela["Quantidade"].sum()
faturamento = tabela["Valor Final"].sum()
print(quantidade)
print(faturamento)
#passo 5 entrar no email
pyg.hotkey("ctrl", "t")
pyc.copy("https://mail.google.com/mail/u/0/#inbox")
pyg.hotkey("ctrl", "v")
pyg.press("enter")
time.sleep(7)

#passo 6 enviar email
pyg.click(x=105, y=209)
time.sleep(3)
pyg.write("algum@gmail.com")
pyg.press("tab")
pyg.press("tab")
pyc.copy("Relatório de Vendas")
pyg.hotkey("ctrl", "v")
pyg.press("tab")
texto = f"""
Bom dia
segue os valores do faturamento e quantidade de itens de venda
Faturamento R$ {faturamento}
Quantidade de: {quantidade}
abs.
"""
pyc.copy(texto)
pyg.hotkey("ctrl", "v")

pyg.hotkey("ctrl", "enter")
