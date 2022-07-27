# Código rodado no jupyter


import pyautogui
import paperclip
import time
import pandas as pd
import datatime

pyautogui.PAUSE = 1 

#Passo 1 
#entrar no sistema (Ex: google drive)

pyautogui.hotkey("ctrl", "t")
pyautogui.copy("link do drive")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#tempo para site carregar
time.sleep(5)

#utilizar pyautogui.position() para extrair o local do click

#Passo 2
#navegar pelo sistema e encontrar a base de dados

pyautogui.click (x=1073, y=725, clicks=2)
time.sleep(5)

#Passo 3 
#download da base de dados
pyautogui.click (x=1074, y=842) # click no arquivo
pyautogui.click (x=3265, y= 413) # click nos 3 pontos 
pyautogui.click (x=2773, y=1527) #iniciar download
time.sleep(7)

#Passo 4 
#calcular indicadores   (faturamento e quantidade de produtos)

#import pandas 

tabela = pd.read_excel(r"C:\Users\mthca\Downloads\Vendas - Dez.xlsx")
display(tabela)

quantidade = tabela["Quantidade"].sum()

faturamento = tabela["Valor Final"].sum()

printf(quantidade)
printf(faturamento)

#Passo 5
#entrar no sistema (Ex: entrar no email)

pyautogui.hotkey("ctrl", "t")
pyautogui.copy("link do email")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(7)

#Passo 6
#enviar email com dados obtidos

#clicar no botão +
pyautogui.click (x=84, y=423)
time.sleep(2)

#escrever o destinatário
pyautogui.write("email@email.com")
pyautogui.press("tab")
pyautogui.press("tab")

#escrever o assunto
pyautogui.write("Relatório de Vendas")
pyautogui.press("tab")

#escrever corpo da mensagem

texto = f"""
Prezados, bom dia

O Faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de : R$ {quantidade:,}

Att, Matheus.

"""


pyautogui.copy(texto) #feito com copy para evitar perder caracteres especiais
pyautogui.hotkey("ctrl", "v")


#enviar o email

pyautogui.hotkey("ctrl", "enter")

