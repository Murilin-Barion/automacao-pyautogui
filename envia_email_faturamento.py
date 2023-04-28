# Automação de Sistemas de Processos com Python
# Passo 1: Entrar no sistema da empresa

from IPython import display
import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 0.2

# Abre o chrome instalado no computador
pyautogui.hotkey("win")
pyautogui.write("chrome")
pyautogui.press("enter")

# Tempo para carregar o chrome
time.sleep(1)

# Atalho para abrir uma nova guia
pyautogui.hotkey("ctrl", "t")
# Copia o link do drive para a área de transferência
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
# Cola o link do drive e logo em seguida preciona Enter
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# Tempo para carregar a página
time.sleep(2.5)

# Abre a pasta do drive
pyautogui.click(x=343, y=251, clicks=2)

# Tempo para carregar a página
time.sleep(2.5)

# Importa o relatório (Faz o download)
pyautogui.click(x=356, y=260, clicks=1)
time.sleep(.3)
pyautogui.click(x=1159, y=162, clicks=1)
time.sleep(.3)
pyautogui.click(x=887, y=560, clicks=1)

time.sleep(1)

# Lê os dados do arquivo baixado
tabela = pd.read_excel(r"C:\Users\Murilo\Downloads\Vendas - Dez.xlsx")

# Calcula os indicadores (Faturamento e quantidade de produtos)
faturamento = tabela["Valor Final"].sum()
quantidade_produtos = tabela["Quantidade"].sum()

time.sleep(2)

# Abre uma nova guia no chrome e entra no gmail
pyautogui.hotkey("ctrl", "t")
pyautogui.write("gmail.com")
pyautogui.press("enter")

# Tempo para carregar a página
time.sleep(4)

# Clica em "Escrever email"
pyautogui.click(x=112, y=174)

time.sleep(2.5)
# Digita o destinatario
pyautogui.write("murilin.hotmart@gmail.com")
pyautogui.press("tab", presses=2)
# Digita o assunto
pyautogui.write("Envio de Faturamento e Total de produtos")
pyautogui.press("tab")
# Copia e cola o corpo do email
corpo_email = f"""Faturamento: {faturamento}
Total de Produtos: {quantidade_produtos}"""
pyautogui.write(corpo_email)
# Atalho para emviar email
pyautogui.hotkey("ctrl", "enter")


# time.sleep(7)
# print(pyautogui.position())
