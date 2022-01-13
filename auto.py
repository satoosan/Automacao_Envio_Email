import pandas as pd
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

time.sleep(5)

# Ler a tabela excel
tabela = pd.read_excel(r'Vendas - Dez.xlsx')
# print(tabela)

# Calcular Faturamento
faturamento = tabela['Valor Final'].sum()
qtd_produto = tabela['Quantidade'].sum()

# print(faturamento)
# print(qtd_produto)

# Enviar para o email

# Abri diretamente a caixa de escrever
link_email = 'https://mail.google.com/mail/u/0/#inbox?compose=new'

pyautogui.hotkey('ctrl', 't')
pyperclip.copy(link_email)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

# Loading
time.sleep(5)

# Destinatario
pyautogui.write('destinatario@gmail.com')
pyautogui.press('tab')  # seleciona o email
pyautogui.press('tab')  # muda para o campo assunto

# Assunto
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')  # muda para o corpo do link_email

# Corpo
texto = f'''
Prezados, bom dia.

o faturamento de ontem foi de R$ {faturamento:,.2f}

A quantidade de produtos vendidos foi de {qtd_produto:,}

Ass.
Gui Sato
'''

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

# Enviar
pyautogui.hotkey('ctrl', 'enter')
