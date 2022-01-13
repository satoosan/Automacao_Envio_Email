# Automação - Envio de Emails com Python

A aplicação funcionará apenas se tiver aberto no navegador.

E se as seguintes teclas funcionar:
- ctrl + t (Abri uma nova guia)
- ctrl + v (Cola algum arquivo ou texto)

## Sobre o Projeto
Foi feito junto com o evento do canal "<a href="https://www.youtube.com/c/HashtagPrograma%C3%A7%C3%A3o">Hashtag Programação</a>"

Algumas das bibliotecas utilizadas:
- pyautogui e pyperclip (Automação do teclado, mouse e do monitor)
- time (Tempo)
- pandas (Manipulação e Análise de dados)

*Obs: Algumas das bibliotecas, necessitam que instalam, via terminal, utilizando o 'pip install'*

Basicamente, a aplicação, calcula o faturamento e a quantidade de produtos vendidos e os envia por email... Por exemplo para ler, os dados do excel
foi utilizado a lib **pandas**.

```Python
tabela = pd.read_excel(r'Vendas - Dez.xlsx')
```

E para "ler" as teclas ou os clicks, foi utilizado o **pyautogui** e o **pyperclip**.
```Python
# Utilizando teclas do teclado

# "Lê" um conjunto de teclas
pyautogui.hotkey('ctrl', 't')

# Escreve algum texto
pyautogui.write('Algum texto')

# Pressiona uma tecla
pyautogui.press('enter')

# Copia um texto, dentro dos parênteses
pyperclip.copy('Relatório de Vendas')

```
Obs:
- Caso utilize, o destino do arquivo a ser lido, se estiver na base, utilize o próprio nome do arquivo
- Altere o Destinatário, para um email válido
- E esteja logado em algum email

##

### Link pra download: 
- <a href="https://github.com/satoosan/Automacao_Envio_Email/archive/refs/tags/v1.0.zip">Automacao_Envio_Emails.zip</a>
- <a href="https://github.com/satoosan/Automacao_Envio_Email/releases/download/v1.0/auto.exe">Automacao_Envio_Emails.exe</a>
