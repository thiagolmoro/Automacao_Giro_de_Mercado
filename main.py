# pip install pandas numpy pyautogui openpyxl pyperclip

# importar os pacotes que serão utilizados
import pyautogui as pa # importa o pacote pyautogui e dá o apelido de pa
import time # importa o pacote time
import pandas as pd # importa o pacote pandas e dá o apelido de pd
import pyperclip # importa o pacote pyperclip


pa.PAUSE = 1 # Comando para daixar o código rodar mais lento em 1 segundo

# Abrir o navegador (chrome)
pa.press("win") # Aperta a Tecla Windows
time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
pa.write("chrome") # Pesquisa pelo Navegador Chrome
time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
pa.press("enter") # Aperta a tecla Enter
time.sleep(20) # Faz o código aguardar 20 segundos antes de executar o próximo comando
# pa.hotkey("win", "up") # Aperta a combinação de tecla Ctrl + Tecla para Cima


# Abrir site do Giro de cada um e salvar imagem de cada um

tabela = pd.read_csv("lista_bovespa.csv") # atribui a variável tabela todos os dados do arquivo csv

for linha in tabela.index: #cria a variável linha já dentro do for e atribui a ela a primeira linha da tabela
    pa.hotkey("ctrl", "l") # Aperta a combinação de tecla Ctrl+l
    time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
    cod = tabela.loc[linha,"codigo"] # atribui a variável cod o valor da linha e coluna codigo da tabela em que for está
    pa.write("https://www.girodemercado.com.br/analise/"+cod+"/") # digita esse texto no local em que o cursor estiver
    time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
    pa.press("enter") # Aperta a tecla Enter
    time.sleep(15) # Faz o código aguardar 20 segundos antes de executar o próximo comando
    pa.click(x=643, y=193) # Clica com o Mouse nessa posição da Tela
    time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
    pa.scroll(-1000) # dá um Scroll na tela para baixo
    time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
    pa.rightClick(x=1119, y=639) # Clica com o Botão Direito do Mouse nessa posição da Tela
    pa.press("down") # Aperta a tecla Tab
    pa.press("down") # Aperta a tecla Tab
    pa.press("enter") # Aperta a tecla Enter
    if linha == 0:
        time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
        pa.hotkey("ctrl", "l") # Aperta a combinação de tecla Ctrl+l
        time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
        pa.write("Desktop\Giro") # digita esse texto no local em que o cursor estiver
        time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
        pa.press("enter") # Aperta a tecla Enter
        time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.write(cod) # digita codigo no local em que o cursor estiver
        time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
        pa.press("enter") # Aperta a tecla Enter
        time.sleep(3) # Faz o código aguardar 3 segundos antes de executar o próximo comando
    else:
        time.sleep(3) # Faz o código aguardar 3 segundos antes de executar o próximo comando
        pa.press("enter") # Aperta a tecla Enter
        time.sleep(3) # Faz o código aguardar 3 segundos antes de executar o próximo comando
        
# Fecha o navegador após terminar
time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
pa.hotkey("alt", "f4") # Aperta a combinação de tecla Alt + f4 para fechar a janela


# Postar no Twitter

# Abrir o Edge

pa.press("win")# Aperta a Tecla Windows
time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
pa.write("Microsoft Edge")# Pesquisa pelo Navegador Edge
time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
pa.press("enter") # Aperta a tecla Enter
time.sleep(20) # Faz o código aguardar 20 segundos antes de executar o próximo comando

# Abrir o Twitter
pa.hotkey("ctrl", "l") # Aperta a combinação de tecla Ctrl+l

pa.write("https://twitter.com/home") # digita esse texto no local em que o cursor estiver
time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
pa.press("enter") # Aperta a tecla Enter

time.sleep(25) # Faz o código aguardar 25 segundos antes de executar o próximo comando

pa.click(x=778, y=246) # Clica com o Mouse nessa posição da Tela

for linha in tabela.index: #cria a variável linha já dentro do for e atribui a ela a primeira linha da tabela
    nome_empresa = tabela.loc[linha,"nome"] # atribui a variável nome_empresa o valor da linha e coluna nome da tabela em que for está
    codigo = tabela.loc[linha,"codigo"] # atribui a variável codigo o valor da linha e coluna codigo da tabela em que for está

    if "WINFUT" in codigo: # testa se dentro do texto tem a palabra WINFUT
        # atribui o texto a variável frase
        frase = "Acesse agora a análise gráfica para hoje de "+nome_empresa+".\n\nA Ação "+codigo+" foi analisada tecnicamente por Carlos Martins CNPI-t CCAT! - https://www.girodemercado.com.br/analise/"+codigo+"/\n\n#"+codigo+" #girodemercado #b3 #bovespa #bolsadevalores #analisegrafica #cnpi #daytrade #trader "
    elif "DOLARFUT" in codigo: # testa se dentro do texto tem a palabra DOLARFUT
        frase = "Acesse agora a análise gráfica para hoje de "+nome_empresa+".\n\nA Ação "+codigo+" foi analisada tecnicamente por Carlos Martins CNPI-t CCAT! - https://www.girodemercado.com.br/analise/"+codigo+"/\n\n#"+codigo+" #girodemercado #b3 #bovespa #bolsadevalores #analisegrafica #cnpi #daytrade #trader "
    else:
        frase = "Acesse agora a análise gráfica para hoje de "+nome_empresa+".\n\nA Ação "+codigo+" foi analisada tecnicamente por Carlos Martins CNPI-t CCAT! - https://www.girodemercado.com.br/analise/"+codigo+"/\n\n#"+codigo+" #girodemercado #b3 #bovespa #bolsadevalores #analisegrafica #cnpi #daytrade #trader #analisetecnica"

    time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando

    pyperclip.copy(frase) # copia a frase para o clipboard para poder passar pelo problema de caracter especial que o pyautogui não reconhece como acentos e ç
    pa.hotkey("ctrl", "v") # cola a frase no local em que o cursor estiver

    pa.press("tab") # Aperta a tecla Tab
    pa.press("tab") # Aperta a tecla Tab
    pa.press("enter") # Aperta a tecla Enter
    time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
    if linha == 0:
        time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
        pa.hotkey("ctrl", "l") # Aperta a combinação de tecla Ctrl+l
        time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
        pa.write("Desktop\Giro") # digita esse texto no local em que o cursor estiver
        time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
        pa.press("enter") # Aperta a tecla Enter
        time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.press("tab") # Aperta a tecla Tab
        pa.write(codigo) # digita esse codigo no local em que o cursor estiver
        time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
        pa.press("enter") # Aperta a tecla Enter
    else:
        pa.write(codigo) # digita esse codigo no local em que o cursor estiver
        time.sleep(2) # Faz o código aguardar 2 segundos antes de executar o próximo comando
        pa.press("enter") # Aperta a tecla Enter

    pa.press("tab") # Aperta a tecla Tab
    pa.press("tab") # Aperta a tecla Tab
    pa.press("tab") # Aperta a tecla Tab
    pa.press("tab") # Aperta a tecla Tab
    pa.press("tab") # Aperta a tecla Tab
    pa.press("enter") # Aperta a tecla Enter
    time.sleep(3) # Faz o código aguardar 3 segundos antes de executar o próximo comando
    pa.click(x=718, y=282) # Clica com o Mouse nessa posição da Tela

# Fecha o navegador após terminar
time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
pa.hotkey("alt", "f4") # Aperta a combinação de tecla Alt + f4 para fechar a janela


# Apagar todas as imagens da pasta Giro da Área de Trabalho

pa.hotkey("win", "e") # Aperta a combinação de tecla Windows + e
time.sleep(5) # Faz o código aguardar 5 segundos antes de executar o próximo comando
pa.hotkey("ctrl", "l") # Aperta a combinação de tecla Ctrl+l
pa.write("Desktop\Giro") # digita esse texto no local em que o cursor estiver
time.sleep(3) # Faz o código aguardar 3 segundos antes de executar o próximo comando
pa.press("enter") # Aperta a tecla Enter
pa.press("tab") # Aperta a tecla Tab
pa.press("tab") # Aperta a tecla Tab
pa.press("tab") # Aperta a tecla Tab
pa.press("tab") # Aperta a tecla Tab
pa.press("tab") # Aperta a tecla Tab
pa.press("tab") # Aperta a tecla Tab
pa.press("tab") # Aperta a tecla Tab
pa.press("tab") # Aperta a tecla Tab
pa.hotkey("ctrl", "a") # Aperta a combinação de tecla Ctrl + a
pa.press("delete") # Aperta a tecla Delete
time.sleep(5) # Faz o código aguardar 2 segundos antes de executar o próximo comando

pa.hotkey("alt", "f4") # Aperta a combinação de tecla Alt + f4 para fechar a janela