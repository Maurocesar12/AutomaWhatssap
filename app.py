import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento}. Favor pagar no link'

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
    try:
        seta = pyautogui.locateCenterOnScreen('seta02.PNG')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(2)
    except:
        print(f'Nao foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8')as arquivo:
            arquivo.write(f'{nome},{telefone}')







    
