import sys
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import PySimpleGUI as sg
import re
import os
import pandas as pd
import time
import urllib
import threading
import shutil
from io import BytesIO 
from pathlib import Path
from PIL import Image, UnidentifiedImageError
import chromedriver_autoinstaller



import cv2
import tempfile

chromedriver_autoinstaller.install() 

INSTRUCOES = """
Com esta ferramenta você pode fazer o envio em massa de mensagens via 
whatsapp a partir de uma planilha, podendo personalizar as mensagens
de acordo com as informações de cada linha.

Parametrização da planilha

- O arquivo precisa ser em Excel
- Deve existir uma coluna chamada CELULAR com os números de whatsapp
- Deve conter apenas uma aba
- Não pode conter filtros
- Deixe na planilha apenas as linhas que serão enviadas

Envio das mensagens

Primeiro, escreva a mensagem que será enviada ou importe de um 
arquivo de texto.

Aqui você pode personalizar as mensagens da seguinte forma:

Para inserir um valor dinâmico dentro da mensagem, escreva a tag 
com o mesmo nome da coluna na sua planilha
dentro da marcação {{% nome da coluna %}}. Exemplo:

Olá {{%NOME%}}! Sua nota na prova foi {{%NOTA%}}

Desta forma a mensagem irá ser montada com o valor da respectiva coluna.

Após configurar a mensagem, escolha a planilha em excel com os dados
e clique em Iniciar

O navegador será aberto no Whatsapp Web. Siga as instruções do site para 
conectar um dispositivo, e após a conexão o bot começará a ser executado

"""

try:
    CAMINHO_DADOS = sys._MEIPASS
except:
    CAMINHO_DADOS = ".\\"

sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.
col = []
width, height = size = 256,256
layout = [  
            [sg.Button('Instruções'), sg.Button("Redefinir Whatsapp", key='-RESETWP-')],
            [sg.Text('Insira o texto:'), sg.Multiline(size=(60, 10), key='-message-'), sg.Input(enable_events=True, visible=False, key='-txtin-'), sg.FileBrowse('Importar texto', key='-txt-')],
            [sg.Text('Selecione a planilha:'), sg.FileBrowse('Procurar', key='-file_enviar-', file_types=(('Planilha Excel', "*.xlsx"),))],
            [sg.Text('Anexar imagem ou vídeo na mensagem:'), sg.Button('Procurar', key='-CHOOSEIMG-')],
            [sg.Input(visible=False, key='-IMGPATH-')],
            [sg.Image(size=size, background_color='grey', key = "Image")],
            [sg.Slider(range=(5, 60), default_value=10, expand_x=True, orientation='horizontal', key='-seconds-'), 
             sg.Text('Tempo de espera entre mensagem em segundos \n(aumente caso a mensagem não esteja sendo enviada)')],
            # [sg.Text('Ordenar por'), sg.Combo(col, key='-COLUMNS-', expand_x=True)],
            [sg.Text(key='-SPACE-')],
            [sg.ProgressBar(100, visible=True, key='-PROGRESS-', orientation='h', expand_x=True, size=(10, 10))],
            [sg.Text('', key='-TXTPROGRESS-')],
            [sg.Button('Iniciar', key='ok'), sg.Button('Cancel')]
         ]

def send_massa(planilha, mensagem: str, ordenar=None, img=''):
    dir_path = os.getcwd()
    chrome_options2 = Options()
    chrome_options2.add_argument(r"user-data-dir=" + os.getenv('APPDATA') + "\\botmassawhatsappdata\\_sessao\\sessao")
    service = Service(executable_path=os.path.join(CAMINHO_DADOS, 'data\\chromedriver.exe'))
    navegador = webdriver.Chrome(chrome_options=chrome_options2)
    navegador.get('https://web.whatsapp.com/')
    time.sleep(4)

    contatos_df = pd.read_excel(planilha)
    window['-PROGRESS-'].update(max=len(contatos_df.index), current_count=0)
    window['-TXTPROGRESS-'].update(f'0 de {len(contatos_df.index)}')
    
    print(contatos_df.head())
    while len(navegador.find_elements(By.ID,"side")) < 1:
        time.sleep(1)

    for i,numero in enumerate(contatos_df['CELULAR']):
        try:
            if 'Unable to evaluate script' in navegador.get_log('driver')[-1]['message']:
                window['-TXTPROGRESS-'].update(f'0 de {len(contatos_df.index)} - Execução interrompida')
                break
        except:
            pass
        print(i, numero)
        numero = str(numero)
        if not numero.startswith('55'):
            numero = '55'+numero
        numero = numero.replace(' ', '').replace('-', '').replace('(', '').replace(')', '')
        window['-PROGRESS-'].update(current_count=i+1)
        columns = re.findall("{{%[a-zA-Z0-9áàâãéèêíïóôõöúçñÁÀÂÃÉÈÍÏÓÔÕÖÚÇÑ_ ]+%}}", values['-message-'])
        window['-TXTPROGRESS-'].update(f'{i+1} de {len(contatos_df.index)}. Enviando agora:\n {contatos_df.loc[[i]]}')
        texto = str(mensagem)
        for c in columns:
            try:
                texto = texto.replace(c, str(contatos_df.loc[i, c[3:-3].strip()]))
            except KeyError as keyerror:
                # sg.popup("Coluna não encontrada:", keyerror)
                window['-SPACE-'].update("Coluna não encontrada:"+ keyerror.__str__(), text_color='red')
                return None
            print(str(contatos_df.loc[i, c[3:-3].strip()]))
            # print([x[3:-3].strip() for x in columns])
        # texto = texto.replace('\n', '%0A')
        link = f"https://web.whatsapp.com/send?phone={str(int(float(numero)))}&text={urllib.parse.quote_plus(texto)}"
        print(link)
        try:
            navegador.get(link)
            while len(navegador.find_elements(By.ID, "side")) < 1:
                time.sleep(int(values['-seconds-']))
                if img!='':
                    navegador.find_element(By.CSS_SELECTOR, "span[data-icon='plus']").click()
                    time.sleep(2)
                    anexa = navegador.find_elements(By.CSS_SELECTOR, "input[type='file']")
                    
                    anexa[1].send_keys(img)
                    time.sleep(2)
                    navegador.find_element(By.CSS_SELECTOR, "span[data-icon='send']").click()
                else:
                    navegador.find_element(By.CSS_SELECTOR, "span[data-icon='send']").click()
                    time.sleep(1)
                    # navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
            time.sleep(4)
        except Exception as erro:
            with open("Logerro.txt", "a") as arquivo:
                arquivo.write(f'\nNumero {numero} invalido')
            print(erro)


def getFirstFrame(videofile):
    vidcap = cv2.VideoCapture(videofile)
    success, image = vidcap.read()
    if success:
        # Criar um arquivo temporário
        temp_file = tempfile.NamedTemporaryFile(suffix='.jpg', delete=False)
        # Salvar o frame no arquivo temporário
        cv2.imwrite(temp_file.name, image)
        return temp_file.name  # Retornar o caminho do arquivo temporário

def image_to_data(im):
    """
    Image object to bytes object.
    : Parameters
      im - Image object
    : Return
      bytes object.
    """
    with BytesIO() as output:
        im.save(output, format="PNG")
        data = output.getvalue()
    return data

# Create the Window
window = sg.Window('Envio de mensagem em massa - Whatsapp', layout, icon=os.path.join(CAMINHO_DADOS, 'data\\botwhatsapp.ico'))
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    
    if event=='Instruções':
        layout2 = [
            [sg.Text(INSTRUCOES)],
            [sg.Button('Ok', key='ok2')]
        ]
        window2 = sg.Window('Instruções', layout2, modal=True)
        while True:
            event2, values2 = window2.read()
            if event2 == sg.WIN_CLOSED or event2 == 'ok2': # if user closes window or clicks cancel
                window2.close()
                break
    if event=='-txtin-':
        txt_path = values['-txtin-']
        print(txt_path)
        with open(txt_path, 'r', encoding='utf-8') as txt:
            window['-message-'].update(txt.read()) 
        # event, values = window.read(timeout=1) 
        # print(values['-message-'])          
        # columns = re.findall("{{%[a-zA-Z0-9]+%}}", values['-message-'])
        # print(columns)
        # col = [x[3:-3].strip() for x in columns]
        # print(col)
        # window['-COLUMNS-'].update(values=[None, *col], value=values['-COLUMNS-'])
    print(event)

    if event=='ok':
        # print(values['-COLUMNS-'])
        t = threading.Thread(target=lambda : send_massa(values['-file_enviar-'], values['-message-'], img=values['-IMGPATH-']))
        t.start()
        # send_massa(values['-file_enviar-'], values['-message-'], values['-COLUMNS-'])
    
    if event=='-RESETWP-':
        ctz = sg.popup_yes_no("A sessão do whatsapp será excluída e será necessário excluir a antiga conexão do celular e realizar a conexão novamente", "Deseja realmente exluir?")
        if ctz=='Yes':
            shutil.rmtree('C:\\botmassawhatsappdata\\_sessao', ignore_errors=True)
            sg.popup_ok("Sessão do whatsapp excluida")
        else:
            pass
    
    if event=='-CHOOSEIMG-':
        img_path = sg.popup_get_file("Escolha uma imagem ou vídeo para anexar à mensagem", no_window=True, file_types=(('Imagens', "*.png;*.jpg;*.jpeg;*.mp4;*.mov"),))
        window['-IMGPATH-'].update(img_path)
        if img_path == '':
            continue
       
        if not Path(img_path).is_file():
            continue
        try:
            if img_path.endswith('png') or img_path.endswith('jpeg') or img_path.endswith('jpg'):
                im = Image.open(img_path)
            else:
                im = Image.open(getFirstFrame(img_path))
        except UnidentifiedImageError:
            continue
        w, h = size = im.size
        window['Image'].update(size=size)
        print(size)
        scale = min(width/w, height/h, 1)
        if scale != 1:
            im = im.resize((int(w*scale), int(h*scale)))
        data = image_to_data(im)
        window['Image'].update(data=data)

window.close()


# espaços do nome da coluna
# CELULAR nas instrucoes

