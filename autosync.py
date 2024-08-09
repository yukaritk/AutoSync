import time
from datetime import datetime
import schedule
import threading
import os
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pyautogui
import cv2
import tkinter as tk
from tkinter import ttk
import json
import numpy as np
import zipfile
import shutil
from pathlib import Path
import sys

data_file_path = 'autoSyncData.json'

loja_satelite = 'https://app.omie.com.br/gestao/perfumaria-0ecjosu8/'
loja_centro = 'https://app.omie.com.br/gestao/perfumaria-0eqi7esk/'
loja_filial4 = 'https://app.omie.com.br/gestao/perfumaria-0yzu80ef/'
loja_filial1 = 'https://app.omie.com.br/gestao/perfumaria-0eceucow/'

loja_list = [loja_satelite, loja_centro, loja_filial4, loja_filial1]

def load_data():
    if os.path.exists(data_file_path):
        with open(data_file_path, 'r') as file:
            return json.load(file)
    else:
        # Cria o arquivo com dados vazios se não existir
        data = {
            'diretorio_lapa': '',
            'diretorio_sjc': '',
            'root_folder': ''
        }
        with open(data_file_path, 'w') as file:
            json.dump(data, file)
        return data

def update_status_label(text):
    def _update():
        status_label.config(text=text)
        root.update_idletasks()
    root.after(0, _update)

def job_sequence():
    try:
        sys.exit()
    except:
        pass
    def async_job():
        func_lapa = False
        while func_lapa == False: 
            func_lapa = generate_file_lapa()
        update_status_label('File Lapa solicitado.')

        # Gerar arquivos SJC para cada loja
        for loja in loja_list:
            func_sjc = False
            while func_sjc == False:
                func_sjc = generate_file_sjc(loja)
            sincronizar()

        # Download
        func_download = False
        while func_download == False:
            func_download = download()
        sincronizar()

        # Atualizar status
        time.sleep(5)
        update_status_label('ON')

    threading.Thread(target=async_job).start()

def generate_file_lapa():
    global status_label

    try:
        # Atualizando Status
        update_status_label("Processando...")

        # Inicializar o navegador
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        navegador = webdriver.Chrome(options=options)
        wait = WebDriverWait(navegador, 10)
        navegador.get('https://tecfiscal.linx.com.br/#/app/emissao/notas/download')
        time.sleep(3)

        # Encontrar e preencher o campo de email
        email = navegador.find_element(By.ID, 'email')
        email.click()
        email.send_keys('xml.venda.sumire@gmail.com')
        time.sleep(1)

        # Encontrar e preencher o campo de senha
        senha = navegador.find_element(By.ID, 'senha')
        senha.click()
        senha.send_keys('sumire@123')
        time.sleep(1)

        # Clicar no botão de conectar
        navegador.find_element(By.ID, 'btn_conectar').click()

        # Aguardar o carregamento da página após login
        time.sleep(5)
        
        # Clicar no primeiro botão de download usando CSS selector
        download_buttons = navegador.find_elements(By.CSS_SELECTOR, "button.ant-btn-primary")
        download_buttons[1].click()
        time.sleep(1)

        # Clicar no elemento <nz-range-picker> para abrir o seletor de datas
        date_picker = navegador.find_element(By.CSS_SELECTOR, "nz-range-picker.ant-picker")
        date_picker.click()
        time.sleep(1)

        # Tirar um screenshot da tela
        screenshot_path = 'screenshot.png'
        navegador.save_screenshot(screenshot_path)

        # Processar a imagem do screenshot para identificar a localização do quadrado ao redor da data de hoje
        img = cv2.imread(screenshot_path, cv2.IMREAD_GRAYSCALE)
        _, img_bin = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY_INV)
        img_bin = cv2.dilate(img_bin, np.ones((5, 5), np.uint8))

        # Encontrar contornos
        contours, _ = cv2.findContours(img_bin, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
        today_x, today_y = None, None
        for contour in contours:
            approx = cv2.approxPolyDP(contour, 0.02 * cv2.arcLength(contour, True), True)
            if len(approx) == 4:
                x, y, w, h = cv2.boundingRect(contour)
                if 20 < w < 50 and 20 < h < 50:
                    today_x = x + w // 2
                    today_y = (y + h // 2) + 135 # Ajuste hight notification
                    break
        
        today = datetime.today().weekday()
        if today == 6:
            today_x = today_x + 150 # Sabado anterior
            today_y = today_y - 25
        else:
            today_x = today_x - 25 # Data anterior

        pyautogui.click(today_x, today_y)
        pyautogui.click(today_x, today_y)

        # Encontrar e preencher o campo de email adicional usando o atributo 'type'
        additional_email = navegador.find_element(By.XPATH, "//input[@type='email']")
        additional_email.click()
        additional_email.clear()
        additional_email.send_keys('xml.venda.sumire@gmail.com')

        next_button = navegador.find_element(By.XPATH, "//button[span[text()=' Próximo ']]")
        next_button.click()
        time.sleep(1)

        # Localizar o campo de busca
        search_input = navegador.find_element(By.CSS_SELECTOR, "input.ant-select-selection-search-input")
        search_input.click()
        
        # Aguardar as opções carregarem e selecionar a primeira opção
        time.sleep(1)
        first_option = navegador.find_element(By.CSS_SELECTOR, "div.ant-select-item-option-content")
        first_option.click()

        # Localizar e clicar no botão "Adicionar" usando XPath
        adicionar_button = navegador.find_element(By.XPATH, "//button[span[text()=' Adicionar ']]")
        adicionar_button.click()

        # Localizar e clicar no botão "Salvar" usando XPath
        salvar_button = navegador.find_element(By.XPATH, "//button[span[text()=' Salvar ']]")
        salvar_button.click()

        # Aguardar para gerar os resultados
        time.sleep(5)

        # Fechar o navegador
        navegador.quit()
        return True
    except Exception as e:
        update_status_label(f"ERRO - {e}")
        return False

def generate_file_sjc(loja):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # Inicializa o driver do Chrome
    navegador = webdriver.Chrome(options=options)

    try:
        # Abre a página desejada
        navegador.get(loja)

        time.sleep(5)
        # Faz login
        user_mail = navegador.find_element(By.ID, 'email')
        user_mail.send_keys('Nagaokakarina@gmail.com')
        navegador.find_element(By.ID, 'btn-continue').click()

        time.sleep(3)
        user_pass = navegador.find_element(By.ID, 'current-password')
        user_pass.send_keys('Knaga0608')

        time.sleep(1)
        navegador.find_element(By.ID, 'btn-login').click()

        time.sleep(10)

        # Fecha possíveis pop-ups ou notificações
        navegador.find_element(By.ID, 'onesignal-slidedown-cancel-button').click()
        time.sleep(2)

        # Clica no primeiro link que leva ao menu desejado
        navegador.find_element(By.CSS_SELECTOR, 'li.tile.VPR a').click()
        time.sleep(2)

        # Move o mouse para o elemento que exibe o menu com hover
        time.sleep(3)
        hover_element = navegador.find_element(By.CSS_SELECTOR, "li[data-id='VPR']")

        # Cria uma instância de ActionChains
        time.sleep(3)
        actions = ActionChains(navegador)
        actions.move_to_element(hover_element).perform()

        # Espera até que o menu apareça
        menu_item = navegador.find_element(By.CSS_SELECTOR, "a[data-slug='painel-nfce-ven']")
        navegador.execute_script("arguments[0].scrollIntoView(true);", menu_item)
        time.sleep(1)
        actions.move_to_element(menu_item).click().perform()

        time.sleep(5)

        # Caminha até Periodo
        actions.send_keys(Keys.TAB).perform()
        actions.send_keys(Keys.TAB).perform()
        actions.send_keys(Keys.TAB).perform()

        # Digita Ontem
        actions.send_keys(Keys.DELETE).perform()
        actions.send_keys(Keys.ARROW_DOWN).perform()
        actions.send_keys(Keys.ARROW_DOWN).perform()
        actions.send_keys(Keys.ENTER).perform()

        time.sleep(2)

        # Atualizar

        atualizar_button = navegador.find_element(By.CSS_SELECTOR, "button[data-did='50688']")
        atualizar_button.click()

        time.sleep(20)
        try:
            # Localiza e clica no botão "EXPORTAR XML" usando partial link
            button_export = navegador.find_element(By.PARTIAL_LINK_TEXT, 'Exportar XML')
            navegador.execute_script("arguments[0].click();", button_export)

            time.sleep(2)

            button_yes = navegador.find_element(By.CSS_SELECTOR, '.fal.fa-check')
            navegador.execute_script("arguments[0].click();", button_yes)

            actions.send_keys(Keys.ENTER).perform()

            time.sleep(10)

            button_descartar = navegador.find_element(By.CSS_SELECTOR, '.fal.fa-times')
            navegador.execute_script("arguments[0].click();", button_descartar)

            time.sleep(2)
        except:
            pass
    except Exception as e:
        navegador.quit()
        return False

    # Fecha o navegador
    navegador.quit()
    return True

def download():
    global status_label

    try:
        # Atualizando Status
        update_status_label("Processando...")

        # Ler o arquivo autoSyncData
        data = load_data()

        # Inicializar o navegador
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        navegador = webdriver.Chrome(options=options)
        navegador.get('https://tecfiscal.linx.com.br/#/app/emissao/notas/download')
        time.sleep(3)

        # Encontrar e preencher o campo de email
        email = navegador.find_element(By.ID, 'email')
        email.click()
        email.send_keys('xml.venda.sumire@gmail.com')
        time.sleep(1)

        # Encontrar e preencher o campo de senha
        senha = navegador.find_element(By.ID, 'senha')
        senha.click()
        senha.send_keys('sumire@123')
        time.sleep(1)

        # Clicar no botão de conectar
        navegador.find_element(By.ID, 'btn_conectar').click()
        time.sleep(30)
        
        # Ajustar X
        x_width = 765
        # Ajusta a altura de Y
        y_hight_notification = 327

        # Move para a posição do retângulo verde
        pyautogui.moveTo(x_width, y_hight_notification)

        time.sleep(1)

        # Clicar no elemento usando CSS
        download_button = navegador.find_element(By.CSS_SELECTOR, 'svg[data-icon="download"]')
        download_button.click()

        time.sleep(10)
    except Exception as e:
        update_status_label(f"Erro: {e}")
        return False
    navegador.quit()
    return True  

def sincronizar():
    global status_label
    now = datetime.now()
    try:
        # Atualizar status para "Processando..."
        update_status_label("Processando...")

        # Diretório do arquivo de configuração
        data = load_data()

        # Entrar na pasta de origem (anteriormente "Downloads")
        origem_folder = data['root_folder']

        # Armazenar os arquivos zipados e suas datas de inclusão
        zip_files = []

        for item in os.listdir(origem_folder):
            item_path = os.path.join(origem_folder, item)
            if os.path.isfile(item_path) and item.endswith('.zip'):
                creation_time = os.path.getctime(item_path)
                zip_files.append((item_path, creation_time))

        # Verificar se há arquivos zipados
        if zip_files:
            # Encontrar o arquivo mais recente
            latest_file = max(zip_files, key=lambda x: x[1])[0]
        else:
            update_status_label("Nenhum arquivo zip encontrado.")
            return

        # Verificar se o arquivo zip existe
        if not os.path.exists(latest_file):
            update_status_label(f"Arquivo zip não encontrado: {latest_file}")
            return

        # Descompactar a pasta zipada
        with zipfile.ZipFile(latest_file, 'r') as zip_ref:
            extracted_folder = os.path.join(origem_folder, os.path.splitext(os.path.basename(latest_file))[0])
            zip_ref.extractall(extracted_folder)

        # Diretório de destino
        destino_lapa = Path(data['diretorio_lapa'])
        destino_sjc = Path(data['diretorio_sjc'])
        count = 0

        destino_lapa_pross = os.path.join(destino_lapa, 'Processado')
        destino_sjc_pross = os.path.join(destino_sjc, 'Processado')

        if os.path.isdir(extracted_folder):
            # Verificar subpastas para Lapa
            lapa_dirs = ['43582584000163', '43582584000325', '43582584000406', '43582584000597']
            for subfolder in os.listdir(extracted_folder):
                subfolder_path = os.path.join(extracted_folder, subfolder)
                if os.path.isdir(subfolder_path) and subfolder in lapa_dirs:
                    for xml_file in os.listdir(subfolder_path):
                        if xml_file.endswith(".xml"):
                            src_file = os.path.join(subfolder_path, xml_file)
                            dest_file = os.path.join(destino_lapa, xml_file)
                            # Verificar essa parte. Se o arquivo xml_file não estiver dentro da pasta destino_lapa_pross, deve haver a cópia do arquivo.
                            if not os.path.exists(os.path.join(destino_lapa_pross, xml_file)):
                                shutil.copy(src_file, dest_file)
                                count += 1

            # Verificar arquivos para São José dos Campos
            for xml_file in os.listdir(extracted_folder):
                if xml_file.startswith("SAT") and xml_file.endswith(".xml"):
                    src_file = os.path.join(extracted_folder, xml_file)
                    dest_file = os.path.join(destino_sjc, xml_file)
                    # Verificar essa parte. Se o arquivo xml_file não estiver dentro da pasta destino_sjc_pross, deve haver a cópia do arquivo.
                    if not os.path.exists(os.path.join(destino_sjc_pross, xml_file)):
                        shutil.copy(src_file, dest_file)
                        count += 1

            shutil.rmtree(extracted_folder)
            os.remove(latest_file)
        else:
            update_status_label("Erro ao processar.")
            return

        # Atualizar status para "XMLs sincronizados"
        update_status_label(f"{count} XMLs sincronizados")
    except Exception as e:
        update_status_label(f"ERRO - {e}")
    finally:
        root.update_idletasks()
    time.sleep(5)

def save_data(data):
    with open(data_file_path, 'w') as file:
        json.dump(data, file)

def page_config():
    # Carregar dados do arquivo JSON
    data = load_data()

    root_config = tk.Tk()
    root_config.title('Configuracao')
    root_config.geometry("400x450")

    frame_root = tk.Frame(root_config)
    frame_root.grid(padx=10, pady=10)

    bold_font = ('Helvetica', 12, 'bold')

    # Campos para Lapa
    frame_lapa = tk.Frame(frame_root)
    frame_lapa.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

    frame_lapa_tittle = ttk.Label(frame_lapa, text='LAPA', font=bold_font)
    frame_lapa_tittle.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="ew")

    folder_label_lapa = ttk.Label(frame_lapa, width=10, text='Diretorio')
    folder_label_lapa.grid(row=3, column=0, padx=10, pady=10, sticky="w")
    folder_entry_lapa = ttk.Entry(frame_lapa, width=40)
    folder_entry_lapa.grid(row=3, column=1, padx=10, pady=10)
    folder_entry_lapa.insert(0, data['diretorio_lapa'])

    # Campos para São José dos Campos
    frame_sjc = tk.Frame(frame_root)
    frame_sjc.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    frame_sjc_tittle = ttk.Label(frame_sjc, text='SAO JOSE DOS CAMPOS', font=bold_font)
    frame_sjc_tittle.grid(row=0, column=0, columnspan=2, pady=(20, 10), sticky="ew")

    folder_label_sjc = ttk.Label(frame_sjc, width=10, text='Diretorio')
    folder_label_sjc.grid(row=3, column=0, padx=10, pady=10, sticky="w")
    folder_entry_sjc = ttk.Entry(frame_sjc, width=40)
    folder_entry_sjc.grid(row=3, column=1, padx=10, pady=10)
    folder_entry_sjc.insert(0, data['diretorio_sjc'])

    # Campo para Origem
    frame_origem = tk.Frame(frame_root)
    frame_origem.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

    origem_tittle = ttk.Label(frame_origem, text='ORIGEM', font=bold_font)
    origem_tittle.grid(row=0, column=0, columnspan=2, pady=(20, 10), sticky="ew")

    origem_label = ttk.Label(frame_origem, width=10, text='Root')
    origem_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
    origem_entry = ttk.Entry(frame_origem, width=40)
    origem_entry.grid(row=1, column=1, padx=10, pady=10)
    origem_entry.insert(0, data['root_folder'])

    def save_and_close():
        # Atualizar dados do JSON com os valores inseridos nos campos
        data['diretorio_lapa'] = folder_entry_lapa.get()
        data['diretorio_sjc'] = folder_entry_sjc.get()
        data['root_folder'] = origem_entry.get()
        save_data(data)
        root_config.destroy()

    # Botão Salvar
    save_button = ttk.Button(frame_root, text="Salvar", command=save_and_close)
    save_button.grid(row=3, column=0, pady=20, sticky="e", padx=10)

    # Configurar centralização
    frame_root.grid_columnconfigure(0, weight=1)
    frame_lapa.grid_columnconfigure(0, weight=1)
    frame_sjc.grid_columnconfigure(0, weight=1)

    root_config.mainloop()


root = tk.Tk()
root.title("autoSync")
root.geometry("400x200")

# Adicionando o ícone à janela principal
icon_path = 'autoSync_icon.png'
icon = tk.PhotoImage(file=icon_path)
root.iconphoto(True, icon)

main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True)

# Botões
button_frame = tk.Frame(main_frame)
button_frame.pack(pady=10, anchor="n")

config_button = ttk.Button(button_frame, text="CONFIG", width=10, command=page_config)
config_button.pack(side="left", padx=5)

generate_button = ttk.Button(button_frame, text="Job", width=10, command=job_sequence)
generate_button.pack(side="left", padx=5)

sync_button = ttk.Button(button_frame, text="Sync", width=10, command=download)
sync_button.pack(side="left", padx=5)

# Status
status_frame = tk.Frame(main_frame)
status_frame.pack(pady=10, expand=True)

global status_label
status_label = ttk.Label(status_frame, text="ON", font=("Helvetica", 10))
status_label.pack(expand=True)

root.mainloop()