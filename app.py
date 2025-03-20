import os
import shutil
import xml.etree.ElementTree as ET
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from anticaptchaofficial.recaptchav2proxyless import recaptchaV2Proxyless
import logging
import time
import re
from datetime import datetime
from chave import chave_api, SENHA_EMAIL, users_data
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from webdriver_manager.chrome import ChromeDriverManager
import pyttsx3  # Biblioteca para texto em fala
import requests
import socket

# Configurações de e-mail
EMAIL_REMETENTE = "leandrolopesleal26@gmail.com"
EMAIL_DESTINATARIO = "leandrolopesleal26@gmail.com"

# Definições de variáveis globais
NOME_PROGRAMA = "Cadastrador de Nota Fiscal Paulista"
RELATORIO_FILE = "relatorio_processamento.txt"
LOG_FILE = "log_cadastro_numeros.txt"

# Inicializa a interface gráfica
root = tk.Tk()
root.title(f"Cadastrador de Nota Fiscal Paulista v2.0")

# Cria um ScrolledText para exibir logs na interface
log_area = scrolledtext.ScrolledText(root, width=50, height=20, state=tk.DISABLED, bg='light grey')
log_area.grid(row=2, column=0, columnspan=3, pady=5, padx=10, sticky='nsew')

# Configuração do logging para registrar logs em tempo real
logger = logging.getLogger("CadastroLogger")
logger.setLevel(logging.INFO)
logger.propagate = False  # Desativa a propagação para evitar duplicação

# Criar um FileHandler para registrar logs em um arquivo
if not any(isinstance(handler, logging.FileHandler) for handler in logger.handlers):
    file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logger.addHandler(file_handler)

# Adicionar um TextHandler para atualizar a interface gráfica
class TextHandler(logging.Handler):
    def __init__(self, widget):
        logging.Handler.__init__(self)
        self.widget = widget

    def emit(self, record):
        def atualizar_text():
            msg = self.format(record)
            self.widget.config(state=tk.NORMAL)
            self.widget.insert(tk.END, msg + '\n')
            self.widget.yview(tk.END)
            self.widget.config(state=tk.DISABLED)
        root.after(0, atualizar_text)

# Adicionar o TextHandler se ainda não estiver adicionado
if not any(isinstance(handler, TextHandler) for handler in logger.handlers):
    text_handler = TextHandler(log_area)
    text_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logger.addHandler(text_handler)

# Função para registrar logs de forma segura na thread principal
# Função para registrar logs de forma segura na thread principal
def registrar_log(mensagem):
    def log():
        logger.info(mensagem)
    root.after(0, log)


# Função para verificar conexão com a internet
def verificar_conexao_internet():
    try:
        # Conectar a um servidor DNS do Google
        socket.create_connection(("8.8.8.8", 53))
        return True
    except OSError:
        return False

# Função para esperar por um elemento
def waiting(driver, by, value, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        return element
    except Exception:
        registrar_log(f"Elemento {value} não encontrado em {timeout} segundos.")
        raise

# Função para extrair o número de identificação dos arquivos XML
def extrair_numero_identificacao(arquivo_xml):
    try:
        tree = ET.parse(arquivo_xml)
        root = tree.getroot()

        # Namespace para XML de NFe
        ns_nfe = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        # Tenta encontrar o elemento <infNFe> para o primeiro tipo de XML
        infNFe = root.find(".//nfe:infNFe", ns_nfe)
        if infNFe is not None:
            numero_identificacao = infNFe.attrib["Id"]
            numero_identificacao = re.sub(r'^\D+', '', numero_identificacao)
            return numero_identificacao

        # Tenta encontrar o elemento <infCFe> para o segundo tipo de XML
        infCFe = root.find(".//infCFe")
        if infCFe is not None:
            numero_identificacao = infCFe.attrib["Id"]
            numero_identificacao = re.sub(r'^\D+', '', numero_identificacao)
            return numero_identificacao

        # Loga um erro se nenhum dos elementos for encontrado
        registrar_log(f"Elemento <infNFe> ou <infCFe> não encontrado no arquivo {arquivo_xml}.")
        return None
    except Exception as e:
        registrar_log(f"Erro ao processar o arquivo {arquivo_xml}: {e}")
        return None

# Função para atualizar as variáveis de usuário e senha com base na seleção do ComboBox
def selecionar_usuario(event):
    global usuario, senha, nome
    usuario_selecionado = combo_usuarios.get()
    usuario = users_data[usuario_selecionado]['user']
    senha = users_data[usuario_selecionado]['password']
    nome = usuario_selecionado
    registrar_log(f"Usuário selecionado: {usuario_selecionado}")

# Função para simular o progresso
def atualizar_progress_bar(progresso):
    def atualizar():
        progress_bar['value'] = progresso
        root.update_idletasks()
    root.after(0, atualizar)

# Função para processar arquivos XML e atualizar a ProgressBar
def processar_arquivos_xml_com_progress(pasta_origem, pasta_destino):
    registrar_log(f"\nIniciando o processamento dos arquivos XML na pasta: {pasta_origem}, aguarde...\n")
    arquivos = [arquivo for arquivo in os.listdir(pasta_origem) if arquivo.endswith('.xml')]
    total_arquivos = len(arquivos)
    
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
        registrar_log(f"Pasta de destino '{pasta_destino}' criada.")

    global numeros_identificacao
    numeros_identificacao = []  # Limpa a lista antes de cada processamento
    contador_identificacoes = 0

    for i, arquivo in enumerate(arquivos, start=1):
        if stop_event.is_set():
            registrar_log("Processamento parado pelo usuário.")
            break

        caminho_arquivo = os.path.join(pasta_origem, arquivo)
        numero_identificacao = extrair_numero_identificacao(caminho_arquivo)
        if numero_identificacao:
            numeros_identificacao.append(numero_identificacao)
            contador_identificacoes += 1
            shutil.move(caminho_arquivo, os.path.join(pasta_destino, arquivo))
        else:
            registrar_log(f"Não foi possível extrair o número de identificação do arquivo '{arquivo}'.")

        # Atualiza a progressão
        progresso = (i / total_arquivos) * 100
        atualizar_progress_bar(progresso)

        # Implementa a funcionalidade de pausa
        while pause_event.is_set():
            time.sleep(0.05)

    if not stop_event.is_set():
        # Atualiza o Label com a quantidade de arquivos na main thread
        def atualizar_label():
            label_arquivos.config(text=f"Total:  {contador_identificacoes} notas")
        root.after(0, atualizar_label)

        registrar_log(f"{contador_identificacoes} identificações foram extraídas.")
        def ativar_start_button():
            start_button.config(state=tk.NORMAL)
        root.after(0, ativar_start_button)
        messagebox.showinfo("Resultado", f"{contador_identificacoes} identificações foram extraídas de {total_arquivos} arquivos.")

# Função para iniciar o processamento dos arquivos XML
def selecionar_pasta_e_processar():
    global pasta_origem_selecionada  # Armazena a pasta selecionada
    pasta_origem_selecionada = filedialog.askdirectory()
    if pasta_origem_selecionada:
        entry_pasta_origem.delete(0, tk.END)
        entry_pasta_origem.insert(0, pasta_origem_selecionada)
        pasta_destino = os.path.join(pasta_origem_selecionada, "Verificados")
        stop_event.clear()  # Limpa o estado de parada
        pause_event.clear()  # Limpa o estado de pausa 
        threading.Thread(target=processar_arquivos_xml_com_progress, args=(pasta_origem_selecionada, pasta_destino)).start()
        
    else:
        registrar_log("Nenhuma pasta selecionada.")

# Função para resolver o CAPTCHA
def resolver_captcha(driver, site_key, url):
    registrar_log("Iniciando resolução do CAPTCHA.")
    solver = recaptchaV2Proxyless()
    solver.set_verbose(1)
    solver.set_key(chave_api)
    solver.set_website_url(url)
    solver.set_website_key(site_key)

    resposta = solver.solve_and_return_solution()
    if resposta != 0:
        driver.execute_script(f"document.getElementById('g-recaptcha-response').value = '{resposta}'")
        driver.execute_script("document.getElementById('btnLogin').removeAttribute('disabled')")
        registrar_log("Desbloqueou o botão de login.")
        time.sleep(2)
        driver.find_element(By.ID, "btnLogin").click()
    else:
        registrar_log(f"Erro ao resolver CAPTCHA: {solver.err_string}")

# Função para enviar um e-mail com o log e relatório
def enviar_email():
    registrar_log("Preparando para enviar o e-mail com o log e relatório.")

    try:
        with open(RELATORIO_FILE, "r") as relatorio_file:
            corpo = relatorio_file.read()
    except Exception as e:
        registrar_log(f"Erro ao ler o arquivo de relatório: {str(e)}")
        corpo = "Não foi possível ler o relatório."

    # Configuração do e-mail
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINATARIO
    msg['Subject'] = f"Relatório de Processamento - {NOME_PROGRAMA}"

    msg.attach(MIMEText(corpo, 'plain'))

    # Enviar o e-mail
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
            servidor.starttls()
            servidor.login(EMAIL_REMETENTE, SENHA_EMAIL)
            servidor.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, msg.as_string())
        registrar_log("E-mail enviado com sucesso.")
    except Exception as e:
        registrar_log(f"Erro ao enviar e-mail: {str(e)}")

# Função para configurar o navegador
def configurar_navegador():
    registrar_log("Configurando opções do navegador.")
    chrome_options = Options()
    # Ajustar o zoom para 80%
    chrome_options.add_argument("--force-device-scale-factor=0.8")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.set_window_size(1600, 1200)  # Ajusta a janela do navegador para garantir que o zoom funcione corretamente
    return driver

# Função para realizar o login no site
def realizar_login(driver):
    try:
        registrar_log("Acessando o site da Nota Fiscal Paulista.")
        driver.get("https://www.nfp.fazenda.sp.gov.br/login.aspx")
        waiting(driver, By.ID, "UserName")

        driver.find_element(By.ID, "UserName").send_keys(usuario)
        registrar_log("Preencheu o nome de usuário.")
        time.sleep(1)
        driver.find_element(By.ID, "Password").send_keys(senha)
        registrar_log("Preencheu a senha.")
        time.sleep(1)

        waiting(driver, By.ID, 'captchaPnl')
        chave_captcha = driver.find_element(By.CLASS_NAME, 'g-recaptcha').get_attribute('data-sitekey')

        if chave_captcha:
            registrar_log(f"Chave do CAPTCHA obtida: {chave_captcha}")
            resolver_captcha(driver, chave_captcha, "https://www.nfp.fazenda.sp.gov.br/login.aspx")
        else:
            registrar_log("Chave do CAPTCHA não encontrada.")
    except Exception as e:
        registrar_log(f"Erro ao realizar login: {str(e)}")
        verificar_tela_atual(driver)

# Função para verificar e clicar no botão "Continuar" se presente
def verificar_e_clicar_continuar(driver):
    try:
        continuar_button = driver.find_element(By.ID, "btnContinuar")
        continuar_button.click()
        registrar_log("Botão 'Continuar' clicado.")
        time.sleep(2)
    except:
        registrar_log("Botão 'Continuar' não encontrado.")

# Função para navegar até a página de cadastro de cupons
def navegar_para_cadastro(driver):
    try:
        verificar_e_clicar_continuar(driver)
        waiting(driver, By.XPATH, "//a[text()='Entidades']").click()
        registrar_log("Navegando para 'Entidades'.")
        time.sleep(2)

        waiting(driver, By.XPATH, "//a[text()='Cadastramento de Cupons']").click()
        registrar_log("Selecionou 'Cadastramento de Cupons'.")
        time.sleep(2)

        waiting(driver, By.XPATH, "//input[@value='Prosseguir']").click()
        registrar_log("Clicou em 'Prosseguir'.")
        time.sleep(2)

        entidade_dropdown = Select(waiting(driver, By.ID, "ddlEntidadeFilantropica"))
        entidade_dropdown.select_by_visible_text("ASSOCIACAO WISE MADDNESS")
        registrar_log("Selecionou a entidade 'ASSOCIACAO WISE MADDNESS'.")
        time.sleep(2)

        waiting(driver, By.XPATH, "//input[@value='Nova Nota']").click()
        registrar_log("Iniciando nova nota.")
        time.sleep(3)

        action = ActionChains(driver)
        action.send_keys(Keys.ESCAPE).perform()
        registrar_log("Enviou tecla 'ESC'.")
        time.sleep(2)
    except Exception as e:
        registrar_log(f"Erro ao navegar para cadastro: {str(e)}")
        verificar_tela_atual(driver)

# Função para verificar a presença de um elemento
def elemento_presente(driver, by, value):
    try:
        driver.find_element(by, value)
        return True
    except:
        return False

# Função para verificar em qual tela o programa está
def verificar_tela_atual(driver):
    try:
        if elemento_presente(driver, By.ID, "UserName") and elemento_presente(driver, By.ID, "Password"):
            realizar_login(driver)
            return "login"

        if elemento_presente(driver, By.XPATH, "//input[@value='Salvar Nota']"):
            return "cadastro"

        if "Principal.aspx" in driver.current_url:
            navegar_para_cadastro(driver)
            return "seleção de entidade"

    except Exception as e:
        registrar_log(f"Erro ao verificar tela atual: {str(e)}")
        return None

# Função para cadastrar um número de identificação
def cadastrar_numero(driver, numero, indice):
    try:
        registrar_log(f"[{indice}] Processando número de identificação: {numero}")
        input_field = waiting(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']")
        input_field.clear()
        input_field.send_keys(Keys.HOME)
        time.sleep(1)
        input_field.send_keys(numero)
        time.sleep(1)

        colado = re.sub(r'\D', '', input_field.get_attribute('value'))
        esperado = re.sub(r'\D', '', numero)

        while colado != esperado:
            registrar_log(f"[{indice}] Número colado incorretamente: {colado}. Tentando novamente.")
            input_field.clear()
            input_field.send_keys(Keys.HOME)
            time.sleep(1)
            input_field.send_keys(numero)
            time.sleep(1)
            colado = re.sub(r'\D', '', input_field.get_attribute('value'))

        waiting(driver, By.XPATH, "//input[@value='Salvar Nota']").click()
        time.sleep(2)
        
        # Verifica a tela após tentar salvar a nota
        tela_atual = verificar_tela_atual(driver)
        if tela_atual != "cadastro":
            return False  # Ou algum outro comportamento adequado

        try:
            sucesso_msg = waiting(driver, By.ID, "lblInfo",timeout = 1).text
            if "Doação registrada com sucesso" in sucesso_msg:
                registrar_log(f"[{indice}] cadastrada com sucesso.")
                return True
            else:
                raise Exception("Mensagem de sucesso não encontrada.")
        except Exception:
            erro_msg = waiting(driver, By.ID, "lblErro").text
            if " já existe no sistema" in erro_msg:
                registrar_log(f"[{indice}] Nota já estava cadastrada no sistema: {numero}")
                return 'ja_cadastrada'
            if "Data da Nota excedeu" in erro_msg:
                registrar_log(f"[{indice}] Data da Nota {numero} excedeu o prazo de cadastro.")
                return 'expirada'
            if "Não foi possível incluir o pedido" in erro_msg:
                registrar_log(f"[{indice}] Erro: Não foi possível incluir o pedido.")
                return 'limite_atingido'
            else:
                registrar_log(f"[{indice}] Erro desconhecido ao cadastrar número: {numero}")
                return False

        time.sleep(2)
    except Exception as e:
        registrar_log(f"[{indice}] Erro ao cadastrar número {numero}: {str(e)}")
        verificar_tela_atual(driver)
        return False

# Função para reprocessar números inválidos
def reprocessar_numeros(driver, numeros_invalidos_lista):
    reprocessados = len(numeros_invalidos_lista)
    sucesso_reprocessados = 0
    if reprocessados > 0:
        registrar_log(f"\nReprocessando {reprocessados} números inválidos...\n")
        for i, numero in enumerate(numeros_invalidos_lista, start=1):
            if stop_event.is_set():
                registrar_log("Reprocessamento interrompido pelo usuário.")
                break

            resultado = cadastrar_numero(driver, numero, f"Reprocessado-{i}")
            if resultado == True:
                sucesso_reprocessados += 1

    return reprocessados, sucesso_reprocessados

# Função para enviar dados ao monitor web
def monitor(qtd):
    registrar_log('Tentando atualizar dados no Monitor.')
    try:
        url = 'https://leandrolopesleal26.pythonanywhere.com/update'

        data = {
            'robot_name': nome,
            'quantity': qtd,
            'descricao': 'Cadastro de Nota Fiscal Paulista'
        }

        response = requests.post(url, json=data)
        registrar_log(f"Resposta da request Monitor: {response.status_code}")

        if response.json().get('success'):
            registrar_log('Dados atualizados com sucesso no Monitor.')
        else:
            registrar_log('Erro ao atualizar os dados no monitor.')
    except Exception as err:
        registrar_log(f"Erro ao tentar atualizar o monitor web: {err}")
        return False

# Função principal para cadastrar números
def cadastrar_numeros():
    global numeros_cadastrados  # Define a variável global aqui
    numeros_cadastrados = 0

    if not numeros_identificacao:
        registrar_log("Nenhum número de identificação encontrado, selecione uma pasta com arquivos XML.")
        messagebox.showwarning("Atenção", "Nenhum número de identificação disponível para cadastro.")
        return

    driver = configurar_navegador()
    start_time = time.time()  # Início do tempo de execução
    try:
        realizar_login(driver)
        navegar_para_cadastro(driver)

        numeros_invalidos = 0
        numeros_ja_cadastrados = 0
        numeros_expirados = 0
        numeros_invalidos_lista = []
        total_numeros = len(numeros_identificacao)

        for i, numero in enumerate(numeros_identificacao, start=1):
            if stop_event.is_set():
                registrar_log("Cadastro interrompido pelo usuário.")
                break

            resultado = cadastrar_numero(driver, numero, i)
            if resultado == True:
                numeros_cadastrados += 1
            elif resultado == 'ja_cadastrada':
                numeros_ja_cadastrados += 1
            elif resultado == 'expirada':
                numeros_expirados += 1
            elif resultado == 'limite_atingido':
                registrar_log("Sistema atingiu o limite de cadastro.")
                break
            else:
                numeros_invalidos += 1
                numeros_invalidos_lista.append(numero)

            # Atualiza a progressão
            progresso = (i / total_numeros) * 100
            atualizar_progress_bar(progresso)

            # Implementa a funcionalidade de pausa
            while pause_event.is_set():
                time.sleep(0.1)

        # Reprocessar números inválidos
        reprocessados, sucesso_reprocessados = reprocessar_numeros(driver, numeros_invalidos_lista)

        # Calcula o tempo de execução
        tempo_execucao = time.time() - start_time
        tempo_execucao_formatado = time.strftime("%H:%M:%S", time.gmtime(tempo_execucao))

        if not stop_event.is_set():
            registrar_log(f"Total de números cadastrados com sucesso: {numeros_cadastrados}, \nTotal de números inválidos: {numeros_invalidos}, \nTotal de notas já cadastradas no sistema: {numeros_ja_cadastrados}, \nTotal de notas expiradas: {numeros_expirados}")
            registrar_log(f"Total de números reprocessados: {reprocessados}, \nTotal de reprocessados cadastrados com sucesso: {sucesso_reprocessados}")
            registrar_log(f"Tempo total de execução: {tempo_execucao_formatado}")

            with open(RELATORIO_FILE, "w") as relatorio_file:
                relatorio_file.write(f"Total de números cadastrados com sucesso: {numeros_cadastrados}\n")
                relatorio_file.write(f"Total de números inválidos: {numeros_invalidos}\n")
                relatorio_file.write(f"Total de notas já cadastradas no sistema: {numeros_ja_cadastrados}\n")
                relatorio_file.write(f"Total de notas expiradas: {numeros_expirados}\n")
                relatorio_file.write(f"Total de números reprocessados: {reprocessados}\n")
                relatorio_file.write(f"Total de reprocessados cadastrados com sucesso: {sucesso_reprocessados}\n")
                relatorio_file.write(f"Tempo total de execução: {tempo_execucao_formatado}\n")
            
            monitor(numeros_cadastrados)
            enviar_email()
            
            # Conversão de texto em fala
            def falar_conclusao():
                engine = pyttsx3.init()
                mensagem_fala = f"Processo concluído. {numeros_cadastrados} notas foram cadastradas com sucesso. {sucesso_reprocessados} reprocessadas com sucesso."
                engine.say(mensagem_fala)
                engine.runAndWait()

            threading.Thread(target=falar_conclusao).start()

            messagebox.showinfo("Concluído", f"Cadastrados com sucesso: {numeros_cadastrados}, \nTotal de {numeros_invalidos} inválidos, \nTotal de notas já cadastradas no sistema {numeros_ja_cadastrados}, \nTotal de notas expiradas: {numeros_expirados}\nTotal de reprocessados cadastrados com sucesso: {sucesso_reprocessados}\nTempo total de execução: {tempo_execucao_formatado}")
        
    except Exception as e:
        registrar_log(f"Erro na navegação: {str(e)}")
        monitor(numeros_cadastrados)
        enviar_email()
    finally:
        driver.quit()
        registrar_log("Fechando o navegador.")

# Função para parar o processamento
def parar_processamento():
    if messagebox.askyesno("Confirmação", "Tem certeza de que deseja parar o processamento?"):
        stop_event.set()
        pause_event.clear()  # Certifique-se de que não está pausado
        registrar_log("Processamento será parado assim que possível.")
        registrar_log("Processo interrompido pelo usuário.")
        root.quit()  # Encerra a interface gráfica

# Função para pausar o processamento
def pausar_processamento():
    if not pause_event.is_set():
        if messagebox.askyesno("Confirmação", "Tem certeza de que deseja pausar o processamento?"):
            pause_event.set()
            registrar_log("Processamento pausado.")
    else:
        if messagebox.askyesno("Confirmação", "Deseja retomar o processamento?"):
            pause_event.clear()
            registrar_log("Processamento retomado.")

# Inicializa a variável nome
nome = list(users_data.keys())[0]  # Define o primeiro usuário como padrão

# Adiciona um ComboBox para selecionar o usuário
label_usuario = tk.Label(root, text="Selecione o usuário:")
label_usuario.grid(row=0, column=0, padx=10, pady=10)

combo_usuarios = ttk.Combobox(root, values=list(users_data.keys()),state="readonly")
combo_usuarios.grid(row=1, column=0, padx=5, pady=5)
combo_usuarios.bind("<<ComboboxSelected>>", selecionar_usuario)

# Define um usuário padrão ao iniciar a interface (opcional)
combo_usuarios.current(0)  # Seleciona o primeiro usuário da lista como padrão
selecionar_usuario(None)  # Atualiza as variáveis de usuário e senha

# Lista para armazenar números de identificação
numeros_identificacao = []

# Controle de threads
stop_event = threading.Event()
pause_event = threading.Event()

# ProgressBar para mostrar progresso
progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=400)
progress_bar.grid(row=3, column=0, columnspan=3, pady=10)

# Define o tamanho da janela
window_width = 450
window_height = 700

# Obtém as dimensões da tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calcula a posição x para a janela ser exibida no lado direito da tela
x_position = screen_width - window_width
y_position = (screen_height - window_height) // 2  # Centraliza verticalmente

# Define a geometria da janela
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

frame = tk.Frame(root, padx=10, pady=10)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

entry_pasta_origem = tk.Entry(frame, width=50)
entry_pasta_origem.grid(row=1, column=0, columnspan=3, pady=5, padx=5)

select_button = tk.Button(frame, text="Selecionar Pasta", command=selecionar_pasta_e_processar)
select_button.grid(row=1, column=3, pady=5, padx=5)

start_button = tk.Button(frame, text="Iniciar Cadastro", state=tk.DISABLED, command=lambda: threading.Thread(target=cadastrar_numeros).start())
start_button.grid(row=2, column=0, pady=5, padx=5)

pause_button = tk.Button(frame, text="Pausar/Retomar", command=pausar_processamento)
pause_button.grid(row=2, column=1, pady=5, padx=5)

stop_button = tk.Button(frame, text="Parar", command=parar_processamento)
stop_button.grid(row=2, column=2, pady=5, padx=5)

# Label para mostrar a quantidade de arquivos
label_arquivos = tk.Label(frame, text="Total: 0 notas")
label_arquivos.grid(row=2, column=3, pady=5, padx=5)

root.grid_rowconfigure(2, weight=1)  # Permite que o log_area expanda
root.grid_columnconfigure(0, weight=1)

# Mensagem de áudio após carregar a interface gráfica
def mensagem_audio_inicio():
    engine = pyttsx3.init()
    mensagem_fala = "Clique em selecionar pasta, e selecione a pasta com os arquivos XML"
    engine.say(mensagem_fala)
    engine.runAndWait()

root.after(300, lambda: threading.Thread(target=mensagem_audio_inicio).start())  # Adiciona um atraso para garantir que a interface esteja carregada

root.mainloop()
