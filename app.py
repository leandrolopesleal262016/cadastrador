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
import logging
import time
import re
from datetime import datetime
from chave import instituicoes_dados
from webdriver_manager.chrome import ChromeDriverManager
import pyttsx3  # Biblioteca para texto em fala
import socket

# Definições de variáveis globais
NOME_PROGRAMA = "Cadastrador de Nota Fiscal Paulista"
RELATORIO_FILE = "relatorio_processamento.txt"
LOG_FILE = "log_cadastro_numeros.txt"

# Inicializa a interface gráfica
root = tk.Tk()
root.title(f"Cadastrador de Nota Fiscal Paulista v3.0")

# Cria um ScrolledText para exibir logs na interface
log_area = scrolledtext.ScrolledText(root, width=50, height=20, state=tk.DISABLED, bg='light grey')
log_area.grid(row=4, column=0, columnspan=3, pady=5, padx=10, sticky='nsew')

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
def registrar_log(mensagem):
    def log():
        logger.info(mensagem)
    root.after(0, log)

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

# Função para atualizar o combobox de usuários com base na instituição selecionada
def atualizar_usuarios_por_instituicao(event):
    global instituicao_selecionada
    instituicao_selecionada = combo_instituicao.get()
    usuarios_da_instituicao = list(instituicoes_dados[instituicao_selecionada]['usuarios'].keys())
    
    # Atualiza o combo de usuários
    combo_usuarios['values'] = usuarios_da_instituicao
    combo_usuarios.current(0)  # Seleciona o primeiro usuário da lista
    
    # Atualiza as variáveis de usuário e senha
    selecionar_usuario(None)
    
    registrar_log(f"Instituição selecionada: {instituicao_selecionada}")

# Função para atualizar as variáveis de usuário e senha com base na seleção do ComboBox
def selecionar_usuario(event):
    global usuario, senha, nome, instituicao_selecionada
    usuario_selecionado = combo_usuarios.get()
    
    # Verifica se uma instituição foi selecionada
    if not instituicao_selecionada:
        instituicao_selecionada = combo_instituicao.get()
    
    # Obtém as informações do usuário da instituição selecionada
    usuario = instituicoes_dados[instituicao_selecionada]['usuarios'][usuario_selecionado]['user']
    senha = instituicoes_dados[instituicao_selecionada]['usuarios'][usuario_selecionado]['password']
    nome = usuario_selecionado
    
    registrar_log(f"Usuário selecionado: {usuario_selecionado} da instituição {instituicao_selecionada}")

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

# Função para verificar a presença de um elemento
def elemento_presente(driver, by, value, timeout=1):
    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        return True
    except:
        return False

# Função para fechar popup/mensagem flutuante com ESC - versão melhorada
def fechar_popup_com_esc(driver):
    try:
        # Verifica de forma mais precisa se há algum popup/mensagem visível
        popup_visivel = False
        
        # Verificar se há um popup específico com a mensagem sobre pressionar ESC
        try:
            mensagens = driver.find_elements(By.XPATH, "//div[contains(@class, 'Mensagem') and contains(text(), 'pressione ESC') or contains(text(), 'Caso o documento')]")
            if mensagens and len(mensagens) > 0 and mensagens[0].is_displayed():
                popup_visivel = True
                registrar_log("Popup específico com mensagem sobre 'pressione ESC' detectado.")
        except:
            pass
            
        # Verificar se há um elemento específico que indica o popup
        if not popup_visivel:
            try:
                popup_div = driver.find_element(By.XPATH, "//div[@class='Mensagem' and @style='display: block;']")
                if popup_div.is_displayed():
                    popup_visivel = True
                    registrar_log("Popup genérico com classe 'Mensagem' detectado.")
            except:
                pass
        
        # Se não detectou especificamente, verifica visualmente por um overlay/popup
        if not popup_visivel:
            try:
                # Verifica se há algum elemento que parece ser um overlay (cobre parte da tela)
                overlay_elements = driver.find_elements(By.XPATH, "//div[contains(@style, 'position: absolute') and contains(@style, 'z-index')]")
                for elem in overlay_elements:
                    if elem.is_displayed() and 'none' not in elem.get_attribute('style'):
                        size = elem.size
                        # Se o elemento for grande o suficiente para ser um overlay
                        if size['width'] > 200 and size['height'] > 200:
                            popup_visivel = True
                            registrar_log(f"Possível overlay/popup detectado: {elem.get_attribute('outerHTML')[:100]}...")
                            break
            except:
                pass
        
        # Se um popup foi detectado, tenta fechá-lo
        if popup_visivel:
            registrar_log("Popup/mensagem flutuante detectado! Tentando fechar...")
            
            # Tenta clicar em botão de fechar se existir
            try:
                botoes = driver.find_elements(By.XPATH, "//input[@value='Não' or @value='Fechar' or @value='OK']")
                for botao in botoes:
                    if botao.is_displayed():
                        botao.click()
                        registrar_log(f"Clicou no botão '{botao.get_attribute('value')}' para fechar a mensagem.")
                        time.sleep(1)
                        return True
            except:
                pass
                
            # Envia tecla ESC para fechar o popup
            action = ActionChains(driver)
            action.send_keys(Keys.ESCAPE).perform()
            registrar_log("Enviou tecla 'ESC' para fechar o popup.")
            time.sleep(1)
            
            # Verifica se o popup foi fechado
            # Faz uma verificação mais simples para não perder tempo
            try:
                mensagens = driver.find_elements(By.XPATH, "//div[contains(@class, 'Mensagem') and @style='display: block;']")
                if not mensagens or len(mensagens) == 0 or not mensagens[0].is_displayed():
                    registrar_log("Popup foi fechado com sucesso!")
                    return True
                else:
                    # Tenta novamente com JavaScript
                    driver.execute_script("document.dispatchEvent(new KeyboardEvent('keydown', {'key': 'Escape'}))")
                    registrar_log("Tentou fechar popup usando JavaScript.")
                    time.sleep(1)
                    return True
            except:
                # Se der erro na verificação, provavelmente o popup sumiu
                registrar_log("Popup parece ter sido fechado (não foi possível verificar).")
                return True
        else:
            # Não detectou popup, então não precisa enviar ESC
            return True
            
    except Exception as e:
        registrar_log(f"Erro ao tentar fechar popup: {str(e)}")
        # Retorna True mesmo em caso de erro para não bloquear o fluxo
        return True

# Função para configurar o navegador
def configurar_navegador():
    registrar_log("Configurando opções do navegador.")
    chrome_options = Options()
    # Ajustar o zoom para 80%
    chrome_options.add_argument("--force-device-scale-factor=0.8")
    
    # Adicionar estas opções para desativar as DevTools e mensagens de console
    chrome_options.add_argument("--disable-logging")
    chrome_options.add_argument("--log-level=3")  # Apenas mensagens fatais
    chrome_options.add_argument("--silent")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Adicionar esta opção para esconder a janela do ChromeDriver
    chrome_service = Service(ChromeDriverManager().install())
    chrome_service.creation_flags = 0x08000000  # CREATE_NO_WINDOW flag
    
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
    driver.set_window_size(1400, 900)  # Ajusta a janela do navegador para garantir que o zoom funcione corretamente
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
        registrar_log("CAPTCHA detectado. Por favor, resolva o CAPTCHA manualmente e clique no botão de login.")
        mensagem_fala = "Por favor, resolva o CAPTCHA manualmente e clique no botão de login."

        # Aguarda que o usuário resolva o CAPTCHA e faça login manualmente
        # Verifica periodicamente se saímos da tela de login
        start_time = time.time()
        timeout = 120  # 2 minutos para resolver o CAPTCHA
        
        while time.time() - start_time < timeout:
            # Verifica se ainda estamos na tela de login
            if "login.aspx" not in driver.current_url or not elemento_presente(driver, By.ID, "UserName", timeout=1):
                registrar_log("Login realizado com sucesso.")
                return
            time.sleep(2)  # Verifica a cada 2 segundos
            
        registrar_log("Tempo limite excedido para resolução do CAPTCHA.")
    except Exception as e:
        registrar_log(f"Erro ao realizar login: {str(e)}")
        verificar_tela_atual(driver)

# Função para verificar e clicar no botão "Continuar" se presente
def verificar_e_clicar_continuar(driver):
    try:
        # Verifica se o botão "Continuar" está presente e clica nele
        if elemento_presente(driver, By.ID, "btnContinuar", timeout=2):
            driver.find_element(By.ID, "btnContinuar").click()
            registrar_log("Botão 'Continuar' clicado.")
            time.sleep(2)
            return True
        return False
    except:
        registrar_log("Botão 'Continuar' não encontrado ou não pôde ser clicado.")
        return False

# Função para verificar em qual tela o programa está
def verificar_tela_atual(driver):
    try:
        # Verificar URL atual para identificação mais precisa
        url_atual = driver.current_url
        #registrar_log(f"Verificando tela atual. URL: {url_atual}")
        
        # Tela de login
        if elemento_presente(driver, By.ID, "UserName") and elemento_presente(driver, By.ID, "Password"):
            registrar_log("Detectada tela de login. Realizando login...")
            realizar_login(driver)
            return "login"
        
        # Verificação mais flexível para a tela de cadastro
        # 1. Verificar URL exata da tela de cadastro (case sensitive)
        if "EntidadesFilantropicas/CadastroNotaEntidade.aspx" in url_atual or "ListagemNotaEntidade.aspx" in url_atual:
            #registrar_log("Detectada tela de cadastro (pela URL).")
            return "cadastro"
        
        # 2. Verificar elementos específicos da tela de cadastro
        if elemento_presente(driver, By.XPATH, "//input[@value='Salvar Nota']"):
            registrar_log("Detectada tela de cadastro (pelo botão 'Salvar Nota').")
            return "cadastro"
        
        if elemento_presente(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']"):
            registrar_log("Detectada tela de cadastro (pelo campo de entrada do código).")
            return "cadastro"
            
        # Tela principal
        if "Principal.aspx" in url_atual:
            registrar_log("Detectada tela principal. Navegando para tela de cadastro...")
            navegar_para_cadastro(driver)
            return "principal"
        
        # Tela de entidades (intermediária)
        if "EntidadesFilantropicas" in url_atual:
            registrar_log("Detectada tela de entidades. Continuando navegação...")
            # Verificar em qual etapa estamos
            if elemento_presente(driver, By.ID, "ddlEntidadeFilantropica"):
                registrar_log("Detectada tela de seleção de entidade.")
                selecionar_entidade_e_continuar(driver)
                return "selecao_entidade"
            
            # Se estamos em outra parte da seção de entidades
            navegar_para_cadastro_a_partir_de_entidades(driver)
            return "entidades"
        
        # Se chegou aqui, está em uma tela desconhecida - fazer diagnóstico detalhado
        registrar_log("Tela não identificada automaticamente. Tentando navegar para a tela de cadastro...")
        
        # Após diagnóstico, tentar navegar para a tela de cadastro
        driver.get("https://www.nfp.fazenda.sp.gov.br/Principal.aspx")
        time.sleep(2)
        navegar_para_cadastro(driver)
        return "navegando_apos_diagnostico"

    except Exception as e:
        registrar_log(f"Erro ao verificar tela atual: {str(e)}")
        # Em caso de erro, tentar voltar para a página principal
        try:
            driver.get("https://www.nfp.fazenda.sp.gov.br/Principal.aspx")
            time.sleep(2)
            navegar_para_cadastro(driver)
        except:
            registrar_log("Não foi possível retornar à página principal após erro.")
        return "erro"

# Função para selecionar a entidade e continuar
def selecionar_entidade_e_continuar(driver):
    try:
        registrar_log("Tentando selecionar a entidade na tela de listagem...")
        
        # Verifica qual URL estamos para diferentes comportamentos
        url_atual = driver.current_url.lower()
        
        # Na tela de Listagem de Notas (ListagemNotaEntidade.aspx)
        if "listagemnotaentidade.aspx" in url_atual:
            registrar_log("Detectada tela de Listagem de Notas.")
            
            # Tenta localizar o combobox de Entidade
            if elemento_presente(driver, By.NAME, "ddlEntidadeFilantropica", timeout=3):
                entidade_dropdown = Select(driver.find_element(By.NAME, "ddlEntidadeFilantropica"))
                nome_entidade = instituicoes_dados[instituicao_selecionada]["nome_entidade"]
                
                # Tenta selecionar pelo texto visível
                try:
                    entidade_dropdown.select_by_visible_text(nome_entidade)
                    registrar_log(f"Selecionou a entidade '{nome_entidade}' pelo texto visível.")
                except:
                    # Se falhar, tenta selecionar o primeiro item
                    try:
                        entidade_dropdown.select_by_index(1)  # Primeiro item (0 pode ser "Selecione")
                        registrar_log("Não encontrou entidade específica. Selecionou a primeira entidade disponível.")
                    except:
                        registrar_log("Falha ao selecionar entidade do dropdown.")
                
                time.sleep(1)
            # Tenta localizar outro formato de combobox se o anterior falhar
            elif elemento_presente(driver, By.XPATH, "//select[contains(@id, 'ddlEntidade')]", timeout=3):
                entidade_dropdown = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'ddlEntidade')]"))
                nome_entidade = instituicoes_dados[instituicao_selecionada]["nome_entidade"]
                
                try:
                    entidade_dropdown.select_by_visible_text(nome_entidade)
                    registrar_log(f"Selecionou a entidade '{nome_entidade}' pelo texto visível (método 2).")
                except:
                    try:
                        entidade_dropdown.select_by_index(1)
                        registrar_log("Selecionou a primeira entidade disponível (método 2).")
                    except:
                        registrar_log("Falha ao selecionar entidade do dropdown (método 2).")
            else:
                registrar_log("Não foi possível encontrar o dropdown de entidades.")
            
            # Tenta clicar no botão "Nova Nota" - Experimentando diferentes formas de localizar
            if elemento_presente(driver, By.XPATH, "//input[@value='Nova Nota']", timeout=2):
                driver.find_element(By.XPATH, "//input[@value='Nova Nota']").click()
                registrar_log("Clicou em 'Nova Nota'.")
            elif elemento_presente(driver, By.XPATH, "//input[contains(@id, 'btnNovaNota')]", timeout=2):
                driver.find_element(By.XPATH, "//input[contains(@id, 'btnNovaNota')]").click()
                registrar_log("Clicou em 'Nova Nota' (método 2).")
            elif elemento_presente(driver, By.LINK_TEXT, "Nova Nota", timeout=2):
                driver.find_element(By.LINK_TEXT, "Nova Nota").click()
                registrar_log("Clicou em 'Nova Nota' (método 3).")
            elif elemento_presente(driver, By.XPATH, "//a[text()='Nova Nota']", timeout=2):
                driver.find_element(By.XPATH, "//a[text()='Nova Nota']").click()
                registrar_log("Clicou em 'Nova Nota' (método 4).")
            else:
                # Tenta localizar qualquer elemento com o texto "Nova Nota"
                elementos = driver.find_elements(By.XPATH, "//*[contains(text(),'Nova Nota')]")
                if elementos:
                    try:
                        elementos[0].click()
                        registrar_log("Clicou em elemento com texto 'Nova Nota'.")
                    except:
                        registrar_log("Encontrou elemento com texto 'Nova Nota' mas não conseguiu clicar.")
                else:
                    registrar_log("Botão 'Nova Nota' não encontrado. Tentando localizar pela posição na página...")
                    
                    # Última tentativa - buscar todos os botões/inputs da página
                    botoes = driver.find_elements(By.XPATH, "//input[@type='button' or @type='submit']")
                    for botao in botoes:
                        try:
                            valor = botao.get_attribute("value")
                            if valor and "nova" in valor.lower():
                                registrar_log(f"Encontrou botão com valor '{valor}'. Tentando clicar.")
                                botao.click()
                                time.sleep(1)
                                break
                        except:
                            continue
            
            # Espera um tempo para a navegação ocorrer
            time.sleep(3)
            
            # Verifica se já estamos na tela de cadastro
            if elemento_presente(driver, By.XPATH, "//input[@value='Salvar Nota']", timeout=3):
                registrar_log("Navegação para tela de cadastro bem-sucedida após selecionar entidade!")
                return True
                
            # Verifica se precisamos pressionar ESC para fechar algum popup
            action = ActionChains(driver)
            action.send_keys(Keys.ESCAPE).perform()
            registrar_log("Enviou tecla 'ESC' para fechar possíveis popups.")
            time.sleep(1)
            
            return True
            
        # Tela padrão de seleção de entidade
        else:
            registrar_log("Detectada tela padrão de seleção de entidade.")
            
            # Seleciona a entidade e clica em "Nova Nota"
            if elemento_presente(driver, By.ID, "ddlEntidadeFilantropica", timeout=3):
                entidade_dropdown = Select(driver.find_element(By.ID, "ddlEntidadeFilantropica"))
                nome_entidade = instituicoes_dados[instituicao_selecionada]["nome_entidade"]
                
                try:
                    entidade_dropdown.select_by_visible_text(nome_entidade)
                    registrar_log(f"Selecionou a entidade '{nome_entidade}'.")
                except:
                    try:
                        opcoes = entidade_dropdown.options
                        registrar_log(f"Opções disponíveis: {[o.text for o in opcoes]}")
                        if len(opcoes) > 1:
                            entidade_dropdown.select_by_index(1)
                            registrar_log("Selecionou a primeira entidade disponível.")
                    except Exception as e:
                        registrar_log(f"Erro ao selecionar entidade: {str(e)}")
                
                time.sleep(1)
                
                if elemento_presente(driver, By.XPATH, "//input[@value='Nova Nota']", timeout=2):
                    driver.find_element(By.XPATH, "//input[@value='Nova Nota']").click()
                    registrar_log("Clicou em 'Nova Nota'.")
                    time.sleep(2)
                    
                    # Verifica se há algum pop-up e pressiona ESC para fechá-lo
                    action = ActionChains(driver)
                    action.send_keys(Keys.ESCAPE).perform()
                    registrar_log("Enviou tecla 'ESC'.")
                    time.sleep(1)
                    return True
                else:
                    registrar_log("Botão 'Nova Nota' não encontrado na tela padrão.")
                    return False
            else:
                registrar_log("Dropdown de entidades não encontrado na tela padrão.")
                return False
                
    except Exception as e:
        registrar_log(f"Erro ao selecionar entidade: {str(e)}")
        return False

# Função para navegar para o cadastro a partir da seção de entidades
def navegar_para_cadastro_a_partir_de_entidades(driver):
    try:
        registrar_log("Tentando navegar a partir da seção de entidades...")
        # Primeiro tenta clicar em "Cadastramento de Cupons" se estiver na página de entidades
        if elemento_presente(driver, By.XPATH, "//a[text()='Cadastramento de Cupons']"):
            driver.find_element(By.XPATH, "//a[text()='Cadastramento de Cupons']").click()
            registrar_log("Clicou em 'Cadastramento de Cupons'.")
            time.sleep(2)
            
            # Verifica se precisa clicar em "Prosseguir"
            if elemento_presente(driver, By.XPATH, "//input[@value='Prosseguir']"):
                driver.find_element(By.XPATH, "//input[@value='Prosseguir']").click()
                registrar_log("Clicou em 'Prosseguir'.")
                time.sleep(2)
                
                # Após clicar em "Prosseguir", deve-se selecionar a entidade
                if elemento_presente(driver, By.ID, "ddlEntidadeFilantropica"):
                    selecionar_entidade_e_continuar(driver)
                    return True
            return True
        else:
            # Se não estiver na página correta, volta para a tela principal
            registrar_log("Não encontrou o link 'Cadastramento de Cupons'. Voltando para o início.")
            driver.get("https://www.nfp.fazenda.sp.gov.br/Principal.aspx")
            time.sleep(2)
            navegar_para_cadastro(driver)
            return False
    except Exception as e:
        registrar_log(f"Erro ao navegar a partir de entidades: {str(e)}")
        return False

# Função para navegar até a página de cadastro de cupons
def navegar_para_cadastro(driver):
    try:
        registrar_log("Iniciando navegação para a tela de cadastro...")
        
        # Primeiro verifica se já estamos na tela de cadastro
        url_atual = driver.current_url
        if "EntidadesFilantropicas/CadastroNotaEntidade.aspx" in url_atual or "ListagemNotaEntidade.aspx" in url_atual:
            if elemento_presente(driver, By.XPATH, "//input[@value='Salvar Nota']"):
                registrar_log("Já estamos na tela de cadastro.")
                # Fechar qualquer popup antes de prosseguir
                fechar_popup_com_esc(driver)
                return True
            
        # Verifica se estamos na tela principal
        if "Principal.aspx" in url_atual:
            registrar_log("Estamos na tela principal, navegando para Entidades...")
            
            # Tenta clicar no link "Entidades" usando diferentes métodos
            try:
                # Método 1: Usar XPath específico para o link na seção "nfpfacil"
                driver.find_element(By.XPATH, "//div[@class='nfpfacil']//a[text()='Entidades']").click()
                registrar_log("Clicou em 'Entidades' (método 1).")
            except:
                try:
                    # Método 2: Tentar com um seletor mais genérico
                    driver.find_element(By.XPATH, "//a[text()='Entidades']").click()
                    registrar_log("Clicou em 'Entidades' (método 2).")
                except:
                    try:
                        # Método 3: Tentar com um seletor de link parcial
                        driver.find_element(By.PARTIAL_LINK_TEXT, "Entidades").click()
                        registrar_log("Clicou em 'Entidades' (método 3).")
                    except:
                        # Método 4: Navegar diretamente para a URL de Entidades
                        driver.get("https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/CadastroEntidades.aspx")
                        registrar_log("Navegação direta para a página de Entidades.")
            
            # Espera para a página carregar
            time.sleep(3)
            
            # Agora tenta clicar em "Cadastramento de Cupons"
            try:
                # Verifica se o elemento está visível antes de clicar
                if elemento_presente(driver, By.XPATH, "//a[text()='Cadastramento de Cupons']", timeout=3):
                    driver.find_element(By.XPATH, "//a[text()='Cadastramento de Cupons']").click()
                    registrar_log("Clicou em 'Cadastramento de Cupons'.")
                    time.sleep(2)
                else:
                    registrar_log("Link 'Cadastramento de Cupons' não encontrado após clicar em 'Entidades'.")
                    
                    # Verifica se estamos na página de CadastroEntidades.aspx
                    if "CadastroEntidades.aspx" in driver.current_url:
                        # Tenta encontrar links relacionados a cupons
                        links = driver.find_elements(By.TAG_NAME, "a")
                        for link in links:
                            try:
                                texto = link.text.lower()
                                if "cupom" in texto or "cupons" in texto or "cadastramento" in texto:
                                    registrar_log(f"Encontrou link relacionado: '{link.text}'. Tentando clicar.")
                                    link.click()
                                    time.sleep(2)
                                    break
                            except:
                                continue
                    
                    # Última tentativa - navegar diretamente para a URL de Cadastramento de Cupons
                    if "CadastroNotaEntidade.aspx" not in driver.current_url and "ListagemNotaEntidade.aspx" not in driver.current_url:
                        driver.get("https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/CadastroNotaEntidade.aspx")
                        registrar_log("Navegação direta para a página de Cadastramento de Cupons.")
                        time.sleep(3)
            except Exception as e:
                registrar_log(f"Erro ao clicar em 'Cadastramento de Cupons': {str(e)}")
                # Tenta navegação direta
                driver.get("https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/CadastroNotaEntidade.aspx")
                registrar_log("Navegação direta para a página de Cadastramento de Cupons após erro.")
                time.sleep(3)
            
            # Verifica se precisamos clicar em "Prosseguir"
            if elemento_presente(driver, By.XPATH, "//input[@value='Prosseguir']", timeout=2):
                driver.find_element(By.XPATH, "//input[@value='Prosseguir']").click()
                registrar_log("Clicou em 'Prosseguir'.")
                time.sleep(2)
                
                # Seleciona a entidade e clica em "Nova Nota"
                if elemento_presente(driver, By.ID, "ddlEntidadeFilantropica", timeout=3):
                    selecionar_entidade_e_continuar(driver)
                    time.sleep(2)
            
            # Verifica se chegamos à tela de cadastro ou listagem
            url_atual = driver.current_url
            if "CadastroNotaEntidade.aspx" in url_atual or "ListagemNotaEntidade.aspx" in url_atual:
                registrar_log("Chegou à tela de cadastro ou listagem. Verificando se há popup para fechar...")
                fechar_popup_com_esc(driver)
                
                if elemento_presente(driver, By.XPATH, "//input[@value='Salvar Nota']", timeout=3):
                    registrar_log("Navegação para tela de cadastro bem-sucedida!")
                    return True
                else:
                    registrar_log("Não encontrou o botão 'Salvar Nota' na tela atual.")
                    return False
            else:
                registrar_log("Não foi possível identificar a tela de cadastro após navegação.")
                return False
                
        # Se não estamos na tela principal, tente navegar para ela primeiro
        else:
            registrar_log("Não estamos na tela principal. Navegando para ela primeiro...")
            driver.get("https://www.nfp.fazenda.sp.gov.br/Principal.aspx")
            time.sleep(3)
            return navegar_para_cadastro(driver)  # Chamada recursiva após ir para a tela principal
            
    except Exception as e:
        registrar_log(f"Erro ao navegar para cadastro: {str(e)}")
        # Em caso de erro, tenta navegar diretamente
        try:
            driver.get("https://www.nfp.fazenda.sp.gov.br/Principal.aspx")
            time.sleep(2)
            registrar_log("Redirecionado para a tela principal após erro. Tentando novamente...")
            # Tenta novamente após um erro, mas sem recursão para evitar loop infinito
            try:
                # Tenta clicar no link "Entidades"
                if elemento_presente(driver, By.XPATH, "//a[text()='Entidades']", timeout=3):
                    driver.find_element(By.XPATH, "//a[text()='Entidades']").click()
                    registrar_log("Clicou em 'Entidades' após erro.")
                    time.sleep(2)
                    
                    # Tenta clicar em "Cadastramento de Cupons"
                    if elemento_presente(driver, By.XPATH, "//a[text()='Cadastramento de Cupons']", timeout=3):
                        driver.find_element(By.XPATH, "//a[text()='Cadastramento de Cupons']").click()
                        registrar_log("Clicou em 'Cadastramento de Cupons' após erro.")
                        time.sleep(2)
                else:
                    # Navegação direta em caso de falha
                    driver.get("https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/CadastroNotaEntidade.aspx")
                    registrar_log("Navegação direta para a tela de cadastro após falha em encontrar links.")
            except:
                registrar_log("Falha na segunda tentativa. Tentando navegação direta final.")
                driver.get("https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/CadastroNotaEntidade.aspx")
        except:
            registrar_log("Falha em todas as tentativas de navegação.")
        return False

# Função para cadastrar um número de identificação
def cadastrar_numero(driver, numero, indice):
    try:
        # Verificar se estamos na tela de cadastro antes de começar
        tela_atual = verificar_tela_atual(driver)
        if tela_atual != "cadastro":
            registrar_log(f"[{indice}] Não está na tela de cadastro. Tentando navegar de volta.")
            return False  # Retorna falso para tentar novamente após a navegação
        
        # Só fecha popup no primeiro item ou depois de um número definido de cadastros
        first_item = (indice == 1 or isinstance(indice, str) and "Reprocessado-1" in indice)
        check_popup_interval = 100  # Verificar a cada 10 cadastros

        if first_item or (isinstance(indice, int) and indice % check_popup_interval == 0):
            registrar_log(f"[{indice}] Verificando se há popup para fechar...")
            fechar_popup_com_esc(driver)
        # else:
        #     registrar_log(f"[{indice}] Pulando verificação de popup (será verificado a cada {check_popup_interval} itens).")
        
        registrar_log(f"[{indice}] Processando número de identificação: {numero}")
        input_field = waiting(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']")
        input_field.clear()
        input_field.send_keys(Keys.HOME)
        time.sleep(0.5)
        input_field.send_keys(numero)
        time.sleep(0.5)
        colado = re.sub(r'\D', '', input_field.get_attribute('value'))
        esperado = re.sub(r'\D', '', numero)
        
        while colado != esperado:
            registrar_log(f"[{indice}] Número colado incorretamente: {colado}. Tentando novamente.")
            input_field.clear()
            input_field.send_keys(Keys.HOME)
            time.sleep(0.5)
            input_field.send_keys(numero)
            time.sleep(0.5)
            colado = re.sub(r'\D', '', input_field.get_attribute('value'))
        
        # Clicar no botão "Salvar Nota"
        if elemento_presente(driver, By.XPATH, "//input[@value='Salvar Nota']"):
            driver.find_element(By.XPATH, "//input[@value='Salvar Nota']").click()
            #registrar_log(f"[{indice}] Clicou em 'Salvar Nota'.")
            time.sleep(0.5)
        else:
            registrar_log(f"[{indice}] Botão 'Salvar Nota' não encontrado.")
            # Verificar novamente a tela atual
            tela_atual = verificar_tela_atual(driver)
            return False
        
        # Verifica a tela após tentar salvar a nota
        url_atual = driver.current_url
        if "EntidadesFilantropicas/CadastroNotaEntidade.aspx" not in url_atual and "ListagemNotaEntidade.aspx" not in url_atual or not elemento_presente(driver, By.XPATH, "//input[@value='Salvar Nota']"):
            registrar_log(f"[{indice}] Saiu da tela de cadastro após salvar. Navegando de volta.")
            tela_atual = verificar_tela_atual(driver)
            return False
        
        try:
            # Verificar mensagem de sucesso
            if elemento_presente(driver, By.ID, "lblInfo"):
                sucesso_msg = driver.find_element(By.ID, "lblInfo").text
                if "Doação registrada com sucesso" in sucesso_msg:
                    registrar_log(f"[{indice}] Cadastrada com sucesso.")
                    return True
            
            # Verificar mensagem de erro
            if elemento_presente(driver, By.ID, "lblErro"):
                erro_msg = driver.find_element(By.ID, "lblErro").text
                if " já existe no sistema" in erro_msg:
                    # registrar_log(f"[{indice}] Nota já estava cadastrada no sistema: {numero}")
                    return 'ja_cadastrada'
                if "Data da Nota excedeu" in erro_msg:
                    registrar_log(f"[{indice}] Data da Nota {numero} excedeu o prazo de cadastro.")
                    return 'expirada'
                if "Não foi possível incluir o pedido" in erro_msg:
                    registrar_log(f"[{indice}] Erro: Não foi possível incluir o pedido.")
                    return 'limite_atingido'
                else:
                    registrar_log(f"[{indice}] Erro desconhecido ao cadastrar número: {erro_msg}")
                    return False
            
            # Se não encontrou nem mensagem de sucesso nem de erro
            registrar_log(f"[{indice}] Não foi possível determinar se o cadastro foi bem-sucedido.")
            return False
            
        except Exception as e:
            registrar_log(f"[{indice}] Erro ao verificar resultado do cadastro: {str(e)}")
            # Verificar se ainda estamos na tela de cadastro
            tela_atual = verificar_tela_atual(driver)
            return False
            
    except Exception as e:
        registrar_log(f"[{indice}] Erro ao cadastrar número {numero}: {str(e)}")
        # Verificar a tela atual em caso de erro
        tela_atual = verificar_tela_atual(driver)
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

            # Verificar a tela atual antes de reprocessar
            tela_atual = verificar_tela_atual(driver)
            if tela_atual != "cadastro":
                registrar_log(f"[Reprocessado-{i}] Não está na tela de cadastro. Tentando navegar de volta.")
                time.sleep(2)

            # Tentar reprocessar até 3 vezes
            tentativas = 0
            sucesso = False
            while not sucesso and tentativas < 3:
                resultado = cadastrar_numero(driver, numero, f"Reprocessado-{i}")
                if resultado == True:
                    sucesso_reprocessados += 1
                    sucesso = True
                else:
                    tentativas += 1
                    tela_atual = verificar_tela_atual(driver)
                    if tela_atual != "cadastro":
                        registrar_log(f"[Reprocessado-{i}] Tentativa {tentativas}: Saiu da tela de cadastro. Tentando navegar de volta.")
                        time.sleep(2)
                    if tentativas >= 3:
                        registrar_log(f"[Reprocessado-{i}] Falha após 3 tentativas.")

            # Implementa a funcionalidade de pausa
            while pause_event.is_set():
                time.sleep(0.1)

    return reprocessados, sucesso_reprocessados

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

            # Verificar a tela atual antes de cadastrar
            tela_atual = verificar_tela_atual(driver)
            if tela_atual != "cadastro":
                registrar_log(f"[{i}] Não está na tela de cadastro. Tentando navegar de volta.")
                navegar_para_cadastro(driver)
                time.sleep(2)

            # Tentar cadastrar o número até 3 vezes se falhar por problemas de navegação
            tentativas = 0
            sucesso = False
            while not sucesso and tentativas < 3:
                resultado = cadastrar_numero(driver, numero, i)
                
                if resultado == True:
                    numeros_cadastrados += 1
                    sucesso = True
                elif resultado == 'ja_cadastrada':
                    numeros_ja_cadastrados += 1
                    sucesso = True
                elif resultado == 'expirada':
                    numeros_expirados += 1
                    sucesso = True
                elif resultado == 'limite_atingido':
                    registrar_log("Sistema atingiu o limite de cadastro.")
                    sucesso = True
                    break  # Sai do loop principal
                else:
                    # Falhou, verifica se ainda estamos na tela de cadastro
                    tentativas += 1
                    tela_atual = verificar_tela_atual(driver)
                    if tela_atual != "cadastro":
                        registrar_log(f"[{i}] Tentativa {tentativas}: Saiu da tela de cadastro. Tentando navegar de volta.")
                        navegar_para_cadastro(driver)
                        time.sleep(2)
                    else:
                        registrar_log(f"[{i}] Tentativa {tentativas}: Falha ao cadastrar, mas ainda na tela de cadastro. Tentando novamente.")
                        
                    if tentativas >= 3:
                        numeros_invalidos += 1
                        numeros_invalidos_lista.append(numero)
                        registrar_log(f"[{i}] Falha após 3 tentativas. Marcando como inválido.")

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

# Inicializa as variáveis globais
instituicao_selecionada = list(instituicoes_dados.keys())[0]  # Define a primeira instituição como padrão
nome = list(instituicoes_dados[instituicao_selecionada]["usuarios"].keys())[0]  # Define o primeiro usuário como padrão

# Frame para organizar os comboboxes
frame_selecao = tk.Frame(root)
frame_selecao.grid(row=0, column=0, columnspan=3, padx=10, pady=5, sticky='w')

# Adiciona um ComboBox para selecionar a instituição
label_instituicao = tk.Label(frame_selecao, text="Selecione a instituição:")
label_instituicao.grid(row=0, column=0, padx=10, pady=10, sticky='w')

combo_instituicao = ttk.Combobox(frame_selecao, values=list(instituicoes_dados.keys()), state="readonly", width=25)
combo_instituicao.grid(row=0, column=1, padx=5, pady=5)
combo_instituicao.bind("<<ComboboxSelected>>", atualizar_usuarios_por_instituicao)

# Adiciona um ComboBox para selecionar o usuário
label_usuario = tk.Label(frame_selecao, text="Selecione o usuário:")
label_usuario.grid(row=1, column=0, padx=10, pady=10, sticky='w')

# Inicialmente, carrega apenas os usuários da primeira instituição
combo_usuarios = ttk.Combobox(frame_selecao, values=list(instituicoes_dados[instituicao_selecionada]["usuarios"].keys()), state="readonly", width=25)
combo_usuarios.grid(row=1, column=1, padx=5, pady=5)
combo_usuarios.bind("<<ComboboxSelected>>", selecionar_usuario)

# Define valores padrão ao iniciar a interface
combo_instituicao.current(0)  # Seleciona a primeira instituição da lista como padrão
combo_usuarios.current(0)     # Seleciona o primeiro usuário da lista como padrão
atualizar_usuarios_por_instituicao(None)  # Atualiza os usuários disponíveis
selecionar_usuario(None)      # Atualiza as variáveis de usuário e senha

# Lista para armazenar números de identificação
numeros_identificacao = []

# Controle de threads
stop_event = threading.Event()
pause_event = threading.Event()

# ProgressBar para mostrar progresso
progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=400)
progress_bar.grid(row=5, column=0, columnspan=3, pady=10, sticky='ew')

# Define o tamanho da janela
window_width = 380  # Aumentado para acomodar novos controles
window_height = 770

# Obtém as dimensões da tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calcula a posição x para a janela ser exibida no lado direito da tela
x_position = screen_width - window_width
y_position = (screen_height - window_height) // 2  # Centraliza verticalmente

# Define a geometria da janela
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

frame = tk.Frame(root, padx=10, pady=10)
frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

entry_pasta_origem = tk.Entry(frame, width=35)
entry_pasta_origem.grid(row=1, column=0, columnspan=3, pady=5, padx=5)

select_button = tk.Button(frame, text="Selecionar Pasta", command=selecionar_pasta_e_processar)
select_button.grid(row=1, column=3, pady=5, padx=0)

start_button = tk.Button(frame, text="Iniciar Cadastro", state=tk.DISABLED, command=lambda: threading.Thread(target=cadastrar_numeros).start())
start_button.grid(row=2, column=0, pady=5, padx=1)

pause_button = tk.Button(frame, text="Pausar/Retomar", command=pausar_processamento)
pause_button.grid(row=2, column=1, pady=5, padx=5)

stop_button = tk.Button(frame, text="Parar", command=parar_processamento)
stop_button.grid(row=2, column=2, pady=5, padx=5)

# Label para mostrar a quantidade de arquivos
label_arquivos = tk.Label(frame, text="Total: 0 notas")
label_arquivos.grid(row=2, column=3, pady=5, padx=5)

root.grid_rowconfigure(4, weight=1)  # Permite que o log_area expanda
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(5, weight=0)

# Mensagem de áudio após carregar a interface gráfica
def mensagem_audio_inicio():
    engine = pyttsx3.init()
    mensagem_fala = "Clique em selecionar pasta, e selecione a pasta com os arquivos XML"
    engine.say(mensagem_fala)
    engine.runAndWait()

root.after(300, lambda: threading.Thread(target=mensagem_audio_inicio).start())  # Adiciona um atraso para garantir que a interface esteja carregada

root.mainloop()