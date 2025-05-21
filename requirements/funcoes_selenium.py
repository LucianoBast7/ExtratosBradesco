from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.ie.service import Service
import time
import pickle
import os
import sys
import logging

class SeleniumAutomator:
    def __init__(self):
        """ Inicializa o navegador com opções configuradas """

        sys.stdout = open(os.devnull, "w")
        sys.stderr = open(os.devnull, "w")

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")  # Maximiza a janela
        options.add_argument("--no-sandbox")  # Necessário para rodar em servidores sem interface gráfica
        options.add_argument("--disable-gpu")  # Desativa aceleração por GPU
        options.add_argument("--disable-software-rasterizer")  # Evita fallback de renderização de software
        options.add_argument("--disable-dev-shm-usage")  # Reduz consumo de memória compartilhada
        options.add_argument("--disable-3d-apis")  # Bloqueia APIs 3D que usam WebGL
        options.add_argument("--disable-webgl")  # Desativa WebGL completamente
        options.add_argument("--disable-webgl2")  # Desativa WebGL 2.0
        options.add_argument("--disable-accelerated-video")  # Desativa aceleração de vídeo
        options.add_argument("--disable-accelerated-mjpeg-decode")  # Desativa decodificação MJPEG acelerada
        options.add_argument("--disable-accelerated-video-decode")  # Desativa decodificação de vídeo acelerada
        options.add_argument("--disable-accelerated-2d-canvas")  # Desativa aceleração 2D de canvas
        options.add_argument("--disable-renderer-accessibility")  # Desativa acessibilidade do renderizador
        options.add_argument("--enable-logging")  # Ativa logs do ChromeDriver
        options.add_argument("--log-level=3")  # Reduz nível de logs do ChromeDriver

        # Configuração de downloads
        prefs = {
            "download.default_directory": "C:/Users/LucianoBoaventuraBas/Downloads/Teste PDF",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }
        options.add_experimental_option("prefs", prefs)

        # Configurar logging para evitar mensagens desnecessárias do Selenium
        logging.getLogger("selenium").setLevel(logging.CRITICAL)

        # Inicializa o WebDriver do Chrome
        service = webdriver.ChromeService(log_output=os.devnull)
        self.driver = webdriver.Chrome(service=service, options=options)

        # Restaurar stdout e stderr após iniciar o driver
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__


    def navegar_para(self, url):
        """ Abre uma URL """
        self.driver.get(url)

    def fechar_navegador(self):
        """ Fecha o navegador """
        self.driver.quit()

    def maximizar_janela(self):
        """ Maximiza a janela do navegador """
        self.driver.maximize_window()

    def obter_titulo(self):
        """ Retorna o título da página atual """
        return self.driver.title

    def obter_texto_por_id(self, elemento_id):
        """ Obtém o texto de um elemento pelo ID """
        return self.driver.find_element(By.ID, elemento_id).text

    def obter_texto_por_nome(self, nome):
        """ Obtém o texto de um elemento pelo Name """
        return self.driver.find_element(By.NAME, nome).text

    def obter_texto_por_xpath(self, xpath):
        """ Obtém o texto de um elemento pelo XPath """
        return self.driver.find_element(By.XPATH, xpath).text

    def digitar_por_xpath(self, xpath, texto):
        """ Digita um texto em um campo de entrada pelo XPath """
        self.driver.find_element(By.XPATH, xpath).send_keys(texto)

    def digitar_por_id(self, elemento_id, texto):
        """ Digita um texto em um campo de entrada pelo ID """
        self.driver.find_element(By.ID, elemento_id).send_keys(texto)

    def clicar_por_xpath(self, xpath):
        """ Clica em um elemento pelo XPath """
        self.driver.find_element(By.XPATH, xpath).click()

    def clicar_por_id(self, elemento_id):
        """ Clica em um elemento pelo ID """
        self.driver.find_element(By.ID, elemento_id).click()

    def clicar_por_nome(self, nome):
        """ Clica em um elemento pelo Name """
        self.driver.find_element(By.NAME, nome).click()

    def clicar_link_por_texto(self, texto):
        """ Clica em um link pelo texto visível """
        self.driver.find_element(By.LINK_TEXT, texto).click()

    def executar_javascript(self, script):
        """ Executa um script JavaScript na página """
        return self.driver.execute_script(script)

    def alterar_frame_por_texto(self, frame_name=""):
        """ Altera para um frame específico pelo nome """
        if frame_name:
            self.driver.switch_to.frame(frame_name)
        else:
            self.driver.switch_to.default_content()

    def alterar_para_default_content(self):
        self.driver.switch_to.default_content()

    def alterar_frame_por_index(self, index):
        """ Altera para um frame pelo índice """
        self.driver.switch_to.frame(index)

    def aguardar_estado_documento(self):
        """ Aguarda a página carregar completamente """
        WebDriverWait(self.driver, 10).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )

    def selecionar_item_lista_por_nome(self, nome, texto):
        """ Seleciona um item em uma lista suspensa pelo Name """
        select = Select(self.driver.find_element(By.NAME, nome))
        select.select_by_visible_text(texto)

    def validar_checkbox_por_id(self, checkbox_id):
        """ Retorna True se um checkbox estiver marcado """
        return self.driver.find_element(By.ID, checkbox_id).is_selected()

    def aceitar_alerta(self):
        """ Aceita um alerta (popup) """
        Alert(self.driver).accept()

    def rejeitar_alerta(self):
        """ Rejeita um alerta (popup) """
        Alert(self.driver).dismiss()

    def obter_texto_alerta(self):
        """ Obtém o texto de um alerta (popup) """
        return Alert(self.driver).text

    def limpar_campo_por_xpath(self, xpath):
        """ Limpa um campo de entrada pelo XPath """
        self.driver.find_element(By.XPATH, xpath).clear()

    def selecionar_item_lista_por_id(self, elemento_id, texto):
        """ Seleciona um item em uma lista suspensa pelo ID """
        select = Select(self.driver.find_element(By.ID, elemento_id))
        select.select_by_visible_text(texto)

    def enviar_tab_por_id(self, elemento_id):
        """ Envia a tecla TAB para um campo de entrada pelo ID """
        self.driver.find_element(By.ID, elemento_id).send_keys(Keys.TAB)

    def enviar_arquivo_por_id(self, elemento_id, caminho_arquivo):
        """ Envia um arquivo para um input type='file' """
        self.driver.find_element(By.ID, elemento_id).send_keys(caminho_arquivo)

    def trocar_para_nova_aba(self):
        """ Troca para a última aba aberta """
        self.driver.switch_to.window(self.driver.window_handles[-1])  # Muda para a última aba aberta

    def tirar_print(self, caminho_arquivo):
        """ Tira um print da tela e salva no caminho especificado """
        self.driver.save_screenshot(caminho_arquivo)

    def voltar_para_aba_anterior(self):
        """ Volta para a aba anterior """
        if len(self.driver.window_handles) > 1:
            self.driver.switch_to.window(self.driver.window_handles[0])  # Volta para a primeira aba

    def fechar_nova_aba(self):
        """ Fecha a aba atual e retorna para a aba anterior """
        if len(self.driver.window_handles) > 1:
            self.driver.close()  # Fecha a aba ativa
            self.driver.switch_to.window(self.driver.window_handles[0])  # Volta para a primeira aba

    def obter_texto_por_class_name(self, class_name):
        """ Obtém o texto de um elemento pelo Class Name """
        self.driver.find_element(By.CLASS_NAME, class_name).text

    def clicar_por_partial_link_text(self, partial_text):
        """ Clica em um link pelo texto parcial """
        self.driver.find_element(By.PARTIAL_LINK_TEXT, partial_text).click()

    def obter_texto_por_partial_link(self, partial_text):
        self.driver.find_element(By.PARTIAL_LINK_TEXT, partial_text).text

    def mudar_frame_por_id(self, frame):
        self.driver.switch_to.frame(frame)

    def obter_elemento_web_por_xpath(self, elemento):
        self.driver.find_element(By.XPATH, elemento).text

    def obter_elemento_web_por_css_selector(self, elemento):
        self.driver.find_element(By.CSS_SELECTOR, elemento).text

    def esperar_time_out(self, timeout, elemento):
        WebDriverWait(self.driver, timeout).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, elemento))
        )

    def salvar_cookies(self):
        with open("cookies.pkl", "wb") as file:
            pickle.dump(self.driver.get_cookies(), file)
        
    def aplicar_cookies(self):
        with open("cookies.pkl", "rb") as file:
            cookies = pickle.load(file)
        for cookie in cookies:
            self.driver.add_cookie(cookie)
    
    def refresh(self):
        self.driver.refresh()
    
    def aguardar_elemento(self, by, valor, tempo=10):
        WebDriverWait(self.driver, tempo).until(
            EC.presence_of_element_located((by, valor))
        )
    
    def clicar_por_link_text(self, texto):
        self.driver.find_element(By.LINK_TEXT, texto).click()
    
    def limpar_cache(self):
        self.driver.execute_cdp_cmd("Network.clearBrowserCache", {})

    def limpar_cookies(self):
        self.driver.execute_cdp_cmd("Network.clearBrowserCookies", {})

    def obter_elemento_por_id(self, valor):
        self.driver.find_element(By.ID, valor)

    def aguardar_frame_e_trocar(self, valor):
        WebDriverWait(self.driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, valor))
        )

    def aguardar_elemento_para_clicar_por_xpath(self, valor):
        v1 = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, valor))
        )
        v1.click()
        time.sleep(2)

    def aguardar_elemento_para_clicar_por_id(self, valor):
        v1 = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, valor))
        )
        v1.click()
        time.sleep(2)
    
    def aguardar_link_para_capturar_por_xpath(self, valor):
        v1 = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, valor))
        ).get_attribute("href")
        return v1


class SeleniumIEAutomator:
    def __init__(self, driver_path=r"W:\PY_017 - Extratos Bradesco\7. Scripts\Certificado Bradesco\IEDriverServer.exe"):
        """ Inicializa o navegador Internet Explorer com configurações específicas """

        # Redireciona logs para evitar saída desnecessária no terminal
        sys.stdout = open(os.devnull, "w")
        sys.stderr = open(os.devnull, "w")

        options = webdriver.IeOptions()
        options.ignore_zoom_level = True  # Ignora o nível de zoom (recomendado para IE)
        options.native_events = False  # Evita problemas de clique
        options.require_window_focus = True  # Melhora a estabilidade dos cliques
        options.enable_persistent_hover = False  # Evita bugs de mouseover
        options.ignore_protected_mode_settings = True  # Necessário para evitar erros no IE

        # Inicializa o WebDriver do Internet Explorer
        service = Service(driver_path)
        self.driver = webdriver.Ie(service=service, options=options)

        # Restaurar stdout e stderr após iniciar o driver
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

    def navegar_para(self, url):
        """ Abre uma URL no Internet Explorer """
        self.driver.get(url)

    def fechar_navegador(self):
        """ Fecha o navegador """
        self.driver.quit()

    def maximizar_janela(self):
        """ Maximiza a janela do navegador """
        self.driver.maximize_window()

    def obter_titulo(self):
        """ Retorna o título da página atual """
        return self.driver.title

    def obter_texto_por_id(self, elemento_id):
        """ Obtém o texto de um elemento pelo ID """
        return self.driver.find_element(By.ID, elemento_id).text
    
    def obter_texto_por_xpath(self, xpath):
        """ Obtém o texto de um elemento pelo XPath """
        return self.driver.find_element(By.XPATH, xpath).text

    def digitar_por_xpath(self, xpath, texto):
        """ Digita um texto em um campo de entrada pelo XPath """
        self.driver.find_element(By.XPATH, xpath).send_keys(texto)

    def digitar_por_id(self, elemento_id, texto):
        """ Digita um texto em um campo de entrada pelo ID """
        self.driver.find_element(By.ID, elemento_id).send_keys(texto)

    def clicar_por_xpath(self, xpath):
        """ Clica em um elemento pelo XPath """
        self.driver.find_element(By.XPATH, xpath).click()

    def clicar_por_id(self, elemento_id):
        """ Clica em um elemento pelo ID """
        self.driver.find_element(By.ID, elemento_id).click()

    def clicar_por_nome(self, nome):
        """ Clica em um elemento pelo Name """
        self.driver.find_element(By.NAME, nome).click()

    def clicar_link_por_texto(self, texto):
        """ Clica em um link pelo texto visível """
        self.driver.find_element(By.LINK_TEXT, texto).click()

    def executar_javascript(self, script):
        """ Executa um script JavaScript na página """
        return self.driver.execute_script(script)

    def alterar_frame_por_texto(self, frame_name=""):
        """ Altera para um frame específico pelo nome """
        if frame_name:
            self.driver.switch_to.frame(frame_name)
        else:
            self.driver.switch_to.default_content()

    def alterar_para_default_content(self):
        self.driver.switch_to.default_content()

    def alterar_frame_por_index(self, index):
        """ Altera para um frame pelo índice """
        self.driver.switch_to.frame(index)

    def aguardar_estado_documento(self, timeout=10):
        """ Aguarda até que a página esteja completamente carregada """
        WebDriverWait(self.driver, timeout).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )

    def selecionar_item_lista_por_nome(self, nome, texto):
        """ Seleciona um item em uma lista suspensa pelo Name """
        select = Select(self.driver.find_element(By.NAME, nome))
        select.select_by_visible_text(texto)

    def selecionar_item_lista_por_id(self, elemento_id, texto):
        """ Seleciona um item em uma lista suspensa pelo ID """
        select = Select(self.driver.find_element(By.ID, elemento_id))
        select.select_by_visible_text(texto)

    def validar_checkbox_por_id(self, checkbox_id):
        """ Retorna True se um checkbox estiver marcado """
        return self.driver.find_element(By.ID, checkbox_id).is_selected()

    def aceitar_alerta(self):
        """ Aceita um alerta (popup) """
        try:
            WebDriverWait(self.driver, 5).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
        except:
            pass  # Nenhum alerta encontrado

    def rejeitar_alerta(self):
        """ Rejeita um alerta (popup) """
        try:
            WebDriverWait(self.driver, 5).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.dismiss()
        except:
            pass  # Nenhum alerta encontrado

    def obter_texto_alerta(self):
        """ Obtém o texto de um alerta (popup) """
        return Alert(self.driver).text

    def trocar_para_nova_aba(self):
        """ Troca para a última aba aberta """
        self.driver.switch_to.window(self.driver.window_handles[-1])

    def voltar_para_aba_anterior(self):
        """ Volta para a aba anterior """
        if len(self.driver.window_handles) > 1:
            self.driver.switch_to.window(self.driver.window_handles[0])

    def tirar_print(self, caminho_arquivo):
        """ Tira um print da tela e salva no caminho especificado """
        self.driver.save_screenshot(caminho_arquivo)

    def fechar_nova_aba(self):
        """ Fecha a aba atual e retorna para a aba anterior """
        if len(self.driver.window_handles) > 1:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    def aguardar_elemento(self, by, valor, tempo=10):
        """ Aguarda um elemento específico estar visível """
        return WebDriverWait(self.driver, tempo).until(
            EC.presence_of_element_located((by, valor))
        )

    def salvar_cookies(self, path="cookies.pkl"):
        """ Salva os cookies da sessão em um arquivo """
        with open(path, "wb") as file:
            pickle.dump(self.driver.get_cookies(), file)

    def carregar_cookies(self, path="cookies.pkl"):
        """ Carrega cookies de um arquivo e adiciona ao navegador """
        if os.path.exists(path):
            with open(path, "rb") as file:
                cookies = pickle.load(file)
                for cookie in cookies:
                    self.driver.add_cookie(cookie)

    def refresh(self):
        """ Recarrega a página atual """
        self.driver.refresh()
    

class SeleniumFirefoxAutomator:
    def __init__(self, driver_path=r"W:\PY_017 - Extratos Bradesco\7. Scripts\Driver FireFox\geckodriver.exe"):
        """ Inicializa o navegador Firefox com configurações específicas """

        # Redireciona logs para evitar saída desnecessária no terminal
        sys.stdout = open(os.devnull, "w")
        sys.stderr = open(os.devnull, "w")

        options = webdriver.FirefoxOptions()
        options.set_preference("browser.download.folderList", 2)  # Define diretório de download
        options.set_preference("browser.download.dir", r"W:\PY_017 - Extratos Bradesco\Downloads")
        options.set_preference("browser.download.useDownloadDir", True)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")

        # Inicializa o WebDriver do Firefox
        service = Service(driver_path)
        self.driver = webdriver.Firefox(service=service, options=options)

        # Restaurar stdout e stderr após iniciar o driver
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

    ### 🔹 CONTROLE DO NAVEGADOR
    def navegar_para(self, url):
        """ Abre uma URL no Firefox """
        self.driver.get(url)

    def fechar_navegador(self):
        """ Fecha o navegador """
        self.driver.quit()

    def maximizar_janela(self):
        """ Maximiza a janela do navegador """
        self.driver.maximize_window()

    def refresh(self):
        """ Recarrega a página atual """
        self.driver.refresh()

    def obter_titulo(self):
        """ Retorna o título da página atual """
        return self.driver.title

    ### 🔹 INTERAÇÃO COM ELEMENTOS
    def obter_texto_por_id(self, elemento_id):
        """ Obtém o texto de um elemento pelo ID """
        return self.driver.find_element(By.ID, elemento_id).text

    def obter_texto_por_xpath(self, xpath):
        """ Obtém o texto de um elemento pelo XPath """
        return self.driver.find_element(By.XPATH, xpath).text

    def clicar_por_xpath(self, xpath):
        """ Clica em um elemento pelo XPath """
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

    def clicar_por_id(self, elemento_id):
        """ Clica em um elemento pelo ID """
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, elemento_id))).click()

    def clicar_por_nome(self, nome):
        """ Clica em um elemento pelo Name """
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.NAME, nome))).click()

    def clicar_link_por_texto(self, texto):
        """ Clica em um link pelo texto visível """
        self.driver.find_element(By.LINK_TEXT, texto).click()

    def digitar_por_xpath(self, xpath, texto):
        """ Digita um texto em um campo de entrada pelo XPath """
        elemento = self.driver.find_element(By.XPATH, xpath)
        elemento.clear()
        elemento.send_keys(texto)

    def digitar_por_id(self, elemento_id, texto):
        """ Digita um texto em um campo de entrada pelo ID """
        elemento = self.driver.find_element(By.ID, elemento_id)
        elemento.clear()
        elemento.send_keys(texto)

    def limpar_campo_por_xpath(self, xpath):
        """ Limpa um campo de entrada pelo XPath """
        self.driver.find_element(By.XPATH, xpath).clear()

    def enviar_tab_por_id(self, elemento_id):
        """ Envia a tecla TAB para um campo de entrada pelo ID """
        self.driver.find_element(By.ID, elemento_id).send_keys(Keys.TAB)

    def enviar_arquivo_por_id(self, elemento_id, caminho_arquivo):
        """ Envia um arquivo para um input type='file' """
        self.driver.find_element(By.ID, elemento_id).send_keys(caminho_arquivo)

    ### 🔹 ESPERAS E CARREGAMENTO DE PÁGINA
    def aguardar_estado_documento(self, timeout=10):
        """ Aguarda até que a página esteja completamente carregada """
        WebDriverWait(self.driver, timeout).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )

    def aguardar_elemento(self, by, valor, tempo=10):
        """ Aguarda um elemento específico estar visível """
        return WebDriverWait(self.driver, tempo).until(
            EC.presence_of_element_located((by, valor))
        )

    ### 🔹 MANIPULAÇÃO DE ALERTAS
    def aceitar_alerta(self):
        """ Aceita um alerta (popup) """
        try:
            WebDriverWait(self.driver, 5).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
        except:
            pass

    def rejeitar_alerta(self):
        """ Rejeita um alerta (popup) """
        try:
            WebDriverWait(self.driver, 5).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.dismiss()
        except:
            pass

    def obter_texto_alerta(self):
        """ Obtém o texto de um alerta (popup) """
        return Alert(self.driver).text

    ### 🔹 MANIPULAÇÃO DE ABAS
    def trocar_para_nova_aba(self):
        """ Troca para a última aba aberta """
        self.driver.switch_to.window(self.driver.window_handles[-1])

    def voltar_para_aba_anterior(self):
        """ Volta para a aba anterior """
        if len(self.driver.window_handles) > 1:
            self.driver.switch_to.window(self.driver.window_handles[0])

    def fechar_nova_aba(self):
        """ Fecha a aba atual e retorna para a aba anterior """
        if len(self.driver.window_handles) > 1:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    ### 🔹 MANIPULAÇÃO DE FRAMES
    def mudar_frame_por_id(self, frame):
        """ Alterna para um frame pelo ID """
        self.driver.switch_to.frame(frame)

    def alterar_frame_por_index(self, index):
        """ Altera para um frame pelo índice """
        self.driver.switch_to.frame(index)

    def alterar_para_default_content(self):
        """ Retorna ao contexto padrão do navegador """
        self.driver.switch_to.default_content()

    ### 🔹 COOKIES
    def salvar_cookies(self, path="cookies.pkl"):
        """ Salva os cookies da sessão em um arquivo """
        with open(path, "wb") as file:
            pickle.dump(self.driver.get_cookies(), file)

    def carregar_cookies(self, path="cookies.pkl"):
        """ Carrega cookies de um arquivo e adiciona ao navegador """
        if os.path.exists(path):
            with open(path, "rb") as file:
                cookies = pickle.load(file)
                for cookie in cookies:
                    self.driver.add_cookie(cookie)
