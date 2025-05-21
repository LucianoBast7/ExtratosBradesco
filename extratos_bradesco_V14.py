# Importação de Bibliotecas
import os
import time
import sys
import logging
import pandas as pd
from datetime import date
import requests
import subprocess
import speech_recognition as sr
import warnings
warnings.filterwarnings("ignore")
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from sqlalchemy import create_engine
from datetime import datetime, date, timedelta

from variaveis.variaveis_estratos_bradesco import Variaveis
from requirements.vortx_calendar import VortxCalendar
from requirements.funcoes_selenium import SeleniumAutomator

bot = SeleniumAutomator()

# Listas de empresas a serem ignoradas
empresas_ignoradas = {}

# Configuração de Logs
data_atual = date.today()
data_menos_1 = data_atual - timedelta(days=1)
data_em_texto = str(data_menos_1)
hoje_atual = date.today().strftime("%Y-%m-%d")
hoje = datetime.today()
ano = hoje.strftime("%Y")
mes = hoje.strftime("%m")
dia = hoje.strftime("%d")

caminho_log = f"W:\\PY_017 - Extratos Bradesco\\0. Log\\{ano}\\{mes}\\{dia}\\ExtratosBradesco_{data_atual}.log"

log_format = '%(asctime)s:%(levelname)s:%(filename)s:%(message)s'
logging.basicConfig(filename=caminho_log, filemode='a', level=logging.INFO, format=log_format)

def add_log(msg: str, level: str = 'info'):
    getattr(logging, level)(msg)

# Carregamento de Variáveis
def carregar_variaveis():
    variaveis = Variaveis()
    return {
        "pasta_analise_de_caixa": variaveis.pasta_analise_de_caixa,
        "pasta_exportacao_extratos": variaveis.pasta_exportacao_extratos,
        "user_bradesco": variaveis.user_bradesco,
        "senha_bradesco": variaveis.senha_bradesco,
        "calendario": VortxCalendar(),
        "host": variaveis.host_postgre,
        "usuario": variaveis.user_postgre,
        "senha": variaveis.senha_postgre,
        "database": variaveis.database_postgre,
        "pasta_input": variaveis.pasta_input
    }

variaveis = carregar_variaveis()

# Verificação de Dia Útil
def verificar_dia_util(calendario):
    if not calendario.is_working_day(data_atual):
        add_log("HOJE NÃO É DIA ÚTIL, PARANDO EXECUÇÃO...", 'info')
        sys.exit(0)
    add_log("HOJE É DIA ÚTIL, INICIANDO EXECUÇÃO...")

verificar_dia_util(variaveis["calendario"])

def main():
    # Inicialização do Selenium
    def iniciar_selenium():
        add_log("INICIANDO BRADESCO")
        print("INICIANDO BRADESCO")
        bot.navegar_para("https://www.ne12.bradesconetempresa.b.br/ibpjlogin/login.jsf")
        bot.aguardar_estado_documento()
        bot.maximizar_janela()
        return bot
    
    iniciar_selenium()

    # Realizar Login
    def realizar_login(user, senha):
        add_log("REALIZANDO LOGIN")
        print("REALIZANDO LOGIN")
        add_log("INSERINDO USER")
        bot.digitar_por_xpath('//*[@id="identificationForm:txtUsuario"]', user)
        add_log("INSERINDO SENHA")
        bot.digitar_por_xpath('//*[@id="identificationForm:txtSenha"]', senha)
        add_log("CLICA EM ENTER")
        bot.digitar_por_xpath('//*[@id="identificationForm:botaoAvancar"]', Keys.ENTER)
        bot.aguardar_estado_documento()
        time.sleep(10)
        try:
            add_log("VERIFICANDO SE MSG DE ACESSO FINALIZADO APARECE")
            mensagem_erro = bot.obter_texto_por_id("modalBoxAlertTexto")
            if mensagem_erro:
                add_log("FECHANDO MSG DE ACESSO FINALIZADO")
                bot.clicar_por_id("btnFecharModal")
                bot.aguardar_estado_documento()
                time.sleep(5)
                add_log("VOLTANDO PARA TELA DE LOGIN")
                bot.aguardar_elemento_para_clicar_por_id("btnAcessoContaTopBar")
                bot.aguardar_elemento(By.XPATH, '//*[@id="identificationForm:txtUsuario"]')
                bot.refresh()
                add_log("INSERINDO USER")
                bot.digitar_por_xpath('//*[@id="identificationForm:txtUsuario"]', user)
                add_log("INSERINDO SENHA")
                bot.digitar_por_xpath('//*[@id="identificationForm:txtSenha"]', senha)
                add_log("CLICA EM ENTER")
                bot.digitar_por_xpath('//*[@id="identificationForm:botaoAvancar"]', Keys.ENTER)
                bot.aguardar_estado_documento()
                time.sleep(5)
                add_log("LOGIN REALIZADO COM SUCESSO, SEGUINDO PARA RECAPTCHA")
            else:
                print("LOGIN REALIZADO COM SUCESSO, SEGUINDO PARA RECAPTCHA")
                pass
        except Exception as e:
            print(f"ERRO:{e}")

    realizar_login(variaveis["user_bradesco"], variaveis["senha_bradesco"])
    time.sleep(10)
        
    # Quebrar Captcha
    def quebrar_captcha(user_bradesco, senha_bradesco):
        add_log("VERIFICANDO COMPONENTE DE SEGURANÇA")
        btn_seguranca = bot.obter_texto_por_xpath('//*[@id="conteudo_interno_oferta_OFDB"]/div[1]/div[1]/div/div[1]/span')        
        if btn_seguranca:
            add_log("CLICANDO EM INSTALAR MAIS TARDE")
            bot.clicar_por_xpath('//*[@id="formInstalarOFDB:btn-avancar-left"]')
            bot.aguardar_estado_documento()
            time.sleep(5)
            add_log("MSG DE COMPONENTE DE SEGURANÇA FECHADA")
        else:
            add_log("MSG DE COMPONENTE DE SEGURANÇA NÃO APARECEU")
            pass
                
        add_log("INICIANDO RESOLUÇÃO DO RECAPTCHA")        
        bot.refresh()
        add_log("CLICANDO NO CHECK BOX")
        bot.aguardar_frame_e_trocar("//iframe[contains(@title, 'reCAPTCHA')]")
        bot.aguardar_elemento_para_clicar_por_xpath("//div[@class='recaptcha-checkbox-border']")
        bot.driver.switch_to.default_content()
        time.sleep(5)
        try:
            add_log("VERIFICA MSG DE ERRO AO RESOLVER RECAPTCHA AUTOMATICAMENTE")
            erro = bot.obter_texto_por_xpath('//*[@id="erroSenhaVazia"]/strong/span')
            if "Erro na validação do token recaptcha." in erro:
                add_log("CLICANDO EM CANCELAR ACESSO")
                bot.clicar_por_xpath('//*[@id="formTeste:_id107"]')
                time.sleep(5)
                bot.aguardar_elemento_para_clicar_por_id("btnAcessoContaTopBar")
                add_log("VOLTANDO PARA TELA DE LOGIN")
                bot.aguardar_elemento(By.XPATH, '//*[@id="identificationForm:txtUsuario"]')
                bot.refresh()
                add_log("INSERINDO USER")
                bot.digitar_por_xpath('//*[@id="identificationForm:txtUsuario"]', user_bradesco)
                add_log("INSERINDO SENHA")
                bot.digitar_por_xpath('//*[@id="identificationForm:txtSenha"]', senha_bradesco)
                add_log("CLICA EM ENTER")
                bot.digitar_por_xpath('//*[@id="identificationForm:botaoAvancar"]', Keys.ENTER)
                bot.aguardar_estado_documento()
                time.sleep(15)
                add_log("LOGIN REALIZADO COM SUCESSO, SEGUINDO PARA RECAPTCHA")
                return quebrar_captcha(variaveis["user_bradesco"], variaveis["senha_bradesco"])
        except Exception as e:
            add_log("MSG DE ERRO NÃO LOCALIZADA, SEGUINDO")
        time.sleep(10)
        try:
            time.sleep(10)
            add_log("VERIFICA O RECAPTCHA FOI RESOLVIDO AUTOMATICAMENTE")
            saldos_extratos = bot.obter_texto_por_id("_id74_0:_id76")
            if saldos_extratos:
                print("CAPTCHA RESOLVIDO AUTOMATICAMENTE, CONTINUANDO...")
                add_log("CAPTCHA RESOLVIDO AUTOMATICAMENTE")
        except:
            print("CAPTCHA AINDA PRECISA SER RESOLVIDO")
            add_log("CAPTCHA AINDA PRECISA SER RESOLVIDO")
            # Voltar para o contexto principal e entrar no iframe do desafio
            bot.aguardar_frame_e_trocar("//iframe[contains(@title, 'desafio reCAPTCHA')]")
            # Clicar na opção de áudio
            add_log("ESCOLHE A OPÇÃO DE RESOLUÇÃO POR AUDIO")
            bot.aguardar_elemento_para_clicar_por_id("recaptcha-audio-button")

            def capturar_audio():
                # Capturar o link do áudio
                add_log("CAPTURA O AUDIO")
                audio_src = bot.aguardar_link_para_capturar_por_xpath("//*[@id='rc-audio']/div[7]/a")
                print(f"Link do áudio encontrado: {audio_src}")

                # Baixar o áudio do CAPTCHA
                audio_mp3 = "captcha_audio.mp3"
                response = requests.get(audio_src)
                with open(audio_mp3, "wb") as file:
                    file.write(response.content)

                # Verificar se o arquivo MP3 foi baixado corretamente
                if not os.path.exists(audio_mp3) or os.path.getsize(audio_mp3) == 0:
                    print("Erro: O arquivo de áudio não foi baixado corretamente.")
                    bot.driver.quit()
                    exit()

                # Converter MP3 para WAV com FFmpeg (taxa de amostragem ajustada)
                audio_wav = "captcha_audio.wav"
                print("Convertendo MP3 para WAV...")
                subprocess.run(["ffmpeg", "-i", audio_mp3, "-ar", "22050", "-ac", "1", "-y", audio_wav],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

                # Verificar se o arquivo WAV foi gerado corretamente
                if not os.path.exists(audio_wav) or os.path.getsize(audio_wav) == 0:
                    print("Erro: A conversão para WAV falhou.")
                    bot.driver.quit()
                    exit()

                # Tentativa 1: Reconhecer o áudio com SpeechRecognition
                recognizer = sr.Recognizer()
                captcha_text = None

                try:
                    with sr.AudioFile(audio_wav) as source:
                        recognizer.adjust_for_ambient_noise(source, duration=0.5)
                        audio_data = recognizer.record(source)
                        captcha_text = recognizer.recognize_google(audio_data, language="en-US")
                        print(f"Texto reconhecido pelo Google Speech: {captcha_text}")
                        add_log(f"TEXTO RECONHECIDO: {captcha_text}")
                except sr.UnknownValueError:
                    print("Falha no reconhecimento pelo Google Speech. Tentando Novamente")
                    return quebrar_captcha(variaveis["user_bradesco"], variaveis["senha_bradesco"])
                except sr.RequestError as e:
                    print(f"Erro na API Google Speech: {e}")
                    
                time.sleep(5)

                try:
                    # Inserir o texto no campo do CAPTCHA
                    add_log("INSERINDO TEXTO NO VALIDADOR")
                    audio_input = WebDriverWait(bot.driver, 10).until(EC.element_to_be_clickable((By.ID, "audio-response")))
                    audio_input.send_keys(captcha_text)
                    time.sleep(1)

                    # Clicar no botão de verificar
                    add_log("VERIFICANDO TEXTO")
                    verify_button = WebDriverWait(bot.driver, 10).until(EC.element_to_be_clickable((By.ID, "recaptcha-verify-button")))
                    verify_button.click()                        
                    time.sleep(2)
                    try:
                        add_log("RECAPTCHA RESOLVIDO... VERIFICANDO SE MSG DE ERRO APARECE")
                        erro = bot.obter_texto_por_xpath('//*[@id="erroSenhaVazia"]/strong/span')
                        if "Erro na validação do token recaptcha." in erro:
                            add_log("CANCELANDO ACESSO")
                            bot.clicar_por_xpath('//*[@id="formTeste:_id107"]')
                            time.sleep(5)
                            add_log("VOLTANDO PARA TELA DE LOGIN")
                            bot.aguardar_elemento_para_clicar_por_id("btnAcessoContaTopBar")
                            bot.aguardar_elemento(By.XPATH, '//*[@id="identificationForm:txtUsuario"]')
                            bot.refresh()
                            add_log("INSERINDO USER")
                            bot.digitar_por_xpath('//*[@id="identificationForm:txtUsuario"]', user_bradesco)
                            add_log("INSERINDO SENHA")
                            bot.digitar_por_xpath('//*[@id="identificationForm:txtSenha"]', senha_bradesco)
                            add_log("CLICA EM ENTER")
                            bot.digitar_por_xpath('//*[@id="identificationForm:botaoAvancar"]', Keys.ENTER)
                            bot.aguardar_estado_documento()
                            time.sleep(15)
                            add_log("LOGIN REALIZADO COM SUCESSO, SEGUINDO PARA RECAPTCHA")
                            return quebrar_captcha(variaveis["user_bradesco"], variaveis["senha_bradesco"])
                    except Exception as e:
                        add_log("RECAPTCHA RESOLVIDO SEM MSG DE ERRO, SEGUINDO")
                except Exception as e:
                    add_log(f"ERRO AO VALIDAR RECAPTCHA: {e}")

            capturar_audio()
            time.sleep(10)
        
    quebrar_captcha(variaveis["user_bradesco"], variaveis["senha_bradesco"])

    def acessar_extratos():
        add_log("ACESSANDO EXTRATOS")
        print("ACESSANDO EXTRATOS")
        time.sleep(10)
        bot.refresh()
        try:
            botao_aviso = bot.aguardar_elemento(By.ID, "_id75")
            if botao_aviso:
                print("EXCLUINDO BOX DE AVISO")
                bot.clicar_por_id("_id75")
                time.sleep(5)
                return acessar_extratos()
        except:
            print("SEM AVISOS, SEGUINDO...")
            time.sleep(5)
            bot.aguardar_elemento(By.ID, "_id74_0:_id76")
            bot.clicar_por_id("_id74_0:_id76")
            bot.aguardar_estado_documento()
            bot.aguardar_elemento(By.LINK_TEXT, "Extrato (Últimos Lançamentos)")
            bot.clicar_por_link_text("Extrato (Últimos Lançamentos)")
            bot.aguardar_estado_documento()
            bot.aguardar_elemento(By.ID, 'paginaCentral')
            bot.mudar_frame_por_id('paginaCentral')
            bot.aguardar_estado_documento()
        
    acessar_extratos()

    def obter_lista_empresas():
        try:
            time.sleep(5)
            try:
                add_log("LOCALIZANDO LISTA DE FUNDOS")
                print("LOCALIZANDO LISTA DE FUNDOS")
                select_element = WebDriverWait(bot.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//select[contains(@id, 'comboEmpresas')]"))
                )
                print("ELEMENTO <select> ENCONTRADO")
            except:
                print("ElEMENTO <select> NÃO ENCONTRADO")

            if select_element:
                add_log("CRIANDO LISTA COM FUNDOS")
                print("CRIANDO LISTA COM FUNDOS")
                bot.driver.execute_script("arguments[0].style.display = 'block';", select_element)
                select_html = select_element.get_attribute("outerHTML")

                soup = BeautifulSoup(select_html, "html.parser")
                options = soup.find_all("option")

                lista_empresas = [
                    {"value": option["value"], "text": option.text.strip()}
                    for option in options
                    if option.text.strip() not in empresas_ignoradas
                ]
                lista_df = pd.DataFrame(lista_empresas)
                lista_df.to_excel("W:\\PY_017 - Extratos Bradesco\\7. Scripts\\lista_fundos.xlsx", index=False, sheet_name="Lista Fundos")
                return lista_empresas
            else:
                print("ERRO AO CRIAR LISTA COM FUNDOS")
                return []

        except Exception as e:
            print(f"ERRO: {e}")
            return []
    
    def baixar_extratos():
        add_log("INICIANDO EXTRAÇÃO DOS EXTRATOS")
        print("\nINICIANDO EXTRAÇÃO DOS EXTRATOS")
        empresas = obter_lista_empresas()
        relatorio = []
        movimentacoes = []

        if not empresas:
            print("NENHUM FUNDO ENCONTRADO")
            return
        
        empresas_filtradas = [empresa for empresa in empresas if empresa["text"].split("|")[0].strip() not in empresas_ignoradas]
        
        for empresa in empresas_filtradas:
            empresa_nome = empresa["text"].split("|")[0].strip()
            empresa_cnpj = empresa["text"].split("|")[1].replace("CNPJ ", "")
            empresa_id = empresa["value"]

            add_log(f"\nSELECIONANDO FUNDO: {empresa_nome}")
            print(f"\nSELECIONANDO FUNDO: {empresa_nome}")

            try:
                select_element = WebDriverWait(bot.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "formFiltroUltimosLancamentos:filtro:comboEmpresas"))
                )
                select = Select(select_element)
                select.select_by_value(empresa_id)
                bot.aguardar_estado_documento()
                print("PÁGINA DO EXTRATO CARREGADA COMPLETAMENTE")
                time.sleep(5)

                # Capturar Conta
                try:
                    time.sleep(1)
                    conta_element = WebDriverWait(bot.driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//a[contains(@id, 'listaContasFiltro')]"))
                    )
                    agencia = conta_element.find_element(By.XPATH, ".//span[1]").text
                    numero_conta = conta_element.find_element(By.XPATH, ".//span[2]").text
                    digito_conta = conta_element.find_element(By.XPATH, ".//span[4]").text
                    conta_texto = f"{agencia} | {numero_conta}"
                    conta = f"{numero_conta}-{digito_conta}"
                    print(f"CONTA {conta_texto} CAPTURADA!")
                except:
                    conta_texto = "CONTA NÃO DISPONÍVEL"
                
                # **Verificar se a tabela de movimentação ("tabelaSaldos") existe**
                try:
                    time.sleep(1)
                    tabela_saldos = bot.driver.find_element(By.CLASS_NAME, "tabelaSaldos")
                    houve_movimentacao = "Sim"
                    saldo_java_script = bot.driver.execute_script("""
                        return document.querySelector("#formularioUltimosLancamentos\\\\:listagem\\\\:_id1931 > div.fontImpressao > div > table > tbody.bordaTotal.total.ignorarOrdenacao.tabelaSaldosTotal > tr > td.tabelaSaldosTd.alignRight.positivo").innerText;
                    """)

                    saldo_final = saldo_java_script.replace(".", "").replace(",", ".")
                    print(f"SALDO TabelaSaldosTd: {saldo_final}")
                    print(f"HOUVE MOVIMENTAÇÃO? {houve_movimentacao}")

                    # Capturando Movimentações
                    tabela_lancamentos = tabela_saldos.find_element(By.TAG_NAME, "tbody")
                    linhas = tabela_lancamentos.find_elements(By.TAG_NAME, "tr")

                    ultima_data = None

                    for linha in linhas:
                        colunas = linha.find_elements(By.TAG_NAME, "td")

                        if len(colunas) >= 6:
                            data = colunas[0].text.strip()
                            lancamento = colunas[1].text.strip()
                            movimento_credito = colunas[4].text.strip().replace(".", "").replace(",", ".")
                            movimento_debito = colunas[5].text.strip().replace(".", "").replace(",", ".")

                            if data:
                                ultima_data = data
                            else:
                                data = ultima_data

                            if movimento_credito:
                                movimento = float(movimento_credito)
                            elif movimento_debito:
                                movimento = -abs(float(movimento_debito))
                            else:
                                movimento = 0.0

                            movimentacoes.append({
                                "Agencia": agencia,
                                "Conta": conta,
                                "CNPJ": empresa_cnpj,
                                "Data": data,
                                "Lançamento": lancamento,
                                "Movimento": movimento
                            })
                except:
                    # Se não há movimentação, pegamos o Saldo da Conta-Corrente
                    try:
                        time.sleep(1)
                        bot.aguardar_elemento_para_clicar_por_xpath("//a[.//span[text()='Veja mais']]")
                        bot.aguardar_elemento(By.XPATH, "//tr[td[text()='Conta-Corrente']]/td[@class='alignRight valor']")
                        houve_movimentacao = "Não"

                        # Capturar saldo da coluna "Conta-Corrente"
                        saldo_total_element = bot.driver.find_element(By.XPATH, "//tr[td[text()='Conta-Corrente']]/td[@class='alignRight valor']")
                        saldo_final = saldo_total_element.text.replace(".", "").replace(",", ".")
                        print(f"SALDO DISPONÍVEL: {saldo_final}")
                        print(f"HOUVE MOVIMENTAÇÃO: {houve_movimentacao}")

                    except:
                        # Se nenhuma tabela existir, saldo = 0
                        print(f"EXTRATO DO FUNDO {empresa_nome} INDISPONÍVEL.")
                        houve_movimentacao = "Não"
                        saldo_final = "Indisponível"

                # Armazenar os dados no relatório
                time.sleep(1)
                relatorio.append({
                    "Nome do Fundo": empresa_nome,
                    "CNPJ": empresa_cnpj,
                    "Agencia": agencia,
                    "Conta": conta,
                    "Data": data_atual,
                    "Houve Movimentação": houve_movimentacao,
                    "Saldo": saldo_final
                })

                add_log(f"FUNDO ({empresa_nome}) PROCESSADO COM SUCESSO")
                print(f"FUNDO ({empresa_nome}) PROCESSADO COM SUCESSO")

            except Exception as e:
                print(f"ERRO AO PROCESSAR FUNDO {empresa_nome}: {e}")
                add_log(f"ERRO AO PROCESSAR FUNDO {empresa_nome}: {e}")
        
        def sair_site_bradesco():
            try:
                # Sair do Bradesco
                bot.driver.switch_to.default_content()
                bot.clicar_por_xpath('//*[@id="botaoSair"]')
            except:
                bot.fechar_navegador()

        sair_site_bradesco()

        # Merge com Mapa Sinqia
        mapa_sinqia = pd.read_excel(f"{variaveis['pasta_input']}\\mapa_sinqia.xlsx", sheet_name="Sheet", dtype={"Carteira": str})
        mapa_sinqia["CNPJ"] = mapa_sinqia["CNPJ"].astype(str).str.strip()
        mapa_sinqia["CNPJ"].str.replace(r"[./-]", "", regex=True)
        mapa_sinqia["CNPJ"] = mapa_sinqia["CNPJ"].astype(str).str.zfill(15)

        # Inserir Saldos no banco 
        df_relatorio = pd.DataFrame(relatorio)             
        df_relatorio["CNPJ"] = df_relatorio["CNPJ"].astype(str).str.strip()
        df_relatorio["CNPJ"] = df_relatorio["CNPJ"].str.replace(r"[./-]", "", regex=True)
        df_relatorio_final = df_relatorio.copy()
        df_relatorio = df_relatorio[[
            "Agencia",
            "Conta",
            "CNPJ",
            "Data",
            "Saldo"
        ]]
        df_relatorio = df_relatorio[df_relatorio['Saldo']!= 'Indisponível']
        engine = create_engine(f"postgresql://{variaveis['usuario']}:{variaveis['senha']}@{variaveis['host']}:5432/{variaveis['database']}")
        df_relatorio.to_sql('saldosBradesco', engine, if_exists='append', index=False, schema='public')
        print("\nSALDOS SALVOS")
        add_log("SALDOS SALVOS")

        # Merge com Sinqia
        df_final = pd.merge(
            df_relatorio_final,
            mapa_sinqia,
            how="outer",
            left_on="CNPJ",
            right_on="CNPJ"
        )
        df_final["Carteira"] = df_final["Carteira"].fillna("Sem Carteira no Mapa")
        df_final["Agencia"] = df_final["Agencia"].fillna("Excluir")
        df_final.query('Agencia != "Excluir"')
        df_final = df_final.dropna(how="any")
        df_final.to_excel(f"{variaveis['pasta_exportacao_extratos']}\\ExtratosBradesco.xlsx", index=False, sheet_name="Saldos")
        print("EXTRATOS EXPORTADOS COM SUCESSO")
        add_log("EXTRATOS EXPORTADOS COM SUCESSO")

        # Inserir Movimentações no Banco 
        df_movimentacoes = pd.DataFrame(movimentacoes)
        df_movimentacoes = df_movimentacoes[df_movimentacoes['Lançamento']!= 'SALDO ANTERIOR']
        df_movimentacoes["CNPJ"] = df_movimentacoes["CNPJ"].astype(str).str.strip()
        df_movimentacoes["CNPJ"] = df_movimentacoes["CNPJ"].str.replace(r"[./-]", "", regex=True) 
        df_movimentacoes.rename(columns={"Lançamento": "Lancamento"}, inplace=True)
        df_movimentacoes["Data"] = pd.to_datetime(df_movimentacoes["Data"], format='%d/%m/%Y').dt.strftime('%Y-%m-%d')
        engine = create_engine(f"postgresql://{variaveis['usuario']}:{variaveis['senha']}@{variaveis['host']}:5432/{variaveis['database']}")
        df_movimentacoes.to_sql('movimentacoesBradesco', engine, if_exists='append', index=False, schema='public')
        print("\nMOVIMENTAÇÕES SALVAS")
        add_log("MOVIMENTAÇÕES SALVAS")

        # Merge com Sinqia
        df_final_2 = pd.merge(
            df_movimentacoes,
            mapa_sinqia,
            how="outer",
            left_on="CNPJ",
            right_on="CNPJ"
        )
        df_final_2["Carteira"] = df_final_2["Carteira"].fillna("Sem Carteira no Mapa")
        df_final_2["Agencia"] = df_final_2["Agencia"].fillna("Excluir")
        df_final_2.query('Agencia != "Excluir"')
        df_final_2 = df_final_2.dropna(how="any")
        df_final_2.to_excel(f"{variaveis['pasta_exportacao_extratos']}\\MovimentacoesBradesco.xlsx", index=False, sheet_name="Movimentações")
        
        # Criar Layout Sinqia
        df_mov_sinqia = df_final_2.copy()
        df_mov_sinqia["Lancamento"] = df_mov_sinqia["Lancamento"].fillna("").astype(str).str.strip()
        df_mov_sinqia = df_mov_sinqia[df_mov_sinqia["Lancamento"].str.contains(r"MANUTENCAO", case=False, na=False, regex=True) |
                                      df_mov_sinqia["Lancamento"].str.contains(r"TED INTERNET", case=False, na=False, regex=True)]
        df_mov_sinqia = df_mov_sinqia[[
            "Data",
            "Movimento",
            "CNPJ",
            "Carteira"
        ]]
        df_mov_sinqia = df_mov_sinqia[df_mov_sinqia["Data"] == hoje_atual]

        # Variaveis String
        Carteiras = df_mov_sinqia["Carteira"]
        Valores = df_mov_sinqia["Movimento"]
        Datas = df_mov_sinqia["Data"]

        # Definindo tamanho da String
        carteira_size = 9
        valor_size = 16
        data_size = 11

        # Criando arquivo Sinqia
        arquivo = open(f"{variaveis['pasta_exportacao_extratos']}\\layout_sinqia_{data_atual}.txt", 'w')

        for linha in df_mov_sinqia.index:
            carteira = str(Carteiras[linha])
            carteira = carteira.ljust(carteira_size)
            cod_op = "400 "
            valor = f"{Valores[linha]:.2f}".replace(".", ",")
            valor = valor.ljust(valor_size)
            tipo_conta = "C/C          "
            tipo_lancamento = "#TARIFBANC   "
            data = str(Datas[linha]).replace("-", "")
            data = data.ljust(data_size)
            descricao_lancamento = "Tarifa Bancaria                           ;\n"

            # Preenchendo o Arquivo
            arquivo.write(carteira)
            arquivo.write(cod_op)
            arquivo.write(valor)
            arquivo.write(tipo_conta)
            arquivo.write(tipo_lancamento)
            arquivo.write(data)
            arquivo.write(descricao_lancamento)
        
        # Salvar Arquivo
        arquivo.close()  
        
        print("MOVIMENTAÇÕES EXPORTADAS COM SUCESSO")
        add_log("MOVIMENTAÇÕES EXPORTADAS COM SUCESSO")  

    baixar_extratos()

main()



