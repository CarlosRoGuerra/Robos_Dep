import pandas as pd
from bs4 import BeautifulSoup
import time
import random
import re
import os
from fake_useragent import UserAgent
from concurrent.futures import ThreadPoolExecutor, as_completed
import psutil
import queue
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from requests_html import HTMLSession
import sys
import threading
from requests.exceptions import RequestException, Timeout

session = HTMLSession()
ua = UserAgent()
lock = threading.Lock()

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--disable-gpu')

if hasattr(sys, '_MEIPASS'):
    chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
else:
    chromedriver_path = 'chromedriver.exe'

service = Service(executable_path=chromedriver_path)

def create_driver():
    return webdriver.Chrome(service=service, options=chrome_options)

driver_pool = queue.Queue()

def get_driver():
    if driver_pool.empty():
        return create_driver()
    return driver_pool.get()

def release_driver(driver):
    driver_pool.put(driver)

def ler_planilha(caminho_planilha):
    try:
        df = pd.read_excel(caminho_planilha)
        df['FABRICANTE/MARCA'] = df['FABRICANTE/MARCA'].fillna('')
        
        if 'FABRICANTE/MARCA' not in df.columns or 'CODIGO' not in df.columns:
            raise ValueError("As colunas 'FABRICANTE/MARCA' e 'CODIGO' precisam estar presentes na planilha.")
        
        #combinacoes = (df['FABRICANTE/MARCA'].astype(str) + ' ' + df['CODIGO'].astype(str)).tolist()
        combinacoes = df['CODIGO'].astype(str).tolist()
        return combinacoes
    
    except ValueError as e:
        print(f"Erro: {e}")
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
    return []

def salvar_em_planilha(dados, caminho_arquivo):
    try:
        df_novos_dados = pd.DataFrame(dados)

        if os.path.exists(caminho_arquivo):
            df_existente = pd.read_excel(caminho_arquivo)
            df_combined = pd.concat([df_existente, df_novos_dados]).drop_duplicates(subset='Código do Fabricante').reset_index(drop=True)
        else:
            df_combined = df_novos_dados
        
        caminho_temp = caminho_arquivo + '.temp.xlsx'
        
        df_combined.to_excel(caminho_temp, index=False, engine='openpyxl')
        
        os.replace(caminho_temp, caminho_arquivo)
        
        print(f'Dados salvos com sucesso em {caminho_arquivo}')
    
    except Exception as e:
        print(f'Erro ao salvar dados na planilha: {e}')

def realizar_login(cnpj, senha, driver):
    login_url = 'https://www.kdapeca.com.br/'
    max_tentativas = 3
    tentativa = 0
    
    while tentativa < max_tentativas:
        try:
            driver.get(login_url)
            
            cnpj_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'cnpj'))
            )
            cnpj_input.send_keys(cnpj)
            
            senha_input = driver.find_element(By.NAME, 'password')
            senha_input.send_keys(senha)
            
            login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Acessar')]")
            login_button.click()
            
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "query"))
            )
            
            print("Login realizado com sucesso!")
            return True
        
        except Exception as e:
            print(f"Tentativa {tentativa + 1} falhou: {e}")
            tentativa += 1
            time.sleep(5)
    
    print("Falha no login após várias tentativas.")
    return False

def verificar_se_login_expirou(driver):
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, 'cnpj'))
        )
        return True
    except TimeoutException:
        return False

def obter_codigos_ja_capturados(caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        df = pd.read_excel(caminho_arquivo)
        return set(df['Código do Fabricante'].astype(str))
    return set()

def extrair_numero_codigo(codigo):
    return ''.join(re.findall(r'\d+', codigo))

def buscar_codigo_no_site(codigo_completo):
    driver = get_driver()
    try:
        df = pd.read_excel('kdpeça-teste1.xlsx')
        
        # Verificar se o código já existe na coluna "Código Original"
        if codigo_completo in df['Código Original'].values:
            print(f'O código {codigo_completo} já existe na planilha. Pulando para o próximo.')
            return None  # Pula para o próximo código_completo
        codigo_formatado = codigo_completo.replace('-', '').replace('/', '').replace('.', '')
        #if codigo_formatado.split(" ")[1] == "":
         #   codigo_puro = codigo_formatado.split(" ")[2]
          #  codigo_limpo = codigo_puro.strip()
          #  search_url = f'https://www.kdapeca.com.br/pesquisa?query={codigo_limpo}'
        #else:
        #    codigo_puro = codigo_formatado.split(" ")[1]
         #   codigo_limpo = codigo_puro.strip()
        search_url = f'https://www.kdapeca.com.br/pesquisa?query={codigo_formatado}'
        max_tentativas = 3
        tentativa = 0
        
        while tentativa < max_tentativas:
            try:
                driver.get(search_url)
                time.sleep(5)
                if verificar_se_login_expirou(driver):
                    print("Login expirado, realizando novo login...")
                    if not realizar_login('67671537000154', 'Dpk543@', driver):
                        print("Erro ao realizar novo login.")
                        return None
                # Verifica se a mensagem de "Nenhum resultado encontrado" aparece
                try:
                    nenhum_resultado = driver.find_element(By.CSS_SELECTOR, 'p.pull-left.ng-scope')
                    if nenhum_resultado.is_displayed() and "Nenhum resultado encontrado" in nenhum_resultado.text:
                        print(f'Nenhum resultado encontrado para o código: {codigo_completo}')
                        return None  # Ou pule para o próximo código conforme sua lógica
                except NoSuchElementException:
                    pass  # Caso o elemento não seja encontrado, continua normalmente
                
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'item'))
                )
                
                #time.sleep(5)
                
                produtos = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, 'item'))
                )
                    
                for produto in produtos:
                    try:
                        nome = WebDriverWait(produto, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.produto a'))
                        ).text.strip()
                        
                        fabricante = WebDriverWait(produto, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.fabricante b'))
                        ).text.strip()
                        
                        codigo_fabricante = WebDriverWait(produto, 10).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.fabricante b'))
                        )[1].text.strip()

                        codigo_fabricante_numerico = extrair_numero_codigo(codigo_fabricante)
                        
                        if codigo_fabricante == codigo_formatado:
                            print(f'Produto encontrado: {nome}, Código: {codigo_fabricante}, Código Planilha: {codigo_completo}')
                            
                            imagem_element = WebDriverWait(produto, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.img-produto'))
                            )
                            imagem_url = imagem_element.get_attribute('style')
                            imagem_url = imagem_url.split('url("')[1].split('")')[0] if 'url("' in imagem_url else ''
                            
                            preco_script = """
                            var preco = arguments[0].querySelector('.preco p');
                            return preco ? preco.textContent.trim() : '';
                            """
                            preco_produto = driver.execute_script(preco_script, produto)
                            
                            if not preco_produto:
                                try:
                                    preco_element = WebDriverWait(produto, 10).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, '.preco p.ng-binding'))
                                    )
                                    preco_produto = preco_element.text.strip()
                                except:
                                    preco_produto = "Preço não disponível"
                            
                            disponibilidade_script = """
                            var status = arguments[0].querySelector('.item-status');
                            return status ? status.textContent.trim() : 'Status não disponível';
                            """
                            disponibilidade = driver.execute_script(disponibilidade_script, produto)
                            
                            distribuidor_script = """
                            var distribuidor = arguments[0].querySelector('.nome-distribuidor p');
                            return distribuidor ? distribuidor.textContent.trim() : 'Distribuidor não disponível';
                            """
                            distribuidor = driver.execute_script(distribuidor_script, produto)
                            
                            return {
                                'Nome': nome,
                                'Fabricante': fabricante,
                                'Disponibilidade': disponibilidade,
                                'Código do Fabricante': codigo_fabricante,
                                'Código Original': codigo_completo,
                                'Preço': preco_produto,
                                'Distribuidor': distribuidor,
                                'Imagem URL': imagem_url
                            }
                    except (NoSuchElementException, StaleElementReferenceException) as e:
                        print(f'Erro ao extrair dados do produto: {e}')
                        continue
                
                print(f'Produto não encontrado para o código: {codigo_completo}')
                return None
            
            except TimeoutException:
                print(f'Tempo excedido ao carregar a página para o código {codigo_completo}. Tentativa {tentativa + 1} de {max_tentativas}')
                tentativa += 1
                time.sleep(2)
            except Exception as e:
                print(f'Erro ao buscar o código {codigo_completo}: {e}. Tentativa {tentativa + 1} de {max_tentativas}')
                tentativa += 1
                time.sleep(3)
        
        print(f'Falha ao buscar o código {codigo_completo} após {max_tentativas} tentativas')
        return None
    finally:
        release_driver(driver)

def processo_principal(combinacoes, caminho_arquivo):
    num_threads = psutil.cpu_count(logical=True)
    
    codigos_ja_capturados = [codigo.replace('-', '').replace('/', '') for codigo in obter_codigos_ja_capturados(caminho_arquivo)]
    combinacoes_faltantes = [combo for combo in combinacoes if combo.split()[-1] not in codigos_ja_capturados]
    
    print(f"Total de combinações: {len(combinacoes)}")
    print(f"Combinações já capturadas: {len(codigos_ja_capturados)}")
    print(f"Combinações faltantes: {len(combinacoes_faltantes)}")
    
    produtos_encontrados = []
    
    try:
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(buscar_codigo_no_site, codigo) for codigo in combinacoes_faltantes]
            for future in as_completed(futures):
                try:
                    resultado = future.result()
                    if resultado:
                        with lock:
                            produtos_encontrados.append(resultado)
                            if len(produtos_encontrados) % 10 == 0:
                                salvar_em_planilha(produtos_encontrados, caminho_arquivo)
                                produtos_encontrados = []
                except Exception as e:
                    print(f"Erro em uma das tarefas: {e}")
    except KeyboardInterrupt:
        print("Execução interrompida pelo usuário.")
    finally:
        print("Encerrando o processo de forma segura.")
        if produtos_encontrados:
            with lock:
                salvar_em_planilha(produtos_encontrados, caminho_arquivo)
        while not driver_pool.empty():
            driver = driver_pool.get()
            driver.quit()

if __name__ == "__main__":
    caminho_planilha = 'C:/Users/crrob/OneDrive/Área de Trabalho/Projeto PeçasCarros/planteste.xlsx'
    caminho_arquivo_saida = 'kdpeça-teste1.xlsx'
    combinacoes = ler_planilha(caminho_planilha)
    
    if combinacoes:
        initial_driver = create_driver()
        if realizar_login('SENHA', 'LOGIN', initial_driver):
            driver_pool.put(initial_driver)
            processo_principal(combinacoes, caminho_arquivo_saida)
        else:
            print("Não foi possível realizar o login. Encerrando o programa.")
            initial_driver.quit()
    else:
        print("Não foi possível ler a planilha. Encerrando o programa.")