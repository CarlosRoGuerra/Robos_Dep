import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from fuzzywuzzy import fuzz
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import random
from selenium.webdriver.common.alert import Alert

def create_driver():
    chrome_options = webdriver.ChromeOptions()
    temp_directory = "G:\\selenium_temp_files2"
    #Configurar o diretório temporário para o Chrome
    chrome_options.add_argument(f'--user-data-dir={temp_directory}')  # Usado para armazenamento de dados do perfil
    chrome_options.add_argument(f'--disk-cache-dir={temp_directory}')  # Usado para cache do navegador
    chrome_options.add_argument('--headless')  # Opcional: rodar sem abrir o navegador
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-gpu')
    return webdriver.Chrome(options=chrome_options)

def realizar_login(cnpj, senha, driver):
    login_url = 'http://www.vendamais.lucios.com.br/Login.aspx'
    max_tentativas = 3
    for tentativa in range(max_tentativas):
        try:
            driver.get(login_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'txtUsuario'))).send_keys(cnpj)
            driver.find_element(By.NAME, 'txtSenha').send_keys(senha)
            driver.find_element(By.ID, 'btnLogin').click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//td[contains(text(),"Usuário:")]')))
            print("Login realizado com sucesso!")
            return True
        except Exception as e:
            print(f"Tentativa {tentativa + 1} falhou: {e}")
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

def buscar_codigo_no_site(driver, search_url,codigo_completo):
    codigo_formatado = codigo_completo.replace('-', '').replace('/', '').replace('.', '').strip()
    partes = codigo_formatado.split()
    fabricante = partes[0] if len(partes) > 0 else ""
    codigo_puro = partes[1] if len(partes) > 1 else ""
    codigo_limpo = codigo_puro.strip()
    driver.get(search_url)

    try:
        time.sleep(random.uniform(2, 4))  # Random delay to mimic human behavior
        if verificar_se_login_expirou(driver):
            print("Login expirado, realizando novo login...")
            if not realizar_login('67671537000154', 'Cnpj67671@', driver):
                print("Erro ao realizar novo login.")
                return None      
        # Expandir o elemento escondido via JavaScript
        try:
            referencia = driver.find_element("xpath", '//div[contains(@style, "height: 0px; display: none;")]')
            driver.execute_script("arguments[0].style.height='auto'; arguments[0].style.display='block';", referencia)
            print("Elemento exibido com sucesso!")
        except NoSuchElementException:
            print("Erro: Referência não encontrada.")
        except Exception as e:
            print(f"Erro inesperado: {e}")

        # Limpar o campo de busca e enviar o novo valor
        campo_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_cphMaster_txtProduto_MaskTextBox'))
        )
        campo_input.clear()
        campo_input.send_keys(codigo_limpo)
        botao_busca = driver.find_element(By.ID, 'ctl00_cphMaster_btnBuscarProduto')
        driver.execute_script("arguments[0].click();", botao_busca)
        # Verificar se não foram encontrados resultados
        try:
            time.sleep(random.uniform(2, 3))
            nenhum_resultado = driver.find_element(By.ID, 'ctl00_cphMaster_lblResultado')
            if nenhum_resultado.is_displayed() and "Nenhum registro foi encontrado." in nenhum_resultado.text:
                print(f'Nenhum registro foi encontrado para o código {codigo_completo}.')
                return None
        except NoSuchElementException:
            pass  # Caso o elemento não seja encontrado, continua normalmente

        # Processar a tabela de resultados
        tabela = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ctl00_cphMaster_gvBusca')))
        linhas = tabela.find_elements(By.XPATH, ".//tr[not(contains(@class,'gridHeader'))]")
        
        for linha in linhas:
            codigo_fabr = linha.find_elements(By.XPATH, ".//td")[1].text.strip()
            fornecedor = linha.find_elements(By.XPATH, ".//td")[3].text.strip()
            similarity_score = fuzz.ratio(fornecedor, fabricante)
            if codigo_fabr.replace('-', '').replace('/', '').replace('.', '') == codigo_limpo: #and similarity_score >= 75:
                # Exibe os dados da linha correspondente
                print(f"Código: {codigo_limpo}, Fornecedor: {fornecedor}")
                return {
                    'Nome': linha.find_elements(By.XPATH, ".//td")[2].text.strip(),
                    'Fornecedor': fornecedor,
                    'Código do Fabricante': codigo_fabr,
                    'Código Original': codigo_completo,
                    'Preço': linha.find_elements(By.XPATH, ".//td")[4].text.strip(),
                    'Ano': linha.find_elements(By.XPATH, ".//td")[6].text.strip(),
                    'Informações Complementares': linha.find_elements(By.XPATH, ".//td")[7].text.strip(),
                    'Modelo': linha.find_elements(By.XPATH, ".//td")[5].text.strip(),
                    'Código': linha.find_elements(By.XPATH, ".//td")[0].text.strip(),
                }

        print(f'Código não encontrado na tabela: {codigo_completo}')
        return None

    except Exception as e:
        print(f'Erro ao buscar o código {codigo_completo}: {e}')
        return None

def salvar_progresso(codigo_completo, arquivo_progresso='progresso2.txt'):
    with open(arquivo_progresso, 'w') as f:
        f.write(codigo_completo)

def carregar_progresso(arquivo_progresso='progresso2.txt'):
    if os.path.exists(arquivo_progresso):
        with open(arquivo_progresso, 'r') as f:
            return f.read().strip()
    return None

def processo_principal(combinacoes, caminho_arquivo):
    driver = create_driver()
    if not realizar_login('SUA_SENHA', 'SEU_LOGIN', driver):
        print("Não foi possível realizar o login. Encerrando o programa.")
        driver.quit()
        return

    ultimo_codigo_processado = carregar_progresso()

    # Se o progresso existe, começamos do último código processado
    if ultimo_codigo_processado:
        index_inicio = combinacoes.index(ultimo_codigo_processado) + 1
        combinacoes = combinacoes[index_inicio:]
    else:
        index_inicio = 0
    time.sleep(random.uniform(2, 3))
    search_url = f'http://www.vendamais.lucios.com.br/Cliente/Pedidos.aspx?a=S4sE6i%2fmfok%3d'
    driver.get(search_url)
    produtos_encontrados = []
    codigos_ja_capturados = set()

    if os.path.exists(caminho_arquivo):
        df_existente = pd.read_excel(caminho_arquivo, engine='openpyxl')
        codigos_ja_capturados = set(df_existente['Código Original'].astype(str))

    try:
        for codigo in combinacoes:
            if codigo not in codigos_ja_capturados:
                resultado = buscar_codigo_no_site(driver, search_url,codigo)
                if resultado:
                    produtos_encontrados.append(resultado)
                    salvar_progresso(codigo)  # Salva o código após processá-lo com sucesso
                    if len(produtos_encontrados) % 10 == 0:
                        salvar_em_planilha(produtos_encontrados, caminho_arquivo)
                        produtos_encontrados = []
                time.sleep(random.uniform(1, 3))  # Random delay between requests
    except KeyboardInterrupt:
        print("Execução interrompida pelo usuário.")
    finally:
        if produtos_encontrados:
            salvar_em_planilha(produtos_encontrados, caminho_arquivo)
        driver.quit()

def salvar_em_planilha(dados, caminho_arquivo):
    df_novos_dados = pd.DataFrame(dados)
    if os.path.exists(caminho_arquivo):
        df_existente = pd.read_excel(caminho_arquivo)
        df_combined = pd.concat([df_existente, df_novos_dados]).drop_duplicates(subset='Código Original').reset_index(drop=True)
    else:
        df_combined = df_novos_dados
    
    df_combined.to_excel(caminho_arquivo, index=False, engine='openpyxl')
    print(f'Dados salvos com sucesso em {caminho_arquivo}')

if __name__ == "__main__":
    caminho_planilha = 'C:/Users/crrob/OneDrive/Área de Trabalho/Projeto PeçasCarros/THOMSON.xlsx'
    caminho_arquivo_saida = 'G:\\planilha vendamais\\vendamais-teste2.xlsx'
    df = pd.read_excel(caminho_planilha)
    combinacoes = (df['FABRICANTE/MARCA'].astype(str) + ' ' + df['CODIGO'].astype(str)).loc[
    df['FABRICANTE/MARCA'].notna() & (df['FABRICANTE/MARCA'] != '') &
    df['CODIGO'].notna() & (df['CODIGO'] != '')
  ].tolist()
    
    if combinacoes:
        processo_principal(combinacoes, caminho_arquivo_saida)
    else:
        print("Não foi possível ler a planilha. Encerrando o programa.")