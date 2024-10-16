from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import time
import os
import sys
import pandas as pd
import tempfile
import chromedriver_autoinstaller  # Biblioteca para baixar automaticamente o chromedriver

def setup_driver():

    chromedriver_autoinstaller.install()  # Esta linha baixará o chromedriver correto se ele não estiver presente


    return webdriver.Chrome()

#service = Service(executable_path=chromedriver_path)
#driver = webdriver.Chrome(service=service)

# Acesse o site
driver = setup_driver()
driver.get("https://bioid-shared.bionexo.com/users/sign_in")

MAX_TENTATIVAS = 3  # Número máximo de tentativas de login

# Contador para rastrear ocorrências consecutivas da condição
contagem_consecutiva = 0
LIMITE_CONSECUTIVO = 5

def verificar_login_bem_sucedido(driver):
    try:
        # Procure por um elemento que indique que o login foi bem-sucedido
        elemento_sucesso = driver.find_element(By.XPATH, '//label[@for="search_category_1_visible"]')
        return True
    except NoSuchElementException:
        return False
    
def realizar_login(driver):
    # Acesse o site

    # Aguarde até que a página esteja totalmente carregada
    time.sleep(3)

    # Selecione a opção "Bionexo" no dropdown
    select_element = Select(driver.find_element(By.ID, "user_login_application"))
    select_element.select_by_value("bio")

    # Aguarde um momento para o campo de email ficar habilitado
    time.sleep(1)

    # Preencha o campo de email
    email_input = driver.find_element(By.ID, "user_email")
    email_input.send_keys("alexsandro.santos@valetmed.com.br")

    # Clique no botão "Próximo"
    next_step_button = driver.find_element(By.ID, "nextStep")
    driver.execute_script("arguments[0].removeAttribute('disabled')", next_step_button)
    next_step_button.click()

    # Aguarde a transição para a próxima etapa
    time.sleep(3)

    # Preencha o campo de senha
    password_input = driver.find_element(By.ID, "user_password")
    password_input.send_keys("Valetmed2024@") 
    # Insira a senha correta

    # Clique no botão "Fazer Login"
    login_button = driver.find_element(By.ID, "sign_in")
    driver.execute_script("arguments[0].removeAttribute('disabled')", login_button)
    login_button.click()

    # Aguarde para ver o resultado do login

    time.sleep(10)
    if verificar_login_bem_sucedido(driver):
        return True
    else:
        return False


# Loop de login com múltiplas tentativas
for tentativa in range(1, MAX_TENTATIVAS + 1):
    print(f"Tentativa de login {tentativa}/{MAX_TENTATIVAS}")
    if realizar_login(driver):
        checkbox_label = driver.find_element(By.XPATH, '//label[@for="search_category_1_visible"]')
        checkbox_label.click()
        time.sleep(3)
        print("Login bem-sucedido!")
        break
    else:
        print("Login falhou. Tentando novamente...")
        time.sleep(5)  # Ajuste o tempo de espera entre as tentativas
        driver.refresh()

if tentativa == MAX_TENTATIVAS:
    print("Todas as tentativas de login falharam.")
    # Adicione aqui a lógica para enviar um email ou notificação, ou encerrar o script
    # Por exemplo:
    print("Login falhou após várias tentativas")
    driver.quit()


# Carregar dados existentes da planilha
file_path = "cotacoes_processadas.xlsx"
if os.path.exists(file_path):
    df_existing = pd.read_excel(file_path)
    df_existing = pd.DataFrame(df_existing)
    codigos_planilha = set(df_existing['Código'].tolist())
else:
    df_existing = pd.DataFrame()
    codigos_planilha = set()

# Lista para armazenar todas as cotações
cotacoes = []

# Conjunto para rastrear códigos de cotação já processados
codigos_processados = set()

# Pergunta ao usuário sobre o tipo de raspagem
def escolher_tipo_de_raspagem():
    while True:
        print("Escolha o tipo de raspagem:")
        print("1. Raspagem Completa")
        print("2. Raspagem Rápida")
        escolha = input("Digite 1 ou 2: ").strip()
        if escolha == '1':
            return 'completa'
        elif escolha == '2':
            return 'rapida'
        else:
            print("Escolha inválida. Por favor, digite 1 ou 2.")

tipo_de_raspagem = escolher_tipo_de_raspagem()

# Modifica a função `process_page` para incluir as condições adequadas
def process_page():
    global contagem_consecutiva
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    cotacoes_divs = soup.find_all(class_="list-group-item")

    if not cotacoes_divs:
        return False

    for cotacao_div in cotacoes_divs:
        try:
            tipo_cotacao = cotacao_div.find(class_="quote-kind-description").text.strip()
            codigo = cotacao_div.find(class_="quote-code").text.strip()
            comprador = cotacao_div.find(class_="quote-name").text.strip()
            estado = cotacao_div.find(class_="quote-state").text.strip()
            cidade = cotacao_div.find(class_="quote-city").text.strip()
            itens = cotacao_div.find(class_="quote-items").text.strip()
            abertura_data = cotacao_div.find(class_="quote-open-date").text.strip()
            abertura_hora = cotacao_div.find(class_="quote-open-time").text.strip()
            vencimento_data = cotacao_div.find(class_="quote-expiring-date").text.strip()
            vencimento_hora = cotacao_div.find(class_="quote-expiring-time").text.strip()

            # Escolhe a condição de acordo com o tipo de raspagem
            if tipo_de_raspagem == 'completa':
                if codigo in codigos_processados:
                    print(f"Código {codigo} já existe no sistema. Pulando...")
                    continue
            elif tipo_de_raspagem == 'rapida':
                if codigo in codigos_planilha:
                    contagem_consecutiva += 1
                    print(f"Código {codigo} já existe na planilha. Pulando...")
                    if contagem_consecutiva >= LIMITE_CONSECUTIVO:
                        print("Limite de códigos duplicados consecutivos atingido. Encerrando o programa...")
                        driver.quit()
                        exit()  # Encerra o programa
                    continue
                else:
                    contagem_consecutiva = 0

            codigos_processados.add(codigo)

            cotacao_link = cotacao_div.find('a', href=True)
            numero_cotacao = cotacao_link['href'].split("/")[-1]
            url_cotacao = f"https://bionexonew.bionexo.com/cotacoes/{numero_cotacao}"
            driver.get(url_cotacao)

            WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "quote-item-span")))

            produtos_divs = scroll_and_collect(driver)
            
            html = driver.page_source
            soup_cotacao = BeautifulSoup(html, 'html.parser')

            itens_solicitados = []
            produtos_solicitados = []
            marcas = []
            quantidades = []
            observacoes = []

            for produto_div in produtos_divs:
                try:
                    item_element = produto_div.find(class_='pull-left quote-item-sequence item-status-')
                    item = item_element.text.strip() if item_element else "N/A"

                    produto_solicitado_element = produto_div.find(class_='product-description')
                    produto_solicitado = produto_solicitado_element.text.strip() if produto_solicitado_element else "N/A"

                    marca = produto_div.find(class_='quote-item-value prefered-brands').get('title')

                    quantidade_element = produto_div.find(class_='quantity_description')
                    quantidade = quantidade_element.text.strip() if quantidade_element else "N/A"

                    itens_solicitados.append(item)
                    produtos_solicitados.append(produto_solicitado)
                    marcas.append(marca)
                    quantidades.append(quantidade)
                    
                except Exception as e:
                    print(f"Erro ao processar item: {e}")
                    continue
                
            observacoes_element = soup_cotacao.find(id="quote-observations")
            observacoes_text = observacoes_element.text.strip() if observacoes_element else "N/A"
            observacoes.append(observacoes_text)

            for item, produto_solicitado, marca, quantidade in zip(itens_solicitados, produtos_solicitados, marcas, quantidades):
                cotacoes.append({
                    "Código": codigo,
                    "Tipo de Cotação": tipo_cotacao,
                    "Comprador": comprador,
                    "Estado": estado,
                    "Cidade": cidade,
                    "Itens": itens,
                    "Data de Abertura": abertura_data,
                    "Hora de Abertura": abertura_hora,
                    "Data de Vencimento": vencimento_data,
                    "Hora de Vencimento": vencimento_hora,
                    "Itens Produtos": item,
                    "Produtos Solicitados": produto_solicitado,
                    "Marcas": marca,
                    "Quantidades": quantidade,
                    "Observações": observacoes
                })

            driver.back()
            WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "list-group-item")))

        except Exception as e:
            print(f"Erro ao processar cotação: {e}")
            continue
        df_novo = pd.DataFrame(cotacoes)

        salvar_cotacoes(df_existing, df_novo, file_path)

    return True

# Função para salvar cotações em Excel, protegendo contra corrupção
def salvar_cotacoes(df_existing, df_novo, file_path):
    try:
        if not df_existing.empty:
            df_total = pd.concat([df_existing, df_novo], ignore_index=True)
            codigos_planilha = set(df_existing['Código'].tolist())
        else:
            df_total = df_novo

        # Cria um arquivo temporário com a extensão correta
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            temp_file_path = tmp.name

        # Salva no arquivo temporário
        df_total.to_excel(temp_file_path, index=False)

        # Substitui o arquivo original pelo temporário
        os.replace(temp_file_path, file_path)
        
        print("Dados salvos com sucesso em 'cotacoes_processadas.xlsx'.")
    except Exception as e:
        print(f"Erro ao salvar cotações: {e}")

def scroll_and_collect(driver, max_elements=50):
    total_elements = 0
    produtos_divs = []

    while True:
        html = driver.page_source
        soup_cotacao = BeautifulSoup(html, 'html.parser')

        novos_produtos_divs = soup_cotacao.find_all(class_="quote-item-span")

        if len(novos_produtos_divs) == len(produtos_divs):
            break

        produtos_divs.extend(novos_produtos_divs[len(produtos_divs):])

        if len(produtos_divs) >= total_elements + max_elements:
            total_elements = len(produtos_divs)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(4)
            continue

        break

    return produtos_divs

def clicar_proxima_pagina(driver, pagina_atual):
    try:
        wait = WebDriverWait(driver, 10)
        elemento_pagina_atual = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "li.active[data-pagination-action]")))
        proxima_pagina = int(elemento_pagina_atual.get_attribute('data-pagination-action')) + 1

        xpath = f"//ul[@class='pagination']/li/a[@data-pagination-action='{proxima_pagina}']"
        proximo_link = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))

        driver.execute_script("arguments[0].click();", proximo_link)
        time.sleep(5)
        return True

    except NoSuchElementException:
        print("Botão da próxima página não encontrado.")
        return False
    except TimeoutException:
        print("Tempo esgotado ao tentar localizar o botão da próxima página.")
        return False

# Processamento de todas as páginas
pagina_atual = 1
while process_page():
    if not clicar_proxima_pagina(driver, pagina_atual):
        break
    pagina_atual += 1

# Salvar as cotações em um arquivo Excel


# Fechar o WebDriver ao final do processo
driver.quit()