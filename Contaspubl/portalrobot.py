import time
import pandas as pd
import requests
from datetime import datetime
import re
import os
import glob
import urllib.parse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.common.exceptions import WebDriverException
import chromedriver_autoinstaller  # Biblioteca para baixar automaticamente o chromedriver


def setup_driver():
    # Verifica e instala a versão correta do chromedriver, se necessário
    chromedriver_autoinstaller.install()  # Esta linha baixará o chromedriver correto se ele não estiver presente
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Executa o navegador em modo headless (sem interface gráfica)
    options.add_argument('--disable-gpu')  # Necessário para alguns sistemas operacionais
    options.add_argument('--window-size=1920,1080')  # Define o tamanho da janela
    options.add_argument('--no-sandbox')  # Evita erros de sandbox, útil para ambientes Linux
    options.add_argument('--disable-dev-shm-usage')  # Necessário para am
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-background-networking')
    options.add_argument('--disable-sync')

    # Inicia o driver com o chromedriver baixado pela biblioteca
    return webdriver.Chrome(options=options)

def wait_for_element(driver, selector, by=By.CSS_SELECTOR, timeout=10):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )

def safe_get_text(driver, selector, by=By.CSS_SELECTOR):
    try:
        return driver.find_element(by, selector).text
    except NoSuchElementException:
        return None

def extract_dates(driver):
    date_script = """
    const dateElements = document.querySelectorAll('.half-side p.subtitle b');
    const dates = {};
    dateElements.forEach(el => {
        const key = el.textContent.trim().replace(':', '');
        const value = el.nextSibling.textContent.trim();
        if (['Data de Publicação', 'Início das Propostas', 'Limite para Impugnações'].includes(key)) {
            dates[key] = value;
        }
    });
    return dates;
    """
    return driver.execute_script(date_script)

def get_process_number(driver):
    title_element = driver.find_element(By.TAG_NAME, "title")
    title_text = title_element.get_attribute("innerHTML")
    match = re.search(r'-(\d+)$', title_text)
    if match:
        return match.group(1)
    return None

def fetch_items_from_api(process_number):
    all_items = []
    page = 1
    while True:
        url = f"https://compras.api.portaldecompraspublicas.com.br/v2/licitacao/{process_number}/itens?filtro=&pagina={page}"
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            items = None
            page_count = None
            # Tenta pegar de 'itens' primeiro
            if 'itens' in data and data['itens'] is not None:
                items = data['itens'].get('result')
                page_count = data['itens'].get('pageCount')

            # Se 'itens' não existir ou for None, tenta pegar de 'lotes'
            if items is None:  # Apenas tenta 'lotes' se 'itens' não retornar dados
                if 'lotes' in data and data['lotes'] is not None:
                    items = data['lotes'].get('result')
                    page_count = data['lotes'].get('pageCount')

            # Se encontrar algum item, processa
            if items:
                all_items.extend(items)
                # Se não encontrar nenhum item, interrompe o loop
        
                # Checa o pageCount para determinar se deve continuar
                if page_count is not None and page >= page_count:
                    break
                page += 1
            else:
                print("Nenhum item encontrado em 'itens' ou 'lotes'.")
        else:
            print(f"Failed to fetch data for process {process_number}, page {page}")
            break
    return all_items

def scrape_process_data(driver, url):
    driver.get(url)
    time.sleep(5)  # Allow time for JavaScript to load content

    try:
        licitação_numero = get_process_number(driver)
        process_number = safe_get_text(driver, "div.col-md-4.col-sm-12.p-0 span b")
        title = safe_get_text(driver, "h1.title.centro")
        
        dates = extract_dates(driver)
        
        if licitação_numero:
            items = fetch_items_from_api(licitação_numero)
        else:
            print(f"Could not extract process number for URL: {url}")
            items = []

        return {
            'Process Number': process_number,
            'Licitação Number': licitação_numero,
            'Title': title,
            'Publication Date': dates.get('Data de Publicação', ''),
            'Proposal Start': dates.get('Início das Propostas', ''),
            'Impugnation Deadline': dates.get('Limite para Impugnações', ''),
            'Items': items
        }
    except Exception as e:
        print(f"Error processing {url}: {e}")
        return None

def accept_cookies(driver):
    while True:
        try:
            cookies_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
            )
            cookies_button.click()
            print("Cookies aceitos.")
            break
        except TimeoutException:
            print("Botão de cookies não encontrado, verificando novamente.")
            time.sleep(2)  # Espera antes de tentar novamente


def get_latest_file(base_filename):
    files = glob.glob(f"{base_filename}_*.xlsx")
    if not files:
        return None
    return max(files, key=os.path.getctime)

def load_existing_data(filename):
    if filename and os.path.exists(filename):
        df = pd.read_excel(filename)
        return set(df['Numero do Processo'].unique())
    return set()

def save_data(df, base_filename):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{base_filename}_{timestamp}.xlsx"
    df.to_excel(filename, index=False)
    print(f"Data saved to {filename}")
    return filename

def main(filtro='materiais hospitalares'):
    driver = setup_driver()
    encoded_filtro = urllib.parse.quote(filtro)
    base_url = f'https://www.portaldecompraspublicas.com.br/processos?objeto={encoded_filtro}&pagina={{}}'
    all_process_links = []
    data = []
    last_save_time = time.time()
    save_interval = 300  # Save every 5 minutes

    # Use the filter in the base filename
    base_filename = f'processos_{filtro.replace(" ", "_").lower()}'

    # Load existing data
    latest_file = get_latest_file(f'{base_filename}_final')
    existing_processes = load_existing_data(latest_file)
    print(f"Loaded {len(existing_processes)} existing process numbers.")

    try:
        # Loop through the first 5 pages
        for page in range(1, 11):
            url = base_url.format(page)
            driver.get(url)
            time.sleep(5)
            
            if page == 1:
                accept_cookies(driver)

            process_links = list({a.get_attribute('href') for a in driver.find_elements(By.CSS_SELECTOR, "div#pesquisa-processo-page > div > main > section.container-fluid.resultado-pesquisa > div > section > div > div a") if a.get_attribute('href')})
            all_process_links.extend(process_links)

        for index, link in enumerate(all_process_links[:100], 1):
            try:
                process_data = scrape_process_data(driver, link)
                if process_data:
                    process_number = process_data['Process Number']
                    if process_number not in existing_processes:
                        data.append(process_data)
                        existing_processes.add(process_number)
                        print(f"Added new process: {process_number}")
                    else:
                        print(f"Skipped existing process: {process_number}")
                
                # Save periodically
                if time.time() - last_save_time > save_interval:
                    df = pd.DataFrame(prepare_dataframe(data))
                    save_data(df, f'{base_filename}_temp')
                    last_save_time = time.time()
                
                print(f"Processed {index} out of {min(100, len(all_process_links))} links")
            
            except WebDriverException as e:
                print(f"Error processing link {link}: {str(e)}")
                continue  # Skip to the next link
    
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")
    
    finally:
        driver.quit()
        
        if data:
            df = pd.DataFrame(prepare_dataframe(data))
            if latest_file:
                existing_df = pd.read_excel(latest_file)
                df = pd.concat([existing_df, df]).drop_duplicates(subset='Numero do Processo', keep='last')
            
            final_filename = save_data(df, f'{base_filename}_final')
            print(f"Data successfully saved to {final_filename}")
        else:
            print("No new data was collected.")

def prepare_dataframe(data):
    df_data = []
    for process in data:
        for item in process['Items']:
            if 'itens' in item and isinstance(item['itens'], list):
                for subitem in item['itens']:
                    df_data.append(prepare_row(process, subitem))
            else:
                df_data.append(prepare_row(process, item))
    return df_data

def prepare_row(process, item):
    return {
        'Numero do Processo': process['Process Number'],
        'Numero de Licitação': process['Licitação Number'],
        'Titulo': process['Title'],
        'Data de Publicação': process['Publication Date'],
        'Inicio das Propostas': process['Proposal Start'],
        'Data Limite para Impugnações': process['Impugnation Deadline'],
        'Item': item.get('codigo', None),
        'Descrição': item.get('descricao', None),
        'Quantidade': item.get('quantidade', None),
        'Unidade': item.get('unidade', None),
        'Melhor Lance': item.get('melhorLance', None),
        'Valor Referência': item.get('valorReferencia', None),
        'Situação': item.get('situacao', {}).get('descricao', None)
    }

if __name__ == "__main__":
    filtro = input("Digite o filtro de busca (ex: materiais hospitalares): ")
    main(filtro)