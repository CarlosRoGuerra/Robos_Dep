import pandas as pd
from bs4 import BeautifulSoup
import time
import random
import os
from fake_useragent import UserAgent
from concurrent.futures import ThreadPoolExecutor, as_completed
import psutil
from requests_html import HTMLSession
from requests.exceptions import RequestException, Timeout
import re

session = HTMLSession()
ua = UserAgent()

def ler_planilha(caminho_planilha):
    try:
        df = pd.read_excel(caminho_planilha)
        
        if 'FABRICANTE/MARCA' not in df.columns or 'CODIGO' not in df.columns:
            raise ValueError("As colunas 'FABRICANTE/MARCA' e 'CODIGO' precisam estar presentes na planilha.")
        
        combinacoes = (df['FABRICANTE/MARCA'].astype(str) + ' ' + df['CODIGO'].astype(str)).tolist()
        
        return combinacoes
    
    except ValueError as e:
        print(f"Erro: {e}")
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")

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

def get_random_headers():
    return {
        'User-Agent': ua.random,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Cache-Control': 'max-age=0',
    }

def obter_codigos_ja_capturados(caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        df = pd.read_excel(caminho_arquivo)
        return set(df['Código do Fabricante'].astype(str))
    return set()

def buscar_codigo_no_site(codigo):
    url = 'https://www.mercadocar.com.br/search'
    max_retries = 5
    retry_delay = 5
    
    for attempt in range(max_retries):
        try:
            headers = get_random_headers()
            params = {'q': codigo}
            
            response = session.get(url, params=params, headers=headers, timeout=15)
            
            if response.status_code == 403:
                print(f"Acesso negado ao buscar o código {codigo}. Erro 403. Tentativa {attempt + 1} de {max_retries}.")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay * (attempt + 1))
                    continue
                else:
                    print(f"Falha ao acessar após {max_retries} tentativas.")
                    return None
            
            soup = BeautifulSoup(response.html.html, 'html.parser')

            produtos = soup.find_all('div', class_='product-item')
            for produto in produtos:
                nome = produto.find('h2', class_='product-title').get_text(strip=True)
                preco = produto.find('span', class_='actual-price').get_text(strip=True)
                
                # Melhorar a comparação do código
                codigo_limpo = re.sub(r'[^a-zA-Z0-9]', '', codigo.lower())
                nome_limpo = re.sub(r'[^a-zA-Z0-9]', '', nome.lower())
                
                if codigo_limpo in nome_limpo:
                    print(f'Produto encontrado: {nome}, Preço: {preco}')               
                    try:
                        produto_nome = produto.find('div', class_='product-name').get_text(strip=True)
                        fabricante = produto.find('div', class_='manufacturers').find('span', class_='value').get_text(strip=True)
                        disponibilidade = produto.find('div', class_='stock').find('span', class_='value').get_text(strip=True)
                        codigo_fabricante = produto.find('div', class_='manufacturer-part-number').get_text(strip=True)
                        preco_produto = produto.find('div', class_='product-price').get_text(strip=True)
                        descricao = produto.find('div', class_='cnt-description').get_text(strip=True)
                        imagem_url = produto.find('div', class_='picture').find('a')['data-full-image-url']
                        
                        return {
                            'Nome': produto_nome,
                            'Fabricante': fabricante,
                            'Disponibilidade': disponibilidade,
                            'Código do Fabricante': codigo_fabricante,
                            'Preço': preco_produto,
                            'Descrição': descricao,
                            'Imagem URL': imagem_url
                        }
                        
                    except Exception as e:
                        print(f'Erro ao extrair dados da página do produto: {e}')
                    
                    break
            
            return None
        
        except (RequestException, Timeout) as e:
            print(f'Erro ao buscar o código {codigo}: {e}. Tentativa {attempt + 1} de {max_retries}.')
            if attempt < max_retries - 1:
                time.sleep(retry_delay * (attempt + 1))
            else:
                print(f"Falha após {max_retries} tentativas para o código {codigo}.")
                return None
            
def processo_principal(combinacoes, caminho_arquivo):
    num_threads = psutil.cpu_count(logical=True)
    
    codigos_ja_capturados = obter_codigos_ja_capturados(caminho_arquivo)
    combinacoes_faltantes = [combo for combo in combinacoes if re.sub(r'[^a-zA-Z0-9]', '', combo.split()[-1]) not in codigos_ja_capturados]
    
    print(f"Total de combinações: {len(combinacoes)}")
    print(f"Combinações já capturadas: {len(codigos_ja_capturados)}")
    print(f"Combinações faltantes: {len(combinacoes_faltantes)}")
    
    produtos_encontrados = []
    
    try:
        with ThreadPoolExecutor(max_workers=num_threads) as executor:
            futures = [executor.submit(buscar_codigo_no_site, codigo) for codigo in combinacoes_faltantes]
            for future in as_completed(futures):
                try:
                    resultado = future.result()
                    if resultado:
                        produtos_encontrados.append(resultado)
                        if len(produtos_encontrados) % 10 == 0:
                            salvar_em_planilha(produtos_encontrados, caminho_arquivo)
                            produtos_encontrados = []  # Limpa a lista após salvar
                except Exception as e:
                    print(f"Erro em uma das tarefas: {e}")
    except KeyboardInterrupt:
        print("Execução interrompida pelo usuário.")
    finally:
        print("Encerrando o processo de forma segura.")
        if produtos_encontrados:
            salvar_em_planilha(produtos_encontrados, caminho_arquivo)

if __name__ == "__main__":
    caminho_planilha = 'C:/Users/crrob/OneDrive/Área de Trabalho/Projeto PeçasCarros/planteste.xlsx'
    caminho_arquivo_saida = 'mercadocar-producao.xlsx'
    combinacoes = ler_planilha(caminho_planilha)
    processo_principal(combinacoes, caminho_arquivo_saida)