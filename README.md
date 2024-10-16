# 🤖 Bots de Automação e Web Scraping

Coleção de scripts de automação e web scraping desenvolvidos para diferentes plataformas, utilizando tecnologias como Selenium, BeautifulSoup e Pandas.

## 📋 Índice

- [Bionexo](#bionexo)
- [MercadoCar](#mercadocar)
- [ContaPublicas](#contapublicas)
- [KdPeças](#kdpeças)
- [VendaMais](#vendamais)

## 🛠️ Tecnologias Utilizadas

- Python
- Selenium
- BeautifulSoup
- Pandas
- ThreadPoolExecutor
- FuzzyWuzzy

## 🤖 Bots Disponíveis

### Bionexo

Bot para automação de login e raspagem de cotações no site Bionexo.

#### Funcionalidades:
- ✅ Login automatizado com múltiplas tentativas
- 📊 Dois modos de raspagem: Completa e Rápida
- 💾 Armazenamento em Excel com backup automático
- 🔄 Paginação automática
- 🔍 Coleta detalhada de produtos e cotações

### MercadoCar

Scraper multithread para busca de peças automotivas.

#### Funcionalidades:
- 🚀 Processamento paralelo com ThreadPoolExecutor
- 🔄 Salvamento automático a cada 10 produtos
- 🛡️ Proteção contra bloqueios com fake_useragent
- 📊 Manipulação otimizada de planilhas Excel

### ContaPublicas

Automatização para coleta de processos de licitação no Portal de Compras Públicas.

#### Funcionalidades:
- 🤖 Modo headless para execução em segundo plano
- 🔄 Paginação automática
- 📊 Coleta via API
- 💾 Backup com timestamp

### KdPeças

Bot para busca automatizada de códigos de produtos.

#### Funcionalidades:
- 🔐 Login automático com CNPJ
- 🔄 Renovação automática de sessão
- 📊 Verificação de similaridade com FuzzyWuzzy
- 💾 Sistema de checkpoint para retomada

### VendaMais

Scraper otimizado para busca de autopeças.

#### Funcionalidades:
- 🔄 Gestão automática de sessões
- 🚀 Execução paralela de buscas
- 🛡️ Sistema anti-bloqueio
- 📊 Manipulação segura de dados

## 🚀 Como Usar

1. Clone o repositório
```bash
git clone https://github.com/CarlosRoGuerra/Robos_Dep

2. Instale as dependências
```bash
pip install -r requirements.txt
```

3. Configure as credenciais necessárias em um arquivo `.env`

4. Execute o bot desejado
```bash
python [nome_do_bot].py
```

## 📈 Melhorias Futuras

- [ ] Implementação de mais funcionalidades multithread
- [ ] Parametrização dinâmica via interface do usuário
- [ ] Melhorias no tratamento de erros
- [ ] Interface gráfica para configuração dos bots

## 🤝 Contribuições

Contribuições são sempre bem-vindas! Sinta-se à vontade para abrir uma issue ou criar um pull request.

## 📝 Licença

Este projeto está sob a licença Carlos Guerra.
