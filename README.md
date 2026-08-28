# RPA de ETL e Injeção de Dados BI - iGaming

Automação ponta a ponta (Web Scraping, Tratamento de Dados e Injeção em Excel) desenvolvida para orquestrar e consolidar relatórios diários de plataformas de iGaming. 

O projeto resolve o problema de gargalo operacional na extração manual de dezenas de relatórios financeiros diários, higienizando os dados estruturados e injetando-os nativamente em Dashboards de BI no Excel.

## Arquitetura do Projeto

O pipeline foi dividido em três microsserviços principais aplicando os princípios SOLID:

1. Web Scraper (Extração via Playwright)
   - Navegação automatizada no painel administrativo utilizando sessões persistentes (user_data_dir) para bypass amigável de restrições de sessão.
   - Extração de planilhas .xlsx com filtros cronológicos dinâmicos (UGS, FTD, Novas Contas, Transações).
   - Interceptação de tráfego de rede (Network Interception) para capturar o payload JSON direto da API interna, agilizando o módulo de General Statistics.

2. Data Cleaner (Higienização via Pandas e Calamine)
   - Performance: Uso do motor calamine (escrito em Rust) para leitura ultrarrápida de arquivos de Excel crus.
   - Continuidade Temporal: Algoritmo de validação (Merge 1:1) para preencher lacunas temporais, imputando dias faltantes com valores zerados para garantir a integridade dos cálculos matemáticos.
   - Blindagem de Infraestrutura: Tratamento isolado do openpyxl com estrutura try/finally para evitar Memory Leaks e corrupção de arquivos, garantindo a formatação precisa de máscaras financeiras (R$) e datas (dd/mm/yyyy).

3. Excel Injector (Injeção via Win32COM)
   - Manipulação nativa da API do Windows (pywin32) para abrir relatórios históricos e de performance em background.
   - Atualização em lote de Pivot Tables (RefreshAll) e injeção de dados formatados na última linha disponível.
   - Logs avançados e sistema de auditoria que compara valores retroativos para evitar injeção de dados duplicados ou perdas matemáticas.

## Como Executar

### Pré-requisitos
- Python 3.10+
- Sistema Operacional Windows (Obrigatório devido à dependência nativa do win32com.client para manipulação do Microsoft Excel).
- Instância do Microsoft Office instalada.

### Instalação

1. Clone o repositório:
    git clone https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
    cd SEU_REPOSITORIO

2. Crie o ambiente virtual e instale as dependências:
    python -m venv venv
    venv\Scripts\activate
    pip install -r requirements.txt

3. Instale os binários do Playwright:
    playwright install chromium

4. Configure as variáveis de ambiente baseando-se no arquivo de exemplo:
    copy .env.example .env

5. Inicie a orquestração do RPA:
    python main.py

## Segurança e Boas Práticas
- Credenciais isoladas e gerenciadas exclusivamente via variáveis de ambiente (.env).
- Sistema interativo de fallback: o bot aguarda intervenção humana em caso de desconexão inesperada que exija resolução manual de reCAPTCHA.
- Arquivos de dados sensíveis, relatórios financeiros e configurações locais estão estritamente protegidos via .gitignore para conformidade com confidencialidade de dados (NDA).