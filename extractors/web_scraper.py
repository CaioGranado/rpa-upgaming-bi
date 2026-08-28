import json
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from config.settings import (
    DIVISOR_LOG,
    MARCAS_CONFIG,
    SEL_BRAND_DROPDOWN,
    SEL_BRAND_OK_BTN,
    SEL_CONFIRM_EXPORT,
    SEL_CREATE_DATE_FROM,
    SEL_CREATE_DATE_TO,
    SEL_EXPORT_BTN,
    SEL_FROM_DATE,
    SEL_GAME_TYPE_DROPDOWN,
    SEL_MENU_FTD,
    SEL_MENU_GEN_STATS,
    SEL_MENU_REPORT,
    SEL_MENU_TRANSACTIONS,
    SEL_MENU_UGS,
    SEL_MENU_USERS,
    SEL_MORE_FILTER_BTN,
    SEL_OK_BTN,
    SEL_SEARCH_ADD_BTN,
    SEL_SEARCH_BRAND_INPUT,
    SEL_SEARCH_BTN_ALT,
    SEL_SEARCH_FTD_BTN,
    SEL_TO_DATE,
    URL_SISTEMA,
)
from utils.date_utils import obter_data_alvo, obter_periodo_extracao
from utils.file_utils import obter_pasta_download_diario, obter_pasta_ugs_diario

logger = logging.getLogger(__name__)

def extrair_dados_upgaming():
    logger.info("Iniciando módulo de Extração Web...")
    
    # 1. PEGA AS DATAS INTELIGENTES DO NOSSO MÓDULO (Sem variáveis ociosas)
    data_inicio, data_fim, data_fim_nc = obter_periodo_extracao()
    arquivos_baixados = [] 

    try:
        with sync_playwright() as p:
            pasta_perfil = str(Path.cwd() / "perfil_robo_chrome")
            context = p.chromium.launch_persistent_context(
                user_data_dir=pasta_perfil,
                headless=False,
                channel="chrome", 
                chromium_sandbox=True, 
                ignore_default_args=["--no-sandbox", "--enable-automation"],
                args=['--disable-blink-features=AutomationControlled'],
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                viewport={'width': 1280, 'height': 720}
            )
            
            page = context.pages[0]
            page.set_default_timeout(300000)

            page.goto(URL_SISTEMA)
            
            # Espera o redirecionamento acontecer (caso o cookie seja inválido)
            try:
                page.wait_for_load_state("domcontentloaded", timeout=10000)
                # Uma pequena pausa extra garante que a URL mude completamente
                page.wait_for_timeout(1500)
            except PlaywrightTimeoutError:
                logger.debug("O carregamento inicial demorou mais que 10s. Seguindo para a avalição visual...")
            
            # 1. VERIFICAÇÃO DA TELA DE LOGIN (URL + Visual da imagem corrigida)
            is_login_url = "login.html" in page.url
            # Atualizado: usa o name="username" conforme o DevTools do seu print
            is_login_visual = page.locator('text="Welcome To Admin Panel"').is_visible() or page.locator('input[name="username"]').is_visible()

            # Se a URL acusou login OU a tela inicial apareceu, pede intervenção
            if is_login_url or is_login_visual:
                logger.warning("Página de login detectada (Sessão expirada).")
                input("\n>>> Faça o login e resolva os reCAPTCHA, espere o painel inicial carregar e então pressione ENTER aqui...\n")
                
                # 2. VALIDAÇÃO PÓS-LOGIN (Garante que a barra lateral apareceu)
                try:
                    page.wait_for_selector(SEL_MENU_REPORT, timeout=15000)
                    logger.info("Login confirmado com sucesso!")
                except PlaywrightTimeoutError:
                    logger.error("Falha ao confirmar o login: Menu lateral não encontrado, abortando por segurança.")
                    return arquivos_baixados
            else:
                # 3. PROVA REAL DO COOKIE (Garante que não é uma tela de erro 502/Cloudflare)
                logger.info("Avaliando sessão salva no cookie...")
                try:
                    page.wait_for_selector(SEL_MENU_REPORT, timeout=15000)
                    logger.info("Sessão ativa confirmada! Menu carregado, pulando login manual...")
                except PlaywrightTimeoutError:
                    logger.error("Estado desconhecido! Não é a tela de login, mas o menu não carregou. Possível erro de rede ou bloqueio.")
                    return arquivos_baixados
            
            # --- LOOP DE MARCAS ---
            for marca_arquivo, marca_bo in MARCAS_CONFIG.items():
                _extrair_relatorios_marca(
                    page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc, arquivos_baixados
                )

            return arquivos_baixados

    except Exception:
        logger.exception("FALHA CRÍTICA NA EXTRAÇÃO:")
        return arquivos_baixados


def _extrair_relatorios_marca(page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc, arquivos_baixados):
    """
    Função auxiliar criada para reduzir a 'Complexidade Cognitiva' do código.
    Ela processa os relatórios individualmente para a marca passada.
    """
    logger.info(DIVISOR_LOG)
    logger.info(f" EXTRAINDO MARCA: {marca_arquivo}")
    
    # 2. PEGA A PASTA DE DOWNLOAD CORRETA LÁ NO DRIVE G:
    pasta_destino = obter_pasta_download_diario(marca_arquivo)
    
    # [1/6] NOVAS CONTAS
    page.click(SEL_MENU_USERS)
    page.wait_for_timeout(2000)
    try:
        page.wait_for_selector(SEL_BRAND_DROPDOWN, timeout=4000)
        page.select_option(SEL_BRAND_DROPDOWN, label=marca_bo)
        page.click(SEL_BRAND_OK_BTN)
    except PlaywrightTimeoutError:
        page.click(SEL_SEARCH_BRAND_INPUT)
        page.fill(SEL_SEARCH_BRAND_INPUT, marca_bo)
        page.wait_for_timeout(1000)
        page.click(f'text="{marca_bo}" >> visible=true')
        page.wait_for_timeout(500)
    
    page.click(SEL_MORE_FILTER_BTN)
    page.fill(SEL_CREATE_DATE_FROM, data_inicio)
    page.fill(SEL_CREATE_DATE_TO, data_fim_nc)
    page.click(SEL_SEARCH_BTN_ALT)
    page.wait_for_timeout(6000) 
    
    with page.expect_download(timeout=120000) as download_info:
        page.click(SEL_EXPORT_BTN)
    
    arq_nc = str(pasta_destino / f"NC - {marca_arquivo}.xlsx")
    download_info.value.save_as(arq_nc)
    arquivos_baixados.append(arq_nc)
    logger.info(f"Salvo: {arq_nc}")

    # [2/6] SYSTEM TRANSACTIONS
    page.click(SEL_MENU_TRANSACTIONS)
    page.wait_for_timeout(3000)
    page.click('div.choosen:visible')
    page.wait_for_timeout(1000)
    
    page.click('text="All Brands" >> visible=true')
    page.wait_for_timeout(500)
    page.click('text="All Brands" >> visible=true')
    page.wait_for_timeout(500)
    page.click(f'text="{marca_bo}" >> visible=true')
    page.wait_for_timeout(500)
    page.click('div.choosen:visible')
    page.wait_for_timeout(500)
    
    page.fill(SEL_FROM_DATE, data_inicio)
    page.fill(SEL_TO_DATE, data_fim)
    page.click(SEL_OK_BTN)
    
    try:
        page.click(SEL_SEARCH_ADD_BTN, timeout=3000)
        logger.debug("Botão SEARCH clicado manualmente.")
    except PlaywrightTimeoutError:
        logger.debug("Botão SEARCH ignorado (busca automática acionada ou botão ausente).")
    page.wait_for_timeout(6000) 
    
    page.click(SEL_EXPORT_BTN)
    page.wait_for_selector(SEL_CONFIRM_EXPORT, timeout=10000)
    
    with page.expect_download(timeout=120000) as download_info:
        page.click(SEL_CONFIRM_EXPORT)
    
    arq_trans = str(pasta_destino / f"Transações - {marca_arquivo}.xlsx")
    download_info.value.save_as(arq_trans)
    arquivos_baixados.append(arq_trans)
    logger.info(f"Salvo: {arq_trans}")

    # [3/6] UGS ACUMULADO
    page.click(SEL_MENU_UGS)
    page.wait_for_timeout(3000)
    page.click(SEL_SEARCH_BRAND_INPUT)
    page.fill(SEL_SEARCH_BRAND_INPUT, marca_bo)
    page.wait_for_timeout(1000)
    page.click(f'text="{marca_bo}" >> visible=true')
    page.wait_for_timeout(500)
    
    page.fill(SEL_FROM_DATE, data_inicio)
    page.fill(SEL_TO_DATE, data_fim)
    page.click(SEL_OK_BTN)
    
    tipos_ugs = {"": "Completo", "1": "ST", "2": "LC", "7": "SB", "8": "MG"}
    for valor, sigla in tipos_ugs.items():
        page.select_option(SEL_GAME_TYPE_DROPDOWN, value=valor)
        page.click(SEL_SEARCH_ADD_BTN)
        page.wait_for_timeout(6000) 
        with page.expect_download(timeout=120000) as download_info:
            page.click(SEL_EXPORT_BTN)
        arq_ugs = str(pasta_destino / f"{marca_arquivo} - UGS {sigla}.xlsx")
        download_info.value.save_as(arq_ugs)
        arquivos_baixados.append(arq_ugs)
        logger.info(f"Salvo: {arq_ugs}")

    # [4/6] UGS DIÁRIO (Buscador Dinâmico de Lacunas)
    page.select_option(SEL_GAME_TYPE_DROPDOWN, value="")
    
    # Variável ajustável: Quantos dias no passado o robô deve checar?
    JANELA_DIAS = 7
    dias_diarios_faltantes = []
    
    hoje_real = datetime.now(timezone.utc).astimezone()
    
    logger.info(f"Checando lacunas de UGS Diário nos últimos {JANELA_DIAS} dias...")
    
    # Loop de trás para frente (ex: dia -7 até dia -1) para manter a ordem cronológica
    for i in range(JANELA_DIAS, 0, -1):
        dia_checar = hoje_real - timedelta(days=i)
        pasta_ugs_checar = obter_pasta_ugs_diario(marca_arquivo, dia_checar.year, dia_checar.month)
        nome_dia_checar = dia_checar.strftime("%d-%m")
        
        arquivo_esperado = pasta_ugs_checar / f"{nome_dia_checar}.xlsx"
        
        # Se o arquivo não existe fisicamente na pasta, entra na lista de download
        if not arquivo_esperado.exists():
            dias_diarios_faltantes.append(dia_checar)
            
    if not dias_diarios_faltantes:
        logger.info(f" Nenhuma lacuna encontrada! Todos os UGS dos últimos {JANELA_DIAS} dias já estão na pasta.")
    else:
        logger.info(f" Foram encontradas {len(dias_diarios_faltantes)} lacunas. Iniciando download...")

    # Agora o Playwright só entra em ação para os dias que realmente faltam
    for dia_alvo in dias_diarios_faltantes:
        d_inicio = dia_alvo.strftime("%d-%m-%Y 00:00")
        d_fim = dia_alvo.strftime("%d-%m-%Y 23:59")
        nome_dia = dia_alvo.strftime("%d-%m")
        
        page.fill(SEL_FROM_DATE, d_inicio)
        page.fill(SEL_TO_DATE, d_fim)
        page.click(SEL_OK_BTN)
        page.click(SEL_SEARCH_ADD_BTN)
        page.wait_for_timeout(6000)
        
        with page.expect_download(timeout=120000) as download_info:
            page.click(SEL_EXPORT_BTN)
        
        pasta_ugs_alvo = obter_pasta_ugs_diario(marca_arquivo, dia_alvo.year, dia_alvo.month)
        arq_ugs_diario = str(pasta_ugs_alvo / f"{nome_dia}.xlsx")
        
        download_info.value.save_as(arq_ugs_diario)
        arquivos_baixados.append(arq_ugs_diario)
        logger.info(f"Salvo UGS Diário: {arq_ugs_diario}")

    # [5/6] FTD
    page.click(SEL_MENU_FTD)
    page.wait_for_timeout(3000)
    page.click(SEL_SEARCH_BRAND_INPUT)
    page.fill(SEL_SEARCH_BRAND_INPUT, marca_bo)
    page.wait_for_timeout(1000)
    page.click(f'text="{marca_bo}" >> visible=true')
    page.wait_for_timeout(500)
    page.fill(SEL_FROM_DATE, data_inicio)
    page.fill(SEL_TO_DATE, data_fim)
    page.click(SEL_OK_BTN)
    page.click(SEL_SEARCH_FTD_BTN)
    page.wait_for_timeout(6000) 
    with page.expect_download(timeout=120000) as download_info:
        page.click(SEL_EXPORT_BTN)
    
    arq_ftd = str(pasta_destino / f"FTD - {marca_arquivo}.xlsx")
    download_info.value.save_as(arq_ftd)
    arquivos_baixados.append(arq_ftd)
    logger.info(f"Salvo: {arq_ftd}")
    
    # =====================================================================
    # [6/6] GENERAL STATISTICS (Scraping da API Invisível)
    # =====================================================================
    logger.info("Extraindo General Statistics via API (JSON)...")
    
    page.click(SEL_MENU_REPORT)
    page.wait_for_timeout(500)
    page.click(SEL_MENU_GEN_STATS)
    page.wait_for_timeout(3000)
    
    # Seleciona a marca
    page.click(SEL_SEARCH_BRAND_INPUT)
    page.fill(SEL_SEARCH_BRAND_INPUT, marca_bo)
    page.wait_for_timeout(1000)
    page.click(f'text="{marca_bo}" >> visible=true')
    page.wait_for_timeout(500)

    # Define a data_alvo buscando a inteligência do date_utils
    data_alvo = obter_data_alvo()

    # Extrai estritamente com base na inteligência da data_alvo
    dias_fechados = data_alvo.day 
    mes_alvo = data_alvo.month
    ano_alvo = data_alvo.year

    dados_json_mensal = []

    if dias_fechados > 0:
        for dia in range(1, dias_fechados + 1):
            # Substitui o 'hoje.replace' por uma formatação de data cravada
            data_loop = f"{dia:02d}-{mes_alvo:02d}-{ano_alvo}"
            logger.info(f" -> Extraindo dados do dia {data_loop}...")
            
            # =========================================================
            # AQUI ESTÁ O SEGREDO: Se um dia falhar, não aborta tudo!
            # =========================================================
            try:
                # TRUQUE ANTI-JS: Clicar, limpar e digitar pausadamente (SEU CÓDIGO ORIGINAL)
                loc_from = page.locator(SEL_FROM_DATE)
                loc_from.click()
                loc_from.clear()
                loc_from.press_sequentially(f"{data_loop} 00:00", delay=50)
                
                loc_to = page.locator(SEL_TO_DATE)
                loc_to.click()
                loc_to.clear()
                loc_to.press_sequentially(f"{data_loop} 23:59", delay=50)
                
                page.click(SEL_OK_BTN)
                
                # OBRIGATÓRIO: Dar 1 segundo para o site "entender" a data antes do Search
                page.wait_for_timeout(1000)
                
                # Escuta a aba "Network" e intercepta a requisição assim que clicar em Search
                # (EXATAMENTE COMO VOCÊ ESCREVEU)
                with page.expect_response(lambda response: response.url and "api/Reporting/Get" in response.url and "GameType" in response.url, timeout=30000) as response_info:
                    page.click(SEL_SEARCH_ADD_BTN)
                    
                # Extrai o JSON direto da resposta e salva no array
                json_do_dia = response_info.value.json()
                dados_json_mensal.append({
                    "Dia": dia,
                    "dados": json_do_dia
                })
                
            except PlaywrightTimeoutError:
                logger.warning(f"Timeout no dia {data_loop}: A API demorou mais de 30s. Ignorando o dia e avançando...")
            except Exception as e:
                logger.error(f"Erro inesperado no dia {data_loop}: {e}")
            
            # Espera um pouco antes de ir para o próximo dia para não derrubar a API
            page.wait_for_timeout(1000)
            
    # Salva os dados no arquivo JSON na pasta do dia
    arq_gs = str(pasta_destino / f"GeneralStats - {marca_arquivo}.json")
    with open(arq_gs, 'w', encoding='utf-8') as f:
        json.dump(dados_json_mensal, f, ensure_ascii=False, indent=4)
        
    arquivos_baixados.append(arq_gs)
    logger.info(f"Salvo (JSON API): {arq_gs}")