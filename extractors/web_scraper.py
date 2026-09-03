import json
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from config.settings import (
    MARCAS_CONFIG,
    URL_SISTEMA,
    LogDivisors,
    Seletores,
)
from utils.date_utils import obter_data_alvo, obter_periodo_extracao
from utils.file_utils import obter_pasta_download_diario, obter_pasta_ugs_diario

logger = logging.getLogger(__name__)

def extrair_dados_upgaming():
    logger.info("Iniciando módulo de Extração Web...")
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
                    page.wait_for_selector(Seletores.Menu.REPORT, timeout=15000)
                    logger.info("Login confirmado com sucesso!")
                except PlaywrightTimeoutError:
                    logger.error("Falha ao confirmar o login: Menu lateral não encontrado, abortando por segurança.")
                    return arquivos_baixados
            else:
                # 3. PROVA REAL DO COOKIE (Garante que não é uma tela de erro 502/Cloudflare)
                logger.info("Avaliando sessão salva no cookie...")
                try:
                    page.wait_for_selector(Seletores.Menu.REPORT, timeout=15000)
                    logger.info("Sessão ativa confirmada! Menu carregado, pulando login manual...")
                except PlaywrightTimeoutError:
                    logger.error("Estado desconhecido! Não é a tela de login, mas o menu não carregou. Possível erro de rede ou bloqueio.")
                    return arquivos_baixados

            logger.info("Avaliando o período de extração...")
            data_inicio, data_fim, data_fim_nc = obter_periodo_extracao()
            
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
    logger.info(LogDivisors.MAIN)
    logger.info(f" EXTRAINDO MARCA: {marca_arquivo}")
    logger.info(LogDivisors.MAIN)
    
    # 2. PEGA A PASTA DE DOWNLOAD CORRETA LÁ NO DRIVE G:
    pasta_destino = obter_pasta_download_diario(marca_arquivo)
    
    # [1/6] NOVAS CONTAS
    page.click(Seletores.Menu.USERS)
    page.wait_for_timeout(2000)
    try:
        page.wait_for_selector(Seletores.Filtros.BRAND_DROPDOWN, timeout=4000)
        page.select_option(Seletores.Filtros.BRAND_DROPDOWN, label=marca_bo)
        page.click(Seletores.Botoes.BRAND_OK)
    except PlaywrightTimeoutError:
        page.click(Seletores.Filtros.SEARCH_BRAND_INPUT)
        page.fill(Seletores.Filtros.SEARCH_BRAND_INPUT, marca_bo)
        page.wait_for_timeout(1000)
        page.click(f'text="{marca_bo}" >> visible=true')
        page.wait_for_timeout(500)
    
    page.click(Seletores.Botoes.MORE_FILTER)
    page.fill(Seletores.Filtros.CREATE_DATE_FROM, data_inicio)
    page.fill(Seletores.Filtros.CREATE_DATE_TO, data_fim_nc)
    page.click(Seletores.Botoes.SEARCH_ALT)
    page.wait_for_timeout(6000) 
    
    with page.expect_download(timeout=120000) as download_info:
        page.click(Seletores.Botoes.EXPORT)
    
    arq_nc = str(pasta_destino / f"NC - {marca_arquivo}.xlsx")
    download_info.value.save_as(arq_nc)
    arquivos_baixados.append(arq_nc)
    logger.info(f"Salvo: {arq_nc}")

    # [2/6] SYSTEM TRANSACTIONS
    page.click(Seletores.Menu.TRANSACTIONS)
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
    
    page.fill(Seletores.Filtros.DATE_FROM, data_inicio)
    page.fill(Seletores.Filtros.DATE_TO, data_fim)
    page.click(Seletores.Botoes.OK)
    
    try:
        page.click(Seletores.Botoes.SEARCH_ADD, timeout=3000)
        logger.debug("Botão SEARCH clicado manualmente.")
    except PlaywrightTimeoutError:
        logger.debug("Botão SEARCH ignorado (busca automática acionada ou botão ausente).")
    page.wait_for_timeout(6000) 
    
    page.click(Seletores.Botoes.EXPORT)
    page.wait_for_selector(Seletores.Botoes.CONFIRM_EXPORT, timeout=10000)
    
    with page.expect_download(timeout=120000) as download_info:
        page.click(Seletores.Botoes.CONFIRM_EXPORT)
    
    arq_trans = str(pasta_destino / f"Transações - {marca_arquivo}.xlsx")
    download_info.value.save_as(arq_trans)
    arquivos_baixados.append(arq_trans)
    logger.info(f"Salvo: {arq_trans}")

    # [3/6] UGS ACUMULADO
    page.click(Seletores.Menu.UGS)
    page.wait_for_timeout(3000)
    page.click(Seletores.Filtros.SEARCH_BRAND_INPUT)
    page.fill(Seletores.Filtros.SEARCH_BRAND_INPUT, marca_bo)
    page.wait_for_timeout(1000)
    page.click(f'text="{marca_bo}" >> visible=true')
    page.wait_for_timeout(500)
    
    page.fill(Seletores.Filtros.DATE_FROM, data_inicio)
    page.fill(Seletores.Filtros.DATE_TO, data_fim)
    page.click(Seletores.Botoes.OK)
    
    tipos_ugs = {"": "Completo", "1": "ST", "2": "LC", "7": "SB", "8": "MG"}
    for valor, sigla in tipos_ugs.items():
        page.select_option(Seletores.Filtros.GAME_TYPE, value=valor)
        page.click(Seletores.Botoes.SEARCH_ADD)
        page.wait_for_timeout(6000) 
        with page.expect_download(timeout=120000) as download_info:
            page.click(Seletores.Botoes.EXPORT)
        arq_ugs = str(pasta_destino / f"{marca_arquivo} - UGS {sigla}.xlsx")
        download_info.value.save_as(arq_ugs)
        arquivos_baixados.append(arq_ugs)
        logger.info(f"Salvo: {arq_ugs}")

    # [4/6] UGS DIÁRIO (Buscador Dinâmico de Lacunas)
    page.select_option(Seletores.Filtros.GAME_TYPE, value="")
    
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
        
        page.fill(Seletores.Filtros.DATE_FROM, d_inicio)
        page.fill(Seletores.Filtros.DATE_TO, d_fim)
        page.click(Seletores.Botoes.OK)
        page.click(Seletores.Botoes.SEARCH_ADD)
        page.wait_for_timeout(6000)
        
        with page.expect_download(timeout=120000) as download_info:
            page.click(Seletores.Botoes.EXPORT)
        
        pasta_ugs_alvo = obter_pasta_ugs_diario(marca_arquivo, dia_alvo.year, dia_alvo.month)
        arq_ugs_diario = str(pasta_ugs_alvo / f"{nome_dia}.xlsx")
        
        download_info.value.save_as(arq_ugs_diario)
        arquivos_baixados.append(arq_ugs_diario)
        logger.info(f"Salvo UGS Diário: {arq_ugs_diario}")

    # [5/6] FTD
    page.click(Seletores.Menu.FTD)
    page.wait_for_timeout(3000)
    page.click(Seletores.Filtros.SEARCH_BRAND_INPUT)
    page.fill(Seletores.Filtros.SEARCH_BRAND_INPUT, marca_bo)
    page.wait_for_timeout(1000)
    page.click(f'text="{marca_bo}" >> visible=true')
    page.wait_for_timeout(500)
    page.fill(Seletores.Filtros.DATE_FROM, data_inicio)
    page.fill(Seletores.Filtros.DATE_TO, data_fim)
    page.click(Seletores.Botoes.OK)
    page.click(Seletores.Botoes.SEARCH_FTD)
    page.wait_for_timeout(6000) 
    with page.expect_download(timeout=120000) as download_info:
        page.click(Seletores.Botoes.EXPORT)
    
    arq_ftd = str(pasta_destino / f"FTD - {marca_arquivo}.xlsx")
    download_info.value.save_as(arq_ftd)
    arquivos_baixados.append(arq_ftd)
    logger.info(f"Salvo: {arq_ftd}")
    
    # =====================================================================
    # [6/6] GENERAL STATISTICS (Scraping da API Invisível)
    # =====================================================================
    logger.info("Extraindo General Statistics via API (JSON)...")
    
    page.click(Seletores.Menu.REPORT)
    page.wait_for_timeout(500)
    page.click(Seletores.Menu.GEN_STATS)
    page.wait_for_timeout(3000)
    
    # Seleciona a marca
    page.click(Seletores.Filtros.SEARCH_BRAND_INPUT)
    page.fill(Seletores.Filtros.SEARCH_BRAND_INPUT, marca_bo)
    page.wait_for_timeout(1000)
    page.click(f'text="{marca_bo}" >> visible=true')
    page.wait_for_timeout(500)

    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except PlaywrightTimeoutError:
        pass

    page.wait_for_timeout(2000)

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
                loc_from = page.locator(Seletores.Filtros.DATE_FROM)
                loc_from.click()
                loc_from.clear()
                loc_from.press_sequentially(f"{data_loop} 00:00", delay=50)
                
                loc_to = page.locator(Seletores.Filtros.DATE_TO)
                loc_to.click()
                loc_to.clear()
                loc_to.press_sequentially(f"{data_loop} 23:59", delay=50)
                
                page.click(Seletores.Botoes.OK)
                
                # OBRIGATÓRIO: Dar 1 segundo para o site "entender" a data antes do Search
                page.wait_for_timeout(1000)
                
                # Escuta a aba "Network" e intercepta a requisição assim que clicar em Search
                # (EXATAMENTE COMO VOCÊ ESCREVEU)
                with page.expect_response(lambda response: response.url and "api/Reporting/Get" in response.url and "GameType" in response.url, timeout=30000) as response_info:
                    page.click(Seletores.Botoes.SEARCH_ADD)
                    
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