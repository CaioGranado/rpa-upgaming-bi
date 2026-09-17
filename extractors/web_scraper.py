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
from utils.date_utils import (
    calcular_limite_seguro,
    obter_data_alvo,
    obter_periodo_extracao,
)
from utils.exceptions_utils import FormatoInvalidoError
from utils.file_utils import (
    obter_pasta_download_diario,
    obter_pasta_ugs_diario,
    validar_formato_xlsx,
)

logger = logging.getLogger(__name__)

# Rótulos de contagem fixa esperados por marca (não inclui UGS_Diario nem
# GeneralStats, que têm contagem variável — tratados via contadores abaixo).
_ROTULOS_FIXOS_ESPERADOS = [
    "NC", "Transacoes", "UGS_Completo", "UGS_ST", "UGS_LC", "UGS_SB", "UGS_MG",
    "FTD",
]


def _novo_checklist() -> dict:
    """Cria um checklist zerado para o início da extração de uma marca."""
    checklist = {rotulo: False for rotulo in _ROTULOS_FIXOS_ESPERADOS}
    checklist["UGS_Diario_esperados"] = 0
    checklist["UGS_Diario_obtidos"] = 0
    # GeneralStats tem loop interno por dia que tolera falha de dias individuais
    # (de propósito — um dia ruim não deveria abortar o mês inteiro). Por isso
    # rastreamos quantos dias eram esperados vs. quantos realmente vieram, em
    # vez de um booleano "arquivo foi salvo" (que seria verdade mesmo faltando
    # metade dos dias).
    checklist["GeneralStats_esperados"] = 0
    checklist["GeneralStats_obtidos"] = 0
    return checklist


def _avaliar_checklist(checklist: dict) -> tuple[bool, list[str]]:
    """
    Avalia se a marca extraiu TODOS os arquivos esperados.
    Retorna (completo, lista_de_faltantes) — a lista nomeia exatamente o
    que não foi obtido, para uso direto na mensagem de log.
    """
    faltantes = [rotulo for rotulo in _ROTULOS_FIXOS_ESPERADOS if not checklist[rotulo]]

    esperados_diario = checklist["UGS_Diario_esperados"]
    obtidos_diario = checklist["UGS_Diario_obtidos"]
    if obtidos_diario < esperados_diario:
        faltantes.append(f"UGS_Diario ({obtidos_diario}/{esperados_diario})")

    esperados_gs = checklist["GeneralStats_esperados"]
    obtidos_gs = checklist["GeneralStats_obtidos"]
    if obtidos_gs < esperados_gs:
        faltantes.append(f"GeneralStats ({obtidos_gs}/{esperados_gs})")

    return (len(faltantes) == 0, faltantes)


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
                    return arquivos_baixados, {}
            else:
                # 3. PROVA REAL DO COOKIE (Garante que não é uma tela de erro 502/Cloudflare)
                logger.info("Avaliando sessão salva no cookie...")
                try:
                    page.wait_for_selector(Seletores.Menu.REPORT, timeout=15000)
                    logger.info("Sessão ativa confirmada! Menu carregado, pulando login manual...")
                except PlaywrightTimeoutError:
                    logger.error("Estado desconhecido! Não é a tela de login, mas o menu não carregou. Possível erro de rede ou bloqueio.")
                    return arquivos_baixados, {}

            logger.info("Avaliando o período de extração...")
            data_inicio, data_fim, data_fim_nc = obter_periodo_extracao()

            marcas_incompletas = {}
            lista_marcas = list(MARCAS_CONFIG.items())

            # --- LOOP DE MARCAS ---
            for indice_marca, (marca_arquivo, marca_bo) in enumerate(lista_marcas):
                logger.info(LogDivisors.SUB)
                logger.info(f" >>> INICIANDO EXTRAÇÃO PARA A MARCA: {marca_arquivo.upper()} <<<")
                logger.info(LogDivisors.SUB)

                checklist = _novo_checklist()
                browser_morreu = False

                try:
                    _extrair_relatorios_marca(
                        page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc,
                        arquivos_baixados, checklist
                    )
                except FormatoInvalidoError as e:
                    # Arquivo chegou num formato inesperado — pipeline dessa marca é interrompido
                    # aqui. O checklist já reflete o que foi obtido até este ponto.
                    logger.error(
                        f"[MARCA INTERROMPIDA] {marca_arquivo.upper()}: formato inválido detectado "
                        f"na extração. Detalhes: {e}"
                    )
                except Exception as e:
                    # TargetClosedError: o browser fechou — o objeto 'page' está morto e não
                    # tem como recuperar no mesmo ciclo. Continuar o loop só geraria o mesmo
                    # erro em cascata para as marcas seguintes. Encerramos o loop aqui.
                    if "TargetClosedError" in type(e).__name__ or "Target page" in str(e):
                        logger.error(
                            f"[BROWSER FECHADO] {marca_arquivo.upper()}: o browser foi encerrado "
                            f"inesperadamente durante a extração desta marca. As marcas restantes "
                            f"não podem ser extraídas neste ciclo — encerrando loop de extração."
                        )
                        browser_morreu = True
                    else:
                        # Qualquer outro erro (timeout, seletor, rede, etc.): loga com traceback
                        # completo para diagnóstico. O checklist reflete o que foi obtido até aqui.
                        logger.exception(
                            f"[ERRO DE EXTRAÇÃO] {marca_arquivo.upper()}: erro inesperado durante "
                            f"a extração."
                        )

                # Avalia o checklist desta marca — independente de ter dado exceção ou não,
                # é a fonte da verdade sobre o que realmente foi obtido.
                completo, faltantes = _avaliar_checklist(checklist)
                if completo:
                    logger.info(f"[MARCA COMPLETA] {marca_arquivo.upper()}: todos os arquivos esperados foram obtidos.")
                else:
                    marcas_incompletas[marca_arquivo] = faltantes
                    logger.warning(
                        f"[MARCA INCOMPLETA] {marca_arquivo.upper()}: faltando {faltantes}. "
                        f"Esta marca será pulada nas Etapas 2 e 3."
                    )

                if browser_morreu:
                    # Marca todas as marcas restantes (que nem chegaram a ser tentadas) como
                    # incompletas também, para que o sinal devolvido seja completo e confiável.
                    for _, (marca_restante, _) in enumerate(lista_marcas[indice_marca + 1:]):
                        marcas_incompletas[marca_restante] = ["TODOS - browser encerrado antes de tentar esta marca"]
                    break

            return arquivos_baixados, marcas_incompletas

    except Exception:
        logger.exception("FALHA CRÍTICA NA EXTRAÇÃO:")
        return arquivos_baixados, {}


def _extrair_relatorios_marca(page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc, arquivos_baixados, checklist):
    """
    Função auxiliar criada para reduzir a 'Complexidade Cognitiva' do código.
    Ela processa os relatórios individualmente para a marca passada.

    `checklist` é um dict mutável (passado por referência) que marca True em
    cada chave conforme o respectivo download é validado com sucesso. Como é
    o mesmo objeto durante toda a chamada, se uma exceção interromper a função
    no meio, o chamador ainda enxerga quais itens ficaram concluídos e quais
    faltaram — usado para decidir se a marca está "completa" o suficiente
    para seguir para as Etapas 2 e 3.
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
    validar_formato_xlsx(Path(arq_nc))
    arquivos_baixados.append(arq_nc)
    checklist["NC"] = True
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
    validar_formato_xlsx(Path(arq_trans))
    arquivos_baixados.append(arq_trans)
    checklist["Transacoes"] = True
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
        validar_formato_xlsx(Path(arq_ugs))
        arquivos_baixados.append(arq_ugs)
        checklist[f"UGS_{sigla}"] = True
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

    checklist["UGS_Diario_esperados"] = len(dias_diarios_faltantes)

    # Agora o Playwright só entra em ação para os dias que realmente faltam
    for dia_alvo in dias_diarios_faltantes:
        d_inicio = dia_alvo.strftime("%d-%m-%Y 00:00")
        # Limite seguro: dia seguinte às 00:00, para não perder o último minuto do dia_alvo.
        d_fim = calcular_limite_seguro(dia_alvo).strftime("%d-%m-%Y %H:%M")
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
        validar_formato_xlsx(Path(arq_ugs_diario))
        arquivos_baixados.append(arq_ugs_diario)
        checklist["UGS_Diario_obtidos"] += 1
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
    validar_formato_xlsx(Path(arq_ftd))
    arquivos_baixados.append(arq_ftd)
    checklist["FTD"] = True
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
    checklist["GeneralStats_esperados"] = dias_fechados

    if dias_fechados > 0:
        for dia in range(1, dias_fechados + 1):
            # Substitui o 'hoje.replace' por uma formatação de data cravada
            data_loop = f"{dia:02d}-{mes_alvo:02d}-{ano_alvo}"
            # Limite seguro: dia seguinte às 00:00, para não perder o último minuto do dia.
            data_loop_fim = calcular_limite_seguro(datetime(ano_alvo, mes_alvo, dia)).strftime("%d-%m-%Y %H:%M")
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
                loc_to.press_sequentially(data_loop_fim, delay=50)
                
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
    checklist["GeneralStats_obtidos"] = len(dados_json_mensal)
    logger.info(
        f"Salvo (JSON API): {arq_gs} — {len(dados_json_mensal)}/{dias_fechados} dias obtidos."
    )