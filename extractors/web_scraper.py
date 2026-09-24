import json
import logging
import time
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
from utils.generalstats_utils import dia_generalstats_e_confiavel

logger = logging.getLogger(__name__)

# Rótulos de contagem fixa esperados por marca (não inclui UGS_Diario nem
# GeneralStats, que têm contagem variável — tratados via contadores abaixo).
_ROTULOS_FIXOS_ESPERADOS = [
    "NC", "Transacoes", "UGS_Completo", "UGS_ST", "UGS_LC", "UGS_SB", "UGS_MG",
    "FTD",
]


def _novo_checklist() -> dict:
    """Cria um checklist zerado para o início da extração de uma marca."""
    checklist = dict.fromkeys(_ROTULOS_FIXOS_ESPERADOS, False)
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


def _lancar_navegador_persistente(p):
    """
    Abre o Chrome com perfil persistente e o disfarce de automação
    aplicado. Não navega nem faz login — só devolve a 'page' pronta
    para uso.
    """
    pasta_perfil = str(Path.cwd() / "perfil_robo_chrome")
    context = p.chromium.launch_persistent_context(
        user_data_dir=pasta_perfil,
        headless=False,
        channel="chrome", 
        chromium_sandbox=True, 
        ignore_default_args=["--no-sandbox", "--enable-automation"],
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        viewport={'width': 1280, 'height': 720}
    )

    # Substitui a flag '--disable-blink-features=AutomationControlled' (que o
    # Chrome sinaliza com o aviso "linha de comando não suportada") por um
    # init_script equivalente: sobrescreve navigator.webdriver via JS, antes de
    # qualquer página carregar. Mesmo efeito de disfarce, sem o aviso visível —
    # e tecnicamente mais discreto, já que não depende de uma flag de linha de
    # comando que sites de detecção anti-bot também podem checar.
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
    )

    page = context.pages[0]
    page.set_default_timeout(300000)
    return page


def _confirmar_sessao(page):
    """
    Navega até o sistema e garante que a sessão está ativa — via cookie
    salvo, ou pedindo login manual se necessário.

    Retorna (sucesso, motivo_falha). motivo_falha só é preenchido
    (como lista, pronta para virar 'faltantes' do chamador) quando
    sucesso=False.
    """
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
    is_login_visual = page.locator('text="Welcome To Admin Panel"').is_visible() or page.locator('input[name="username"]').is_visible()

    # Se a URL acusou login OU a tela inicial apareceu, pede intervenção
    if is_login_url or is_login_visual:
        logger.warning("Página de login detectada (Sessão expirada).")
        input("\n>>> Faça o login e resolva os reCAPTCHA, espere o painel inicial carregar e então pressione ENTER aqui...\n")

        # 2. VALIDAÇÃO PÓS-LOGIN (Garante que a barra lateral apareceu)
        try:
            page.wait_for_selector(Seletores.Menu.REPORT, timeout=15000)
            logger.info("Login confirmado com sucesso!")
            return True, None
        except PlaywrightTimeoutError:
            logger.error("Falha ao confirmar o login: Menu lateral não encontrado, abortando por segurança.")
            return False, ["Falha ao confirmar login"]
    else:
        # 3. PROVA REAL DO COOKIE (Garante que não é uma tela de erro 502/Cloudflare)
        logger.info("Avaliando sessão salva no cookie...")
        try:
            page.wait_for_selector(Seletores.Menu.REPORT, timeout=15000)
            logger.info("Sessão ativa confirmada! Menu carregado, pulando login manual...")
            return True, None
        except PlaywrightTimeoutError:
            logger.error("Estado desconhecido! Não é a tela de login, mas o menu não carregou. Possível erro de rede ou bloqueio.")
            return False, ["Estado de sessão desconhecido"]


def _extrair_com_tratamento_erros(page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc, arquivos_baixados, checklist):
    """
    Roda _extrair_relatorios_marca e categoriza o resultado — sucesso,
    formato inválido, crash de navegador, ou erro inesperado.

    Retorna (completo, faltantes, browser_morreu), sempre calculado a
    partir do checklist real: mesmo quando uma exceção interrompe a
    extração no meio, o checklist já reflete o que foi obtido até ali,
    então o retorno reflete o estado verdadeiro, não um "tudo ou nada".
    """
    try:
        _extrair_relatorios_marca(
            page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc,
            arquivos_baixados, checklist
        )
        completo, faltantes = _avaliar_checklist(checklist)
        return completo, faltantes, False

    except FormatoInvalidoError as e:
        # Arquivo chegou num formato inesperado — o checklist já reflete o que foi
        # obtido até este ponto.
        logger.error(
            f"[MARCA INTERROMPIDA] {marca_arquivo.upper()}: formato inválido detectado "
            f"na extração. Detalhes: {e}"
        )
        completo, faltantes = _avaliar_checklist(checklist)
        return completo, faltantes, False

    except Exception as e:
        # TargetClosedError: o browser fechou — o objeto 'page' está morto, não tem
        # como continuar nesta tentativa. browser_morreu=True só afeta o texto do log
        # da próxima tentativa (extrair_dados_upgaming decide o retry de qualquer forma).
        if "TargetClosedError" in type(e).__name__ or "Target page" in str(e):
            logger.error(
                f"[BROWSER FECHADO] {marca_arquivo.upper()}: o browser foi encerrado "
                f"inesperadamente durante a extração desta marca."
            )
            completo, faltantes = _avaliar_checklist(checklist)
            return completo, faltantes, True
        else:
            logger.exception(
                f"[ERRO DE EXTRAÇÃO] {marca_arquivo.upper()}: erro inesperado durante "
                f"a extração."
            )
            completo, faltantes = _avaliar_checklist(checklist)
            return completo, faltantes, False


def _extrair_uma_marca_uma_tentativa(marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc):
    """
    Uma tentativa completa de extração de UMA marca: abre um navegador
    novo, confirma login/sessão, extrai os relatórios dessa marca, e
    fecha tudo. Cada marca tem seu próprio orçamento de tentativas (ver
    extrair_dados_upgaming, logo abaixo) — uma marca com crash crônico
    nunca consome as tentativas de outra marca.

    Orquestra 3 funções, cada uma com uma responsabilidade só:
    _lancar_navegador_persistente (abre o browser), _confirmar_sessao
    (login/cookie), _extrair_com_tratamento_erros (roda a extração e
    categoriza o resultado). Nenhuma lógica de negócio mora aqui — só a
    sequência de passos e o que fazer quando um deles falha.

    Retorna (arquivos_baixados, completo, faltantes, browser_morreu).
    completo/faltantes vêm direto de _avaliar_checklist(). browser_morreu
    sinaliza que a falha foi especificamente um crash de navegador
    (TargetClosedError) — informação usada só para a mensagem de log da
    tentativa seguinte, não muda a lógica de retry em si (qualquer
    motivo de incompletude gera nova tentativa, dentro do orçamento).
    """
    logger.info("Iniciando módulo de Extração Web...")
    arquivos_baixados = []

    try:
        with sync_playwright() as p:
            page = _lancar_navegador_persistente(p)

            sessao_ok, motivo_falha = _confirmar_sessao(page)
            if not sessao_ok:
                return arquivos_baixados, False, motivo_falha, False

            # =========================================================
            # PAUSA DE ESTABILIZAÇÃO: padrão observado em produção mostra
            # o TargetClosedError concentrado perto do primeiro download
            # (NC), logo após um navegador recém-lançado — nunca depois
            # que a sessão já processou algo. Como o retry agora é por
            # marca (cada tentativa abre um navegador novo), essa janela
            # de instabilidade pós-lançamento passou a se repetir a cada
            # marca, não só uma vez por execução. Essa pausa é uma
            # hipótese fundamentada nesse padrão, não uma certeza — o
            # retry por marca continua como rede de segurança de
            # qualquer forma.
            # =========================================================
            page.wait_for_timeout(3000)

            checklist = _novo_checklist()
            completo, faltantes, browser_morreu = _extrair_com_tratamento_erros(
                page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc,
                arquivos_baixados, checklist
            )
            return arquivos_baixados, completo, faltantes, browser_morreu

    except Exception:
        logger.exception(f"FALHA CRÍTICA NA EXTRAÇÃO DE {marca_arquivo.upper()}:")
        return arquivos_baixados, False, ["Falha crítica antes de iniciar a extração"], True


def extrair_dados_upgaming():
    """
    Ponto de entrada público — mesma assinatura de antes, usada pelo
    main.py sem nenhuma mudança na chamada.

    Cada marca recebe seu PRÓPRIO orçamento de MAX_TENTATIVAS_POR_MARCA
    tentativas, totalmente independente das demais — uma marca com crash
    crônico (ex: sempre falha logo no primeiro download) nunca consome
    as tentativas de outra marca. Diferente do desenho anterior (retry
    de sessão inteira compartilhado entre marcas), aqui não existe mais
    risco de uma marca "azarada" consumir o orçamento e deixar as
    seguintes sem nenhuma tentativa real.

    Uma marca que complete plenamente (em qualquer tentativa, a 1ª ou a
    última) segue normalmente para as Etapas 2 e 3. Só as marcas que
    esgotarem seu próprio orçamento sem completar ficam de fora.
    """
    logger.info("Iniciando módulo de Extração Web...")
    logger.info("Avaliando o período de extração...")
    data_inicio, data_fim, data_fim_nc = obter_periodo_extracao()

    MAX_TENTATIVAS_POR_MARCA = 5
    ESPERAS_ENTRE_TENTATIVAS = [10, 20, 30, 45]  # segundos: cresce a cada tentativa nova, mantém o último valor se sobrar

    arquivos_baixados_total = []
    marcas_incompletas_final = {}

    for marca_arquivo, marca_bo in MARCAS_CONFIG.items():
        logger.info(LogDivisors.SUB)
        logger.info(f" >>> INICIANDO EXTRAÇÃO PARA A MARCA: {marca_arquivo.upper()} <<<")
        logger.info(LogDivisors.SUB)

        sucesso = False
        ultimo_faltantes = ["motivo desconhecido"]

        for tentativa in range(1, MAX_TENTATIVAS_POR_MARCA + 1):
            arquivos, completo, faltantes, browser_morreu = _extrair_uma_marca_uma_tentativa(
                marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc
            )

            if completo:
                arquivos_baixados_total.extend(arquivos)
                logger.info(f"[MARCA COMPLETA] {marca_arquivo.upper()}: todos os arquivos esperados foram obtidos.")
                sucesso = True
                break

            ultimo_faltantes = faltantes
            tentativas_restantes = MAX_TENTATIVAS_POR_MARCA - tentativa

            if tentativas_restantes > 0:
                espera = ESPERAS_ENTRE_TENTATIVAS[min(tentativa - 1, len(ESPERAS_ENTRE_TENTATIVAS) - 1)]
                motivo_retry = "crash de navegador" if browser_morreu else f"faltando {faltantes}"
                logger.warning(
                    f"[MARCA REINICIADA] {marca_arquivo.upper()}: tentativa {tentativa}/{MAX_TENTATIVAS_POR_MARCA} "
                    f"incompleta ({motivo_retry}). Aguardando {espera}s e tentando novamente..."
                )
                time.sleep(espera)

        if not sucesso:
            marcas_incompletas_final[marca_arquivo] = ultimo_faltantes
            logger.error(
                f"[MARCA FALHOU] {marca_arquivo.upper()}: esgotou as {MAX_TENTATIVAS_POR_MARCA} tentativas "
                f"(faltando: {ultimo_faltantes}). Esta marca será pulada nas Etapas 2 e 3."
            )

    return arquivos_baixados_total, marcas_incompletas_final


def _salvar_e_registrar_download(download_info, caminho_arquivo, arquivos_baixados, atualizar_checklist, rotulo_log="Salvo"):
    """
    Parte final, idêntica em todo download simples do BackOffice: salva o
    arquivo, valida o formato, registra na lista de arquivos baixados,
    atualiza o checklist e loga.

    `atualizar_checklist` é um callback sem argumentos (ex: lambda) porque
    a atualização do checklist varia entre os chamadores — às vezes é
    "marcar True" numa chave fixa, às vezes é "incrementar um contador"
    (UGS Diário). Manter isso como parâmetro evita forçar uma uniformidade
    que não existe de verdade entre os relatórios.

    A parte ANTES disso (abrir o `with page.expect_download(...)` e
    clicar no botão certo) fica no chamador, porque o gatilho do download
    muda de relatório para relatório — Transações, por exemplo, precisa
    de uma confirmação extra antes do download real começar.
    """
    download_info.value.save_as(caminho_arquivo)
    validar_formato_xlsx(Path(caminho_arquivo))
    arquivos_baixados.append(caminho_arquivo)
    atualizar_checklist()
    logger.info(f"{rotulo_log}: {caminho_arquivo}")


def _baixar_nc(page, marca_arquivo, marca_bo, data_inicio, data_fim_nc, pasta_destino, arquivos_baixados, checklist):
    """[1/6] Novas Contas."""
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
    _salvar_e_registrar_download(
        download_info, arq_nc, arquivos_baixados,
        atualizar_checklist=lambda: checklist.__setitem__("NC", True),
    )


def _baixar_transacoes(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist):
    """[2/6] System Transactions."""
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
    _salvar_e_registrar_download(
        download_info, arq_trans, arquivos_baixados,
        atualizar_checklist=lambda: checklist.__setitem__("Transacoes", True),
    )


def _baixar_ugs_acumulado(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist):
    """[3/6] UGS Acumulado (Completo + 4 tipos)."""
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
        _salvar_e_registrar_download(
            download_info, arq_ugs, arquivos_baixados,
            atualizar_checklist=lambda s=sigla: checklist.__setitem__(f"UGS_{s}", True),
        )


def _detectar_lacunas_ugs_diario(marca_arquivo, janela_dias=7):
    """
    Verifica os últimos `janela_dias` dias (sem contar hoje) e devolve
    quais ainda não têm arquivo salvo em disco. Checagem pura de
    sistema de arquivos — nenhuma interação com o navegador aqui, o que
    a torna testável sem mockar 'page'.
    """
    dias_faltantes = []
    hoje_real = datetime.now(timezone.utc).astimezone()

    # Loop de trás para frente (ex: dia -7 até dia -1) para manter a ordem cronológica
    for i in range(janela_dias, 0, -1):
        dia_checar = hoje_real - timedelta(days=i)
        pasta_ugs_checar = obter_pasta_ugs_diario(marca_arquivo, dia_checar.year, dia_checar.month)
        nome_dia_checar = dia_checar.strftime("%d-%m")

        arquivo_esperado = pasta_ugs_checar / f"{nome_dia_checar}.xlsx"

        # Se o arquivo não existe fisicamente na pasta, entra na lista de download
        if not arquivo_esperado.exists():
            dias_faltantes.append(dia_checar)

    return dias_faltantes


def _baixar_ugs_diario(page, marca_arquivo, arquivos_baixados, checklist):
    """[4/6] UGS Diário (Buscador Dinâmico de Lacunas)."""
    page.select_option(Seletores.Filtros.GAME_TYPE, value="")

    # Variável ajustável: Quantos dias no passado o robô deve checar?
    JANELA_DIAS = 7
    logger.info(f"Checando lacunas de UGS Diário nos últimos {JANELA_DIAS} dias...")
    dias_diarios_faltantes = _detectar_lacunas_ugs_diario(marca_arquivo, JANELA_DIAS)

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

        _salvar_e_registrar_download(
            download_info, arq_ugs_diario, arquivos_baixados,
            atualizar_checklist=lambda: checklist.__setitem__(
                "UGS_Diario_obtidos", checklist["UGS_Diario_obtidos"] + 1
            ),
            rotulo_log="Salvo UGS Diário",
        )


def _baixar_ftd(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist):
    """[5/6] FTD."""
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
    _salvar_e_registrar_download(
        download_info, arq_ftd, arquivos_baixados,
        atualizar_checklist=lambda: checklist.__setitem__("FTD", True),
    )


def _extrair_dia_generalstats(page, data_loop, data_loop_fim, max_tentativas=2):
    """
    Tenta extrair o JSON de UM dia do General Statistics, com retry se o
    dado vier suspeito (vazio, ou vertical sempre-ativa zerada) — só
    aqui, com o navegador ainda autenticado, é possível tentar de novo.

    Retorna (json_do_dia_ou_none, motivo_falha). motivo_falha só importa
    quando o retorno é None — é o que o chamador loga.
    """
    json_do_dia_valido = None
    motivo_falha = "falha na requisição"

    # =========================================================
    # AQUI ESTÁ O SEGREDO: Se um dia falhar, não aborta tudo!
    # Além disso, um dia que "funciona" tecnicamente (sem timeout)
    # mas vem com dado suspeito (vazio, ou verticais sempre-ativas
    # zeradas) é tratado como falha e ganha 1 retry antes de
    # desistir — só aqui, com o navegador ainda autenticado, é
    # possível tentar o dia de novo.
    # =========================================================
    for tentativa in range(1, max_tentativas + 1):
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

            # Extrai o JSON direto da resposta e valida a confiabilidade
            candidato = response_info.value.json()
            confiavel, motivo_falha = dia_generalstats_e_confiavel(candidato)

            if confiavel:
                json_do_dia_valido = candidato
                break

            tentativas_restantes = max_tentativas - tentativa
            logger.warning(
                f"Dia {data_loop} (tentativa {tentativa}/{max_tentativas}): "
                f"dado suspeito — {motivo_falha}. "
                + ("Tentando novamente..." if tentativas_restantes > 0 else "Desistindo após retry.")
            )
            if tentativas_restantes > 0:
                page.wait_for_timeout(1500)

        except PlaywrightTimeoutError:
            motivo_falha = "timeout — API demorou mais de 30s"
            logger.warning(
                f"Timeout no dia {data_loop} (tentativa {tentativa}/{max_tentativas})."
            )
        except Exception as e:
            motivo_falha = f"erro inesperado: {e}"
            logger.error(
                f"Erro inesperado no dia {data_loop} (tentativa {tentativa}/{max_tentativas}): {e}"
            )

    return json_do_dia_valido, motivo_falha


def _salvar_generalstats_mensal(dados_json_mensal, pasta_destino, marca_arquivo, dias_fechados, arquivos_baixados, checklist):
    """
    Salva o JSON consolidado do mês em disco e atualiza checklist e
    arquivos_baixados. I/O puro — nenhuma interação com o navegador.
    """
    arq_gs = str(pasta_destino / f"GeneralStats - {marca_arquivo}.json")
    with open(arq_gs, 'w', encoding='utf-8') as f:
        json.dump(dados_json_mensal, f, ensure_ascii=False, indent=4)

    arquivos_baixados.append(arq_gs)
    checklist["GeneralStats_obtidos"] = len(dados_json_mensal)
    logger.info(
        f"Salvo (JSON API): {arq_gs} — {len(dados_json_mensal)}/{dias_fechados} dias obtidos."
    )


def _baixar_general_stats(page, marca_arquivo, marca_bo, pasta_destino, arquivos_baixados, checklist):
    """[6/6] General Statistics (Scraping da API Invisível)."""
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

            json_do_dia, motivo_falha = _extrair_dia_generalstats(page, data_loop, data_loop_fim)

            if json_do_dia is not None:
                dados_json_mensal.append({
                    "Dia": dia,
                    "dados": json_do_dia
                })
            else:
                logger.error(
                    f"Dia {data_loop}: dado não confiável mesmo após retry ({motivo_falha}). "
                    f"Este dia NÃO será contado como obtido."
                )

            # Espera um pouco antes de ir para o próximo dia para não derrubar a API
            page.wait_for_timeout(1000)

    _salvar_generalstats_mensal(dados_json_mensal, pasta_destino, marca_arquivo, dias_fechados, arquivos_baixados, checklist)


def _extrair_relatorios_marca(page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc, arquivos_baixados, checklist):
    """
    Orquestra a extração dos 6 relatórios de uma marca, na ordem certa.

    `checklist` é um dict mutável (passado por referência) que marca True em
    cada chave conforme o respectivo download é validado com sucesso. Como é
    o mesmo objeto durante toda a chamada, se uma exceção interromper a função
    no meio, o chamador ainda enxerga quais itens ficaram concluídos e quais
    faltaram — usado para decidir se a marca está "completa" o suficiente
    para seguir para as Etapas 2 e 3.

    A lógica de cada relatório vive em sua própria função (_baixar_nc,
    _baixar_transacoes, etc.) — esta função só orquestra a ordem de
    chamada, sem repetir nenhuma lógica de negócio.
    """
    logger.info(LogDivisors.MAIN)
    logger.info(f" EXTRAINDO MARCA: {marca_arquivo}")
    logger.info(LogDivisors.MAIN)

    # 2. PEGA A PASTA DE DOWNLOAD CORRETA LÁ NO DRIVE G:
    pasta_destino = obter_pasta_download_diario(marca_arquivo)

    _baixar_nc(page, marca_arquivo, marca_bo, data_inicio, data_fim_nc, pasta_destino, arquivos_baixados, checklist)
    _baixar_transacoes(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist)
    _baixar_ugs_acumulado(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist)
    _baixar_ugs_diario(page, marca_arquivo, arquivos_baixados, checklist)
    _baixar_ftd(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist)
    _baixar_general_stats(page, marca_arquivo, marca_bo, pasta_destino, arquivos_baixados, checklist)