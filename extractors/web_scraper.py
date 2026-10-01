import json
import logging
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import openpyxl
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from config.settings import (
    ACEITAR_DIAS_ZERADOS,
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
    _detectar_formato_real,
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
    """
    Cria o checklist de progresso de uma marca.

    Um checklist pertence à MARCA, não à tentativa: o mesmo objeto acompanha
    todas as tentativas, e cada relatório concluído fica marcado para ser
    pulado na tentativa seguinte (retomada), em vez de baixado de novo. Por
    isso ele também guarda os dados já obtidos do General Statistics, dia a
    dia (GeneralStats_dias), para que só os dias que faltam sejam pedidos de
    novo.
    """
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
    checklist["GeneralStats_dias"] = {}  # {numero_do_dia: json_do_dia} já obtidos
    # Só usados com ACEITAR_DIAS_ZERADOS (ver settings): quantas vezes cada UGS Diário
    # veio quebrado, e quais dias acabaram aceitos como zerados (para o resumo no log).
    checklist["UGS_Diario_falhas"] = {}
    checklist["Dias_zerados_aceitos"] = []
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


def _fechar_navegador_com_seguranca(page):
    """
    Fecha o contexto do navegador de forma defensiva — usado sempre que
    vamos abandonar um 'page' (crashou ou a sessão falhou), para
    garantir que a pasta de perfil (user_data_dir) seja liberada antes
    da próxima tentativa tentar abrir um navegador novo no mesmo
    perfil. Se o navegador já morreu sozinho (TargetClosedError),
    fechar de novo pode falhar — isso é esperado e é só ignorado.
    """
    try:
        page.context.close()
    except Exception:
        pass


def _extrair_uma_marca_uma_tentativa(p, page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc,
                                     arquivos_baixados, checklist):
    """
    Uma tentativa de extração de UMA marca.

    `arquivos_baixados` e `checklist` pertencem à MARCA, não à tentativa:
    são criados por extrair_dados_upgaming e chegam aqui já com o progresso
    das tentativas anteriores. Os relatórios que já constam como obtidos são
    pulados — a tentativa só refaz o que ainda falta. Os dois objetos são
    alterados no próprio lugar, então o chamador enxerga o progresso mesmo
    quando a tentativa termina em exceção.

    Se `page` vier None, abre um navegador novo e confirma a sessão
    antes de extrair — isso acontece na primeira tentativa da primeira
    marca, e sempre que a tentativa anterior (desta marca ou da
    anterior) crashou. Se `page` já vier de uma marca/tentativa
    anterior que terminou saudável, pula direto para a extração — sem
    gastar tempo reabrindo o Chrome nem re-confirmando login à toa.

    Retorna (page, completo, faltantes, browser_morreu). O `page`
    retornado é a mesma instância recebida (se ainda viva) ou None (se
    crashou, ou se nunca chegou a abrir) — o chamador
    (extrair_dados_upgaming) usa isso para decidir se relança o
    navegador na próxima tentativa, seja dela mesma ou da marca seguinte.

    Cada marca tem seu próprio orçamento de tentativas (ver
    extrair_dados_upgaming, logo abaixo) — uma marca com crash crônico
    nunca consome as tentativas de outra marca. Isso é independente de
    reaproveitar ou não o navegador: o orçamento é sobre contagem de
    tentativas, o reaproveitamento é só sobre evitar reabrir o Chrome
    à toa quando ele está saudável.
    """
    if page is None:
        logger.info("Iniciando módulo de Extração Web...")
        try:
            page = _lancar_navegador_persistente(p)
        except Exception:
            logger.exception(f"FALHA CRÍTICA AO ABRIR NAVEGADOR PARA {marca_arquivo.upper()}:")
            return None, False, ["Falha crítica ao abrir o navegador"], True

        sessao_ok, motivo_falha = _confirmar_sessao(page)
        if not sessao_ok:
            _fechar_navegador_com_seguranca(page)
            return None, False, motivo_falha, False

        # =========================================================
        # PAUSA DE ESTABILIZAÇÃO: padrão observado em produção mostra
        # o TargetClosedError concentrado perto do primeiro download
        # (NC), logo após um navegador recém-lançado — nunca depois
        # que a sessão já processou algo. Só faz sentido logo após
        # ABRIR um navegador novo — um 'page' reaproveitado de uma
        # marca anterior saudável não precisa dessa pausa de novo.
        # =========================================================
        page.wait_for_timeout(3000)
    else:
        logger.info(f"Reaproveitando navegador já aberto para {marca_arquivo.upper()}.")

    try:
        completo, faltantes, browser_morreu = _extrair_com_tratamento_erros(
            page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc,
            arquivos_baixados, checklist
        )
    except Exception:
        logger.exception(f"FALHA CRÍTICA NA EXTRAÇÃO DE {marca_arquivo.upper()}:")
        _fechar_navegador_com_seguranca(page)
        return None, False, ["Falha crítica durante a extração"], True

    if browser_morreu:
        _fechar_navegador_com_seguranca(page)
        return None, completo, faltantes, browser_morreu

    # Navegador segue saudável — devolvido para a próxima tentativa (desta
    # marca, se incompleta por outro motivo, ou da marca seguinte) reaproveitar.
    return page, completo, faltantes, browser_morreu


def extrair_dados_upgaming():
    """
    Ponto de entrada público — mesma assinatura de antes, usada pelo
    main.py sem nenhuma mudança na chamada.

    Cada marca recebe seu PRÓPRIO orçamento de MAX_TENTATIVAS_POR_MARCA
    tentativas, totalmente independente das demais — uma marca com crash
    crônico (ex: sempre falha logo no primeiro download) nunca consome
    as tentativas de outra marca. Uma marca que complete plenamente (em
    qualquer tentativa, a 1ª ou a última) segue normalmente para as
    Etapas 2 e 3. Só as marcas que esgotarem seu próprio orçamento sem
    completar ficam de fora.

    Cada nova tentativa RETOMA a marca: o progresso (lista de arquivos e
    checklist) é criado uma vez por marca e atravessa as tentativas, então
    um relatório já obtido não é baixado de novo — se o UGS Diário falha,
    a próxima tentativa refaz só o UGS Diário (e o que vier depois), não a
    marca inteira. Os arquivos só seguem para as Etapas 2 e 3 quando a
    marca fica completa.

    O navegador é reaproveitado entre marcas e tentativas ENQUANTO
    estiver saudável — só é relançado quando de fato crasha (ou quando
    a confirmação de sessão falha). Isso evita pagar o "preço" de abrir
    um Chrome novo (e a instabilidade que isso historicamente trouxe
    logo no primeiro download) toda vez que a marca muda, mesmo quando
    tudo está indo bem. `page` é a variável que carrega esse estado
    entre as iterações do loop: None significa "preciso abrir um novo
    na próxima tentativa", qualquer outro valor significa "este aqui
    ainda está de pé, reaproveite".
    """
    logger.info("Iniciando módulo de Extração Web...")
    logger.info("Avaliando o período de extração...")
    data_inicio, data_fim, data_fim_nc = obter_periodo_extracao()
    if ACEITAR_DIAS_ZERADOS:
        logger.warning(
            "[MODO TESTE] ACEITAR_DIAS_ZERADOS=true: dias sem dados reais serão ACEITOS como zero depois do "
            "retry. Use só para validar o pipeline; para voltar ao normal, remova a linha do .env."
        )

    MAX_TENTATIVAS_POR_MARCA = 5
    ESPERAS_ENTRE_TENTATIVAS = [10, 20, 30, 45]  # segundos: cresce a cada tentativa nova, mantém o último valor se sobrar

    arquivos_baixados_total = []
    marcas_incompletas_final = {}

    with sync_playwright() as p:
        page = None  # nenhum navegador aberto ainda — a primeira tentativa da primeira marca abre um

        for marca_arquivo, marca_bo in MARCAS_CONFIG.items():
            logger.info(LogDivisors.SUB)
            logger.info(f" >>> INICIANDO EXTRAÇÃO PARA A MARCA: {marca_arquivo.upper()} <<<")
            logger.info(LogDivisors.SUB)

            sucesso = False
            ultimo_faltantes = ["motivo desconhecido"]

            # Progresso da MARCA: criado aqui, fora do loop de tentativas, para
            # sobreviver entre elas (ver docstring).
            arquivos_marca = []
            checklist = _novo_checklist()

            for tentativa in range(1, MAX_TENTATIVAS_POR_MARCA + 1):
                page, completo, faltantes, browser_morreu = _extrair_uma_marca_uma_tentativa(
                    p, page, marca_arquivo, marca_bo, data_inicio, data_fim, data_fim_nc,
                    arquivos_marca, checklist
                )

                if completo:
                    arquivos_baixados_total.extend(arquivos_marca)
                    if checklist["Dias_zerados_aceitos"]:
                        logger.warning(
                            f"[MODO TESTE] {marca_arquivo.upper()}: dias aceitos como zerados: "
                            f"{checklist['Dias_zerados_aceitos']}"
                        )
                    logger.info(f"[MARCA COMPLETA] {marca_arquivo.upper()}: todos os arquivos esperados foram obtidos.")
                    sucesso = True
                    break

                ultimo_faltantes = faltantes
                tentativas_restantes = MAX_TENTATIVAS_POR_MARCA - tentativa

                if tentativas_restantes > 0:
                    espera = ESPERAS_ENTRE_TENTATIVAS[min(tentativa - 1, len(ESPERAS_ENTRE_TENTATIVAS) - 1)]
                    motivo_retry = "crash de navegador" if browser_morreu else f"faltando {faltantes}"
                    logger.warning(
                        f"[MARCA RETOMADA] {marca_arquivo.upper()}: tentativa {tentativa}/{MAX_TENTATIVAS_POR_MARCA} "
                        f"incompleta ({motivo_retry}). Aguardando {espera}s e retomando só o que falta..."
                    )
                    time.sleep(espera)

            if not sucesso:
                marcas_incompletas_final[marca_arquivo] = ultimo_faltantes
                logger.error(
                    f"[MARCA FALHOU] {marca_arquivo.upper()}: esgotou as {MAX_TENTATIVAS_POR_MARCA} tentativas "
                    f"(faltando: {ultimo_faltantes}). Esta marca será pulada nas Etapas 2 e 3."
                )

    return arquivos_baixados_total, marcas_incompletas_final


def _isolar_arquivo_invalido(caminho: Path) -> None:
    """
    Tira do caminho esperado um arquivo que falhou na validação de formato,
    renomeando-o para '<nome>.invalido' (fica guardado para diagnóstico).

    Sem isso, o arquivo inválido continuava no lugar do arquivo bom. O UGS
    Diário decide o que baixar pela EXISTÊNCIA do arquivo (ver
    _detectar_lacunas_ugs_diario), então um dia inválido passava a parecer
    "já baixado": a retomada o pularia e o loader leria um arquivo quebrado.
    """
    destino = caminho.with_name(caminho.name + ".invalido")
    try:
        caminho.replace(destino)
        logger.warning(f"Arquivo inválido isolado como '{destino.name}' (guardado para diagnóstico).")
    except OSError:
        logger.exception(f"Não foi possível isolar o arquivo inválido: {caminho}")
        try:
            caminho.unlink(missing_ok=True)
        except OSError:
            logger.exception(f"Também não foi possível remover o arquivo inválido: {caminho}")


def _salvar_e_registrar_download(download_info, caminho_arquivo, arquivos_baixados, atualizar_checklist, rotulo_log="Salvo"):
    """
    Parte final, idêntica em todo download simples do BackOffice: salva o
    arquivo, valida o formato, registra na lista de arquivos baixados,
    atualiza o checklist e loga.

    Se a validação de formato falhar, o arquivo inválido é isolado
    (_isolar_arquivo_invalido) antes de a exceção seguir adiante, para que
    ele não fique no lugar de um arquivo bom.

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
    try:
        validar_formato_xlsx(Path(caminho_arquivo))
    except FormatoInvalidoError:
        _isolar_arquivo_invalido(Path(caminho_arquivo))
        raise
    arquivos_baixados.append(caminho_arquivo)
    atualizar_checklist()
    logger.info(f"{rotulo_log}: {caminho_arquivo}")


def _baixar_nc(page, marca_arquivo, marca_bo, data_inicio, data_fim_nc, pasta_destino, arquivos_baixados, checklist):
    """[1/6] Novas Contas."""
    if checklist["NC"]:
        logger.info(f"[JÁ OBTIDO] NC - {marca_arquivo}: obtido em tentativa anterior, pulando.")
        return

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
    if checklist["Transacoes"]:
        logger.info(f"[JÁ OBTIDO] Transações - {marca_arquivo}: obtido em tentativa anterior, pulando.")
        return

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


def _abrir_tela_ugs(page, marca_bo, data_inicio, data_fim):
    """
    Abre a tela de User Game Statistics já com a marca e o período
    selecionados. Compartilhada por UGS Acumulado e UGS Diário: o diário
    depende dessa tela aberta, mas numa retomada o acumulado pode ter sido
    pulado (já obtido), então o diário precisa conseguir abri-la sozinho.
    """
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


def _baixar_ugs_acumulado(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist):
    """
    [3/6] UGS Acumulado (Completo + 4 tipos).

    Baixa só os tipos ainda não obtidos nesta marca (retomada). Devolve True
    se abriu a tela de UGS (que o UGS Diário aproveita) e False se todos os
    tipos já estavam obtidos e nada foi aberto.
    """
    tipos_ugs = {"": "Completo", "1": "ST", "2": "LC", "7": "SB", "8": "MG"}
    pendentes = {valor: sigla for valor, sigla in tipos_ugs.items() if not checklist[f"UGS_{sigla}"]}

    if not pendentes:
        logger.info(f"[JÁ OBTIDO] UGS acumulado - {marca_arquivo}: todos os tipos obtidos em tentativa anterior, pulando.")
        return False

    if len(pendentes) < len(tipos_ugs):
        ja_obtidos = [sigla for sigla in tipos_ugs.values() if sigla not in pendentes.values()]
        logger.info(f"[JÁ OBTIDO] UGS acumulado - {marca_arquivo}: {ja_obtidos} obtidos em tentativa anterior, pulando.")

    _abrir_tela_ugs(page, marca_bo, data_inicio, data_fim)

    for valor, sigla in pendentes.items():
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

    return True


def _detectar_lacunas_ugs_diario(marca_arquivo, janela_dias=7):
    """
    Verifica os últimos `janela_dias` dias (sem contar hoje) e devolve
    quais ainda não têm um xlsx válido em disco. Checagem pura de
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

        # Entra na lista de download se o arquivo não existe OU existe mas não é um xlsx
        # de verdade (ex: sobra de um download que veio quebrado). Um arquivo inválido
        # não pode contar como "dia já baixado".
        if not arquivo_esperado.exists() or _detectar_formato_real(arquivo_esperado) != "xlsx":
            dias_faltantes.append(dia_checar)

    return dias_faltantes


def _criar_ugs_diario_vazio(destino: Path) -> None:
    """
    Cria um UGS Diário sem nenhuma linha de dados (só o cabeçalho) no lugar de
    um dia que o BackOffice não devolveu válido, para o Step 7 contar 0 usuários.

    O cabeçalho é copiado de um UGS Diário válido já existente: primeiro da
    pasta do próprio mês, depois das outras pastas de mês do mesmo ano. Se não
    houver nenhum, levanta FileNotFoundError — melhor falhar do que inventar colunas.
    """
    mesma_pasta = sorted((f for f in destino.parent.glob("*.xlsx") if f != destino),
                         key=lambda f: f.stat().st_mtime, reverse=True)
    outras_pastas = sorted((f for f in destino.parent.parent.glob("*/*.xlsx") if f.parent != destino.parent),
                           key=lambda f: f.stat().st_mtime, reverse=True)

    cabecalho = None
    for referencia in mesma_pasta + outras_pastas:
        try:
            wb_ref = openpyxl.load_workbook(referencia, read_only=True)
            try:
                primeira_linha = next(wb_ref.active.iter_rows(min_row=1, max_row=1, values_only=True), None)
            finally:
                wb_ref.close()
        except Exception:
            continue  # arquivo ilegível não serve de referência
        if primeira_linha and any(celula is not None for celula in primeira_linha):
            cabecalho = list(primeira_linha)
            break

    if cabecalho is None:
        raise FileNotFoundError(
            f"Nenhum UGS Diário válido encontrado para copiar o cabeçalho (procurado em {destino.parent.parent})."
        )

    novo = openpyxl.Workbook()
    novo.active.append(cabecalho)
    novo.save(destino)
    novo.close()


def _tratar_ugs_diario_invalido(nome_dia, arq_destino, arquivos_baixados, checklist):
    """
    Chamada quando um UGS Diário falha na validação de formato. Devolve True se o
    dia foi aceito como zerado (arquivo só com cabeçalho criado no lugar) e False
    se o chamador deve levantar o erro, como sempre.

    Só aceita com ACEITAR_DIAS_ZERADOS ligado e a partir da 2ª falha do mesmo dia
    (a 1ª pode ser transitória; a retomada baixa o dia de novo antes de desistir).
    """
    falhas = checklist["UGS_Diario_falhas"]
    falhas[nome_dia] = falhas.get(nome_dia, 0) + 1

    if not ACEITAR_DIAS_ZERADOS or falhas[nome_dia] < 2:
        return False

    try:
        _criar_ugs_diario_vazio(Path(arq_destino))
    except Exception:
        logger.exception(f"Não foi possível criar o UGS Diário vazio de {nome_dia}; o erro original será mantido.")
        return False

    arquivos_baixados.append(arq_destino)
    checklist["UGS_Diario_obtidos"] += 1
    checklist["Dias_zerados_aceitos"].append(f"UGS_Diario {nome_dia}")
    logger.warning(
        f"[DIA ZERADO ACEITO] UGS Diário {nome_dia}: veio inválido {falhas[nome_dia]}x; criado arquivo só com "
        f"cabeçalho (0 usuários) porque ACEITAR_DIAS_ZERADOS=true."
    )
    return True


def _baixar_ugs_diario(page, marca_arquivo, marca_bo, data_inicio, data_fim, arquivos_baixados, checklist, tela_ugs_aberta):
    """
    [4/6] UGS Diário (Buscador Dinâmico de Lacunas).

    As lacunas são recalculadas a cada chamada a partir dos arquivos que
    existem em disco, então numa retomada só os dias que ainda faltam são
    baixados. Se a tela de UGS não foi aberta pelo UGS Acumulado desta
    tentativa (tela_ugs_aberta=False), ela é aberta aqui — mas só quando há
    lacuna a baixar, para não navegar à toa.
    """
    # Variável ajustável: Quantos dias no passado o robô deve checar?
    JANELA_DIAS = 7
    logger.info(f"Checando lacunas de UGS Diário nos últimos {JANELA_DIAS} dias...")
    dias_diarios_faltantes = _detectar_lacunas_ugs_diario(marca_arquivo, JANELA_DIAS)

    if not dias_diarios_faltantes:
        logger.info(f" Nenhuma lacuna encontrada! Todos os UGS dos últimos {JANELA_DIAS} dias já estão na pasta.")
    else:
        logger.info(f" Foram encontradas {len(dias_diarios_faltantes)} lacunas. Iniciando download...")

    # 'esperados' vale só para o que falta AGORA, e 'obtidos' recomeça do zero:
    # numa retomada, os dias baixados em tentativas anteriores já estão em disco
    # e não entram mais na conta.
    checklist["UGS_Diario_esperados"] = len(dias_diarios_faltantes)
    checklist["UGS_Diario_obtidos"] = 0

    if not dias_diarios_faltantes:
        return

    if not tela_ugs_aberta:
        _abrir_tela_ugs(page, marca_bo, data_inicio, data_fim)

    page.select_option(Seletores.Filtros.GAME_TYPE, value="")

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

        try:
            _salvar_e_registrar_download(
                download_info, arq_ugs_diario, arquivos_baixados,
                atualizar_checklist=lambda: checklist.__setitem__(
                    "UGS_Diario_obtidos", checklist["UGS_Diario_obtidos"] + 1
                ),
                rotulo_log="Salvo UGS Diário",
            )
        except FormatoInvalidoError:
            if not _tratar_ugs_diario_invalido(nome_dia, arq_ugs_diario, arquivos_baixados, checklist):
                raise


def _baixar_ftd(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist):
    """[5/6] FTD."""
    if checklist["FTD"]:
        logger.info(f"[JÁ OBTIDO] FTD - {marca_arquivo}: obtido em tentativa anterior, pulando.")
        return

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
                if tentativa > 1:
                    logger.info(
                        f"Dia {data_loop}: recuperado na tentativa {tentativa}/{max_tentativas}."
                    )
                break

            tentativas_restantes = max_tentativas - tentativa
            # Só aceita uma RESPOSTA que chegou e parece vazia/zerada. Timeout ou erro de
            # requisição não chegam aqui (caem nos except abaixo) e nunca são aceitos.
            aceitar = ACEITAR_DIAS_ZERADOS and tentativas_restantes == 0 and isinstance(candidato, list)
            if tentativas_restantes > 0:
                acao = "Tentando novamente..."
            elif aceitar:
                acao = "[DIA ZERADO ACEITO] ACEITAR_DIAS_ZERADOS=true: seguindo com o dado como veio."
            else:
                acao = "Desistindo após retry."
            logger.warning(
                f"Dia {data_loop} (tentativa {tentativa}/{max_tentativas}): "
                f"dado suspeito — {motivo_falha}. {acao}"
            )
            if tentativas_restantes > 0:
                page.wait_for_timeout(1500)
            elif aceitar:
                # motivo_falha segue preenchido DE PROPÓSITO: é o sinal, para quem chamou,
                # de que este dia foi aceito apesar de suspeito.
                json_do_dia_valido = candidato

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

    # Numa retomada o mesmo arquivo é regravado (agora mais completo): não duplica na lista.
    if arq_gs not in arquivos_baixados:
        arquivos_baixados.append(arq_gs)
    checklist["GeneralStats_obtidos"] = len(dados_json_mensal)
    logger.info(
        f"Salvo (JSON API): {arq_gs} — {len(dados_json_mensal)}/{dias_fechados} dias obtidos."
    )


def _baixar_general_stats(page, marca_arquivo, marca_bo, pasta_destino, arquivos_baixados, checklist):
    """
    [6/6] General Statistics (Scraping da API Invisível).

    Numa retomada, os dias já obtidos ficam em checklist["GeneralStats_dias"]
    e só os dias que faltam são pedidos de novo. Se o mês já estava completo,
    a função inteira é pulada.
    """
    if checklist["GeneralStats_esperados"] and checklist["GeneralStats_obtidos"] >= checklist["GeneralStats_esperados"]:
        logger.info(f"[JÁ OBTIDO] General Statistics - {marca_arquivo}: obtido em tentativa anterior, pulando.")
        return

    logger.info("Extraindo General Statistics via API (JSON)...")

    # Define a data_alvo buscando a inteligência do date_utils
    data_alvo = obter_data_alvo()

    # Extrai estritamente com base na inteligência da data_alvo
    dias_fechados = data_alvo.day
    mes_alvo = data_alvo.month
    ano_alvo = data_alvo.year

    dias_obtidos = checklist["GeneralStats_dias"]
    checklist["GeneralStats_esperados"] = dias_fechados
    dias_pendentes = [dia for dia in range(1, dias_fechados + 1) if dia not in dias_obtidos]

    if dias_obtidos and dias_pendentes:
        logger.info(
            f"[RETOMADA] General Statistics - {marca_arquivo}: {len(dias_obtidos)} dia(s) já obtidos; "
            f"buscando só {dias_pendentes}."
        )

    # Só navega até a tela se ainda há dia a buscar
    if dias_pendentes:
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

        for dia in dias_pendentes:
            # Substitui o 'hoje.replace' por uma formatação de data cravada
            data_loop = f"{dia:02d}-{mes_alvo:02d}-{ano_alvo}"
            # Limite seguro: dia seguinte às 00:00, para não perder o último minuto do dia.
            data_loop_fim = calcular_limite_seguro(datetime(ano_alvo, mes_alvo, dia)).strftime("%d-%m-%Y %H:%M")
            logger.info(f" -> Extraindo dados do dia {data_loop}...")

            json_do_dia, motivo_falha = _extrair_dia_generalstats(page, data_loop, data_loop_fim)

            if json_do_dia is not None:
                dias_obtidos[dia] = json_do_dia
                if motivo_falha:  # aceito apesar de suspeito (ver _extrair_dia_generalstats)
                    checklist["Dias_zerados_aceitos"].append(f"GeneralStats {data_loop}")
            else:
                logger.error(
                    f"Dia {data_loop}: dado não confiável mesmo após retry ({motivo_falha}). "
                    f"Este dia NÃO será contado como obtido."
                )

            # Espera um pouco antes de ir para o próximo dia para não derrubar a API
            page.wait_for_timeout(1000)

    dados_json_mensal = [{"Dia": dia, "dados": dias_obtidos[dia]} for dia in sorted(dias_obtidos)]
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
    # O UGS Diário reaproveita a tela aberta pelo UGS Acumulado; se o acumulado foi
    # pulado numa retomada, o diário abre a tela sozinho (só se houver lacuna).
    tela_ugs_aberta = _baixar_ugs_acumulado(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist)
    _baixar_ugs_diario(page, marca_arquivo, marca_bo, data_inicio, data_fim, arquivos_baixados, checklist, tela_ugs_aberta)
    _baixar_ftd(page, marca_arquivo, marca_bo, data_inicio, data_fim, pasta_destino, arquivos_baixados, checklist)
    _baixar_general_stats(page, marca_arquivo, marca_bo, pasta_destino, arquivos_baixados, checklist)