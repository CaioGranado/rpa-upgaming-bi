import logging
import time
from datetime import datetime, timezone
from pathlib import Path

from config.settings import MARCAS_CONFIG, LogDivisors
from extractors.web_scraper import extrair_dados_upgaming

# Importa todas as funções de injeção
from loaders.excel_injector import (
    atualizar_base_completa_historica,
    carregar_base_ftd,
    carregar_base_kyc,
    carregar_base_mtd,
    carregar_base_nc,
    carregar_base_performance_step1,
    carregar_base_performance_step2,
    carregar_base_performance_step3,
    carregar_base_performance_step4,
    carregar_base_performance_step5,
    carregar_base_performance_step6,
    carregar_base_performance_step7,
    carregar_base_transacoes,
    carregar_base_ugs,
)
from transformers.data_cleaner import tratar_relatorios_crus
from utils.file_utils import MESES_PT


def setup_logger():
    """
    Configura o log para salvar em arquivo (um por dia, mesma convenção de
    pastas usada em obter_pasta_download_diario) e mostrar no terminal.

    Antes, todo dia gravava no mesmo robo_execucao.log, em modo append,
    indefinidamente — o arquivo nunca parava de crescer. Agora cada dia
    tem seu próprio arquivo, sob logs/{ano}/{mes_num} {mes_nome} {ano_2d}/{dia}.log
    """
    hoje = datetime.now(timezone.utc).astimezone()
    ano_4d = hoje.strftime("%Y")
    ano_2d = hoje.strftime("%y")
    mes_num = hoje.strftime("%m")
    dia_str = hoje.strftime("%d-%m-%y")
    nome_mes, _ = MESES_PT[hoje.month]

    pasta_mes = f"{mes_num} {nome_mes} {ano_2d}"
    pasta_logs = Path("logs") / ano_4d / pasta_mes
    pasta_logs.mkdir(parents=True, exist_ok=True)

    caminho_log = pasta_logs / f"{dia_str}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(caminho_log, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

def main():
    setup_logger()
    logger = logging.getLogger(__name__)
    
    logger.info("INICIANDO RPA DE PIPELINE DE DADOS UPGAMING")
    start_time = time.time()

    try:
        # =====================================================================
        # ETAPA 1: EXTRAÇÃO WEB (Playwright)
        # =====================================================================
        logger.info(LogDivisors.MAIN)
        logger.info(" ETAPA 1: EXTRAÇÃO WEB (PLAYWRIGHT) ")
        logger.info(LogDivisors.MAIN)
        
        # O Web Scraper faz o loop nas 3 marcas, cria as pastas diárias e devolve
        # a lista de arquivos + o sinal explícito de quais marcas ficaram incompletas
        # (com a lista do que faltou em cada uma).
        arquivos_baixados, marcas_incompletas = extrair_dados_upgaming()
        
        if not arquivos_baixados:
            logger.error("Nenhum arquivo foi baixado. Abortando pipeline.")
            return

        # =====================================================================
        # ETAPA 2: TRANSFORMAÇÃO (Limpeza de Dados Crus com Pandas)
        # =====================================================================
        logger.info(LogDivisors.MAIN)
        logger.info(" ETAPA 2: TRANSFORMAÇÃO DOS ARQUIVOS(PANDAS/CALAMINE) ")
        logger.info(LogDivisors.MAIN)
        
        # Aplica a máscara contábil, formata datas e converte textos para números
        arquivos_limpos = tratar_relatorios_crus(arquivos_baixados)

        # =====================================================================
        # ETAPA 3: CARREGAMENTO (Injeção no Servidor G: via Win32COM)
        # =====================================================================
        logger.info(LogDivisors.MAIN)
        logger.info(" ETAPA 3: CARREGAMENTO NAS BASES OFICIAIS ")
        logger.info(LogDivisors.MAIN)
        
        for marca in MARCAS_CONFIG:
            logger.info(LogDivisors.SUB)
            logger.info(f" >>> INICIANDO INJEÇÃO PARA A MARCA: {marca.upper()} <<< ")
            logger.info(LogDivisors.SUB)

            if marca in marcas_incompletas:
                logger.warning(
                    f"[MARCA PULADA] {marca.upper()}: extração incompleta na Etapa 1 "
                    f"(faltando: {marcas_incompletas[marca]}). Injeção ignorada para esta marca."
                )
                continue

            try:
                # 3.1 - Relatórios Históricos e Individuais
                atualizar_base_completa_historica(marca)
                carregar_base_nc(marca, arquivos_limpos)
                carregar_base_ftd(marca, arquivos_limpos)
                carregar_base_transacoes(marca, arquivos_limpos)
                carregar_base_ugs(marca, arquivos_limpos)
                carregar_base_kyc(marca, arquivos_limpos)
                carregar_base_mtd(marca)

                # 3.2 - Base de Performance (Steps 1 ao 7)
                carregar_base_performance_step1(marca)
                carregar_base_performance_step2(marca)
                carregar_base_performance_step3(marca)
                carregar_base_performance_step4(marca)
                carregar_base_performance_step5(marca)
                carregar_base_performance_step6(marca, arquivos_limpos)
                carregar_base_performance_step7(marca)
            except Exception:
                # Isola a falha nesta marca: loga com traceback completo (mostra
                # exatamente qual das 14 funções e qual linha quebrou) e segue
                # para a próxima marca do loop, sem abortar o restante do pipeline.
                logger.exception(
                    f"[INJEÇÃO INTERROMPIDA] {marca.upper()}: erro inesperado durante a "
                    f"injeção. Os passos restantes desta marca foram pulados. As demais "
                    f"marcas seguem normalmente."
                )
            
    except Exception:
        # Removido o f-string com a variável de erro redundante
        logger.exception("ERRO FATAL NO PIPELINE:")
    finally:
        elapsed = (time.time() - start_time) / 60
        logger.info(LogDivisors.MAIN)
        logger.info(f" TEMPO TOTAL DE EXECUÇÃO: {elapsed:.2f} minutos.")
        logger.info(LogDivisors.MAIN)

if __name__ == "__main__":
    main()