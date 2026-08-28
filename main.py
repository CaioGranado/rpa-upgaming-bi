import logging
import time

from dotenv import load_dotenv

# Carrega as senhas e variáveis ocultas
load_dotenv()

from config.settings import MARCAS_CONFIG
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


def setup_logger():
    """Configura o log para salvar em arquivo e mostrar no terminal."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("robo_execucao.log", encoding='utf-8'),
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
        logger.info("\n" + "="*70 + "\nETAPA 1: EXTRAÇÃO WEB (PLAYWRIGHT)\n" + "="*70)
        
        # O Web Scraper faz o loop nas 3 marcas, cria as pastas diárias e devolve a lista de arquivos
        arquivos_baixados = extrair_dados_upgaming()
        
        if not arquivos_baixados:
            logger.error("Nenhum arquivo foi baixado. Abortando pipeline.")
            return

        # =====================================================================
        # ETAPA 2: TRANSFORMAÇÃO (Limpeza de Dados Crus com Pandas)
        # =====================================================================
        logger.info("\n" + "="*70 + "\nETAPA 2: TRANSFORMAÇÃO DOS ARQUIVOS(PANDAS/CALAMINE)\n" + "="*70)
        
        # Aplica a máscara contábil, formata datas e converte textos para números
        arquivos_limpos = tratar_relatorios_crus(arquivos_baixados)

        # =====================================================================
        # ETAPA 3: CARREGAMENTO (Injeção no Servidor G: via Win32COM)
        # =====================================================================
        logger.info("\n" + "="*70 + "\nETAPA 3: CARREGAMENTO NAS BASES OFICIAIS\n" + "="*70)
        
        # Loop pythônico limpo (sem o .keys())
        for marca in MARCAS_CONFIG:
            logger.info(f"\n>>> INICIANDO INJEÇÃO PARA A MARCA: {marca.upper()} <<<")
            
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
        # Removido o f-string com a variável de erro redundante
        logger.exception("ERRO FATAL NO PIPELINE:")
    finally:
        elapsed = (time.time() - start_time) / 60
        logger.info(f"\nTEMPO TOTAL DE EXECUÇÃO: {elapsed:.2f} minutos.")

if __name__ == "__main__":
    main()