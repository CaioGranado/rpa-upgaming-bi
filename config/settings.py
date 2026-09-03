import json
import logging
import os
import sys
from pathlib import Path
from typing import Final

from dotenv import load_dotenv

# Criando logger específico (__name__) para esse módulo, rastreabilidade de erro
logger = logging.getLogger(__name__)

# Carrega o arquivo .env e força a parada se ele não existir
if not load_dotenv():
    logger.critical("Arquivo .env não encontrado, abortando a execução!")
    sys.exit(1)

# ------------ Dados Sensíveis e Configurações ------------ 
URL_SISTEMA: Final[str] = os.getenv("URL_SISTEMA", "https://google.com")

try:
    _marcas_raw = os.getenv("MARCAS_JSON")
    if not _marcas_raw:
        raise ValueError("A variavél MARCAS_JSON está vazia no .env.")
    MARCAS_CONFIG: Final[dict] = json.loads(_marcas_raw)
except ValueError as e:
    logger.critical(f"Erro fatal ao carregar as configurações das marcas: {e}")
    sys.exit(1)

# ------------ Diretórios e Cosntantes Básicas ------------ 
PASTA_RAIZ: Final[Path] = Path.home() / "Downloads"
TXT_ORGANICOS: Final[str] = "Orgânicos"

_drive_raw = os.getenv("DRIVE_ROOT_PATH")
if not _drive_raw:
    logger.critical("A variável DRIVE_ROOT_PATH não foi encontrada no .env!")
    sys.exit(1)
BASE_PATH: Final[Path] = Path(_drive_raw)

class LogDivisors:
    """Divisores físicos padronizados para os arquivos de log."""
    MAIN: Final[str] = "=" * 50
    SUB: Final[str] = "-" * 50

# ------------ SELETORES CSS - Agrupados por Contexto ------------
class Seletores:
    """Organização hierárquica dos seletores do Playwright."""

    class Menu:
        USERS: Final[str] = 'span[text_key="SIDEBAR__USERS"]:visible'
        TRANSACTIONS: Final[str] = 'span[text_key="SIDEBAR__SYSTEM_TRANSACTIONS"]:visible'
        UGS: Final[str] = 'span[text_key="SIDEBAR__USER_GAME_STATISTICS"]:visible'
        REPORT: Final[str] = 'span[text_key="SIDEBAR__REPORT"]:visible'
        FTD: Final[str] = 'span[text_key="SIDEBAR__FIRST_DEPOSITORS"]:visible'
        GEN_STATS: Final[str] = 'span[text_key="SIDEBAR__GENERAL_STATISTICS"]:visible'

    class Filtros:
        DATE_FROM: Final[str] = 'input[name="customFromDate"]:visible'
        DATE_TO: Final[str] = 'input[name="customToDate"]:visible'
        CREATE_DATE_FROM: Final[str] = 'input[name="CreateDateFrom"]:visible'
        CREATE_DATE_TO: Final[str] = 'input[name="CreateDateTo"]:visible'
        BRAND_DROPDOWN: Final[str] = 'select[name="searchBrand"]:visible'
        SEARCH_BRAND_INPUT: Final[str] = 'input[name="searchBrand"]:visible'
        GAME_TYPE: Final[str] = 'select[name="gameType"]:visible'

    class Botoes:
        OK: Final[str] = 'div.btn.filter-by-date[text_key="INTERVALS__OK"]:visible'
        BRAND_OK: Final[str] = 'div.btn:has-text("OK"):visible'
        EXPORT: Final[str] = 'span[text_key="EXPORT_TO_EXCEL"]:visible'
        CONFIRM_EXPORT: Final[str] = 'button.btn.export[text_key="EXPORT"]:visible'
        SEARCH_ADD: Final[str] = 'button[name="add"][text_key="SEARCH"]:visible'
        SEARCH_ALT: Final[str] = 'button[name="search"][text_key="SEARCH"]:visible'
        SEARCH_FTD: Final[str] = 'button.btn[type="submit"][text_key="SEARCH"]:visible'
        MORE_FILTER: Final[str] = 'div[text_key="MORE_FILTER"]:visible'
