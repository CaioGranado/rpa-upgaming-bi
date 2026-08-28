import json
import os
from pathlib import Path

from dotenv import load_dotenv

# Carrega o arquivo .env
load_dotenv()

# --- DADOS SENSÍVEIS (Puxados do .env) ---
URL_SISTEMA = os.getenv("URL_SISTEMA", "https://google.com")

try:
    MARCAS_CONFIG = json.loads(os.getenv("MARCAS_JSON", "{}"))
except json.JSONDecodeError:
    MARCAS_CONFIG = {}

# --- CAMINHOS DE PASTAS ---
PASTA_RAIZ = Path.home() / "Downloads"

# --- CONSTANTES DE SELETORES CSS ---
SEL_FROM_DATE = 'input[name="customFromDate"]:visible'
SEL_TO_DATE = 'input[name="customToDate"]:visible'
SEL_OK_BTN = 'div.btn.filter-by-date[text_key="INTERVALS__OK"]:visible'
SEL_EXPORT_BTN = 'span[text_key="EXPORT_TO_EXCEL"]:visible'
SEL_CONFIRM_EXPORT = 'button.btn.export[text_key="EXPORT"]:visible'
SEL_SEARCH_BRAND_INPUT = 'input[name="searchBrand"]:visible'
SEL_SEARCH_ADD_BTN = 'button[name="add"][text_key="SEARCH"]:visible'
SEL_MENU_USERS = 'a[href="/html/users/users.html"]'
SEL_MENU_TRANSACTIONS = 'span[text_key="SIDEBAR__SYSTEM_TRANSACTIONS"]:visible'
SEL_MENU_UGS = 'span[text_key="SIDEBAR__USER_GAME_STATISTICS"]:visible'
SEL_MENU_REPORT = 'span[text_key="SIDEBAR__REPORT"]:visible'
SEL_BRAND_DROPDOWN = 'select[name="searchBrand"]:visible'
SEL_MORE_FILTER_BTN = 'div[text_key="MORE_FILTER"]:visible'
SEL_SEARCH_BTN_ALT = 'button[name="search"][text_key="SEARCH"]:visible'
SEL_CREATE_DATE_FROM = 'input[name="CreateDateFrom"]:visible'
SEL_CREATE_DATE_TO = 'input[name="CreateDateTo"]:visible'
SEL_MENU_FTD = 'span[text_key="SIDEBAR__FIRST_DEPOSITORS"]:visible'
SEL_MENU_GEN_STATS = 'span[text_key="SIDEBAR__GENERAL_STATISTICS"]:visible'
SEL_SEARCH_FTD_BTN = 'button.btn[type="submit"][text_key="SEARCH"]:visible'
SEL_GAME_TYPE_DROPDOWN = 'select[name="gameType"]:visible'
SEL_BRAND_OK_BTN = 'div.btn:has-text("OK"):visible'

# --- CONSTANTES SIMPLES ---
TXT_ORGANICOS = "Orgânicos"

# --- DIVISORES DE LOG ---
DIVISOR_LOG = "=" * 50
DIVISOR_MENOR = "-" * 50