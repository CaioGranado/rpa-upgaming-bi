import logging
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path

from config.settings import BASE_PATH

logger = logging.getLogger(__name__)

# =============================================================================
# 1. DICIONÁRIOS DE MAPEAMENTO
# =============================================================================
MESES_PT = {
    1: ('Janeiro', 'Jan'), 2: ('Fevereiro', 'Fev'), 3: ('Março', 'Mar'),
    4: ('Abril', 'Abr'), 5: ('Maio', 'Mai'), 6: ('Junho', 'Jun'),
    7: ('Julho', 'Jul'), 8: ('Agosto', 'Ago'), 9: ('Setembro', 'Set'),
    10: ('Outubro', 'Out'), 11: ('Novembro', 'Nov'), 12: ('Dezembro', 'Dez')
}

MARCAS_PATHS = {
    "Betfast": {
        "base": BASE_PATH  / "Betfast",
        "abv": "Bet", "kyc": "Betfast", "ugs_abv": "Bet"
    },
    "Faz1Bet": {
        "base": BASE_PATH  / "Faz1Bet",
        "abv": "Faz1", "kyc": "Faz1Bet", "ugs_abv": "F1B"
    },
    "TivoBet": {
        "base": BASE_PATH  / "TivoBet",
        "abv": "Tivo", "kyc": "Tivobet", "ugs_abv": "Tivo"
    }
}

# =============================================================================
# 2. FUNÇÃO AUXILIAR: CLONAR ARQUIVO ANTERIOR (TEMPLATE)
# =============================================================================
def garantir_arquivo_existente(caminho_alvo: Path, pasta_busca: Path):
    if caminho_alvo.exists():
        return caminho_alvo

    logger.warning(f"Arquivo não encontrado: {caminho_alvo.name}. Buscando template para cópia...")
    
    arquivos_disponiveis = list(pasta_busca.glob("*.xls*"))
    
    if not arquivos_disponiveis:
        ano_atual = pasta_busca.name
        if ano_atual.isdigit():
            pasta_ano_anterior = pasta_busca.parent / str(int(ano_atual) - 1)
            if pasta_ano_anterior.exists():
                arquivos_disponiveis = list(pasta_ano_anterior.glob("*.xls*"))
    
    if arquivos_disponiveis:
        arquivo_base = max(arquivos_disponiveis, key=os.path.getmtime)
        shutil.copy2(arquivo_base, caminho_alvo)
        logger.info(f"✅ Arquivo criado por cópia: {arquivo_base.name} -> {caminho_alvo.name}")
    else:
        logger.error(f"❌ Não foi possível criar {caminho_alvo.name}: Nenhum template anterior encontrado!")
        
    return caminho_alvo

# =============================================================================
# 3. ROTAS DE DOWNLOADS, UGS E BASE COMPLETA
# =============================================================================
def obter_pasta_download_diario(marca: str) -> Path:
    config = MARCAS_PATHS[marca]
    hoje = datetime.now(timezone.utc).astimezone()
    
    ano_4d = hoje.strftime("%Y")
    ano_2d = hoje.strftime("%y")
    mes_num = hoje.strftime("%m")
    dia_str = hoje.strftime("%d-%m-%y")
    nome_mes, _ = MESES_PT[hoje.month]
    
    pasta_mes = f"{mes_num} {nome_mes} {ano_2d}"
    caminho = config["base"] / "Arquivos de Download Diários" / ano_4d / pasta_mes / dia_str
    
    if not caminho.exists():
        caminho.mkdir(parents=True, exist_ok=True)
        
    return caminho

def obter_pasta_ugs_diario(marca: str, ano: int, mes: int) -> Path:
    nome_pasta_mes = f"{mes:02d} - {MESES_PT[mes][0]}"
    
    # Busca a chave "base" do dicionário corretamente
    pasta_base = MARCAS_PATHS[marca]["base"]
    
    pasta_ugs = pasta_base / "Histórico User Game Statistics" / "Diário" / str(ano) / nome_pasta_mes
    pasta_ugs.mkdir(parents=True, exist_ok=True)
    
    return pasta_ugs

def obter_caminho_base_completa(marca: str) -> Path:
    marcas_corretas = {"BetFast": "Betfast", "Faz1Bet": "Faz1Bet", "TivoBet": "TivoBet"}
    nome_oficial = marcas_corretas.get(marca, marca)
    pasta = BASE_PATH / "xFAST" / "Bases Completas"

    return pasta / f"{nome_oficial} - Base Completa.xlsx"

# =============================================================================
# 4. ROTAS DAS BASES FINAIS (PARA ATUALIZAÇÃO)
# =============================================================================
def obter_caminho_base(marca: str, relatorio: str, data_alvo: datetime) -> Path:
    config = MARCAS_PATHS[marca]
    base_path = config["base"]
    abv = config["abv"]
    
    ano_4d = data_alvo.strftime("%Y")
    ano_2d = data_alvo.strftime("%y")
    mes_num = data_alvo.strftime("%m")
    nome_mes_completo, mes_abv = MESES_PT[data_alvo.month]
    
    pasta_destino = None
    nome_arquivo = None

    if relatorio == "NC":
        pasta_destino = base_path / "Histórico Novas Contas" / ano_4d
        nome_arquivo = f"{mes_num}{mes_abv}_NovasContas_{abv}_{ano_2d}.xlsx"
        
    elif relatorio == "FTD":
        pasta_destino = base_path / "Histórico Primeiros Depositantes" / ano_4d
        nome_arquivo = f"{mes_num}{mes_abv}_FTD_{abv}_{ano_2d}.xlsx"
        
    elif relatorio == "Transacoes":
        pasta_destino = base_path / "Histórico Transações" / ano_4d
        nome_arquivo = f"{mes_num}{mes_abv}_Transações_{abv}_{ano_2d}.xlsx"
        
    elif relatorio == "UGS":
        pasta_destino = base_path / "Histórico User Game Statistics" / "Completo" / ano_4d
        if marca == "Faz1Bet":
            nome_arquivo = f"{config['ugs_abv']}_{mes_num}_UserPlay_{nome_mes_completo}{ano_2d}.xlsx"
        else:
            nome_arquivo = f"{mes_num}{mes_abv}_UserPlay_{config['ugs_abv']}_{ano_2d}.xlsx"
            
    elif relatorio == "KYC":
        pasta_destino = base_path / "Relatório KYC" / ano_4d
        nome_arquivo = f"{mes_num}{mes_abv}_KYC_{config['kyc']}_{ano_2d}.xlsx"
        
    elif relatorio == "MTD":
        pasta_destino = base_path / "Relatório MTD"
        nome_arquivo = f"{config['kyc']} - MTD {ano_4d}.xlsx"
        
    elif relatorio == "Performance":
        pasta_destino = base_path / "Relatório Performance"
        nome_arquivo = f"{config['kyc']}_BasePerformance.xlsm"
    else:
        raise ValueError(f"Relatório '{relatorio}' desconhecido.")

    if not pasta_destino.exists():
        pasta_destino.mkdir(parents=True, exist_ok=True)
        logger.info(f"📁 Pasta de Base criada: {pasta_destino}")

    caminho_final = pasta_destino / nome_arquivo
    
    if relatorio != "Performance":
        caminho_final = garantir_arquivo_existente(caminho_final, pasta_destino)

    return caminho_final