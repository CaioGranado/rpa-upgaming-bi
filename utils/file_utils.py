import json
import logging
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path

from config.settings import BASE_PATH
from utils.date_utils import obter_data_alvo
from utils.exceptions_utils import FormatoInvalidoError

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

def _obter_caminho_download(marca: str, nome_arquivo: str) -> Path:
    pasta = obter_pasta_download_diario(marca)
    caminho = pasta / nome_arquivo
    if not caminho.exists():
        logger.warning(f"Arquivo não encontrado no download diário: {caminho}")
    return caminho

def _fazer_backup(caminho_arquivo: Path):
    if not caminho_arquivo.exists(): return
    
    data_referencia = obter_data_alvo()
    ano = data_referencia.strftime("%Y")
    mes_numero = data_referencia.strftime("%m")

    mes_nome = MESES_PT[data_referencia.month][0]
    pasta_mes = f"{mes_numero} - {mes_nome}"

    pasta_bkp = caminho_arquivo.parent / "Backups" / ano / pasta_mes
    pasta_bkp.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now(timezone.utc).astimezone().strftime("%Y%m%d_%H%M%S")
    nome_bkp = f"{caminho_arquivo.stem}_BKP_{timestamp}{caminho_arquivo.suffix}"

    shutil.copy2(caminho_arquivo, pasta_bkp / nome_bkp)
    logger.info(f" § Backup salvo em [{ano}/{pasta_mes}]: {nome_bkp}")

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

# =============================================================================
# 5. VALIDAÇÃO DE FORMATO DE ARQUIVO PÓS-DOWNLOAD
# =============================================================================
def _detectar_formato_real(caminho: Path) -> str:
    """
    Detecta o formato real do arquivo pelos primeiros bytes (assinatura/magic
    number), independente da extensão do nome — que pode estar errada, como
    vimos no caso do Export de Transações vindo em JSON com extensão .xlsx.

    Retorna "xlsx" (assinatura ZIP, 'PK'), "json" (começa com '{' ou '['),
    ou "desconhecido" (nenhum dos dois — ex: HTML de erro, arquivo vazio).

    LIMITAÇÃO CONHECIDA (aceita deliberadamente em 16/09/2026): não reconhece
    .xls legado (formato OLE2, assinatura D0 CF 11 E0) — seria classificado
    como "desconhecido" e rejeitado. Não implementado porque o fluxo atual de
    download sempre salva com extensão/conteúdo .xlsx; revisar se algum dia
    isso mudar.
    """
    try:
        with open(caminho, 'rb') as f:
            inicio = f.read(4)
    except Exception:
        return "desconhecido"

    if inicio[:2] == b'PK':
        return "xlsx"

    inicio_sem_espacos = inicio.lstrip()
    if inicio_sem_espacos[:1] in (b'{', b'['):
        return "json"

    return "desconhecido"


def _diagnosticar_json_inesperado(caminho: Path, nome_arquivo: str) -> str:
    """
    Tenta extrair um diagnóstico específico quando o arquivo é JSON — caso
    mais comum: endpoint de export devolvendo resposta de API paginada
    (campo 'count' indica o total real de registros) em vez do xlsx completo.

    Retorna a mensagem de diagnóstico formatada para uso em FormatoInvalidoError.
    """
    try:
        with open(caminho, 'r', encoding='utf-8') as f:
            dados = json.load(f)

        if isinstance(dados, list) and dados and isinstance(dados[0], dict):
            qtd_recebida = len(dados)
            total_esperado = dados[0].get('count')

            if isinstance(total_esperado, int) and total_esperado > qtd_recebida:
                return (
                    f"{nome_arquivo}: Export voltou em JSON paginado (recebidos "
                    f"{qtd_recebida} de {total_esperado} registros esperados pelo campo "
                    f"'count'). Possível regressão no endpoint do BackOffice — o export "
                    f"está devolvendo a resposta de API usada para paginação on-screen "
                    f"em vez do relatório completo. Contatar time responsável pelo BO."
                )

        return (
            f"{nome_arquivo}: Export voltou em JSON em vez de xlsx, "
            f"mas em estrutura inesperada (não é uma lista de registros com 'count'). "
            f"Verificar manualmente."
        )

    except Exception:
        return (
            f"{nome_arquivo}: arquivo não veio em xlsx e também não pôde ser "
            f"interpretado como JSON válido. Verificar manualmente."
        )


def validar_formato_xlsx(caminho: Path) -> None:
    """
    Valida que o arquivo salvo é de fato um xlsx (assinatura ZIP/PK).
    Deve ser chamada logo após cada download_info.value.save_as(), antes
    de appendar o caminho em arquivos_baixados.

    Levanta FormatoInvalidoError com diagnóstico específico se o formato
    não for xlsx — interrompendo o pipeline da marca afetada antes de
    tentar tratar ou injetar dados incompletos/incorretos.

    Não deve ser usada para o General Statistics, que é capturado
    intencionalmente como JSON via interceptação de rede.
    """
    nome_arquivo = caminho.name
    formato = _detectar_formato_real(caminho)

    if formato == "xlsx":
        logger.info(f"[FORMATO OK] {nome_arquivo}: xlsx confirmado por assinatura de bytes.")
        return

    if formato == "json":
        diagnostico = _diagnosticar_json_inesperado(caminho, nome_arquivo)
        logger.error(f"[FORMATO INVÁLIDO] {diagnostico}")
        raise FormatoInvalidoError(diagnostico)

    # Formato desconhecido (HTML de erro, arquivo vazio, etc.)
    msg = (
        f"{nome_arquivo}: arquivo não reconhecido como xlsx nem como JSON "
        f"(assinatura de bytes não corresponde a nenhum dos dois). "
        f"Verificar manualmente."
    )
    logger.error(f"[FORMATO INVÁLIDO] {msg}")
    raise FormatoInvalidoError(msg)