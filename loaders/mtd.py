"""
Atualização incremental do MTD (Multi Time Deposit), com fórmulas VLOOKUP
dinâmicas apontando para o arquivo de Transações do mês/marca correspondente.
"""
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import blindar_dados
from utils.date_utils import obter_data_alvo
from utils.excel_utils import (
    _fechar_excel_seguro,
    aplicar_filtro_dinamica,
    atualizar_dinamicas,
)
from utils.file_utils import (
    MESES_PT,
    _fazer_backup,
    _obter_caminho_download,
    obter_caminho_base,
)

logger = logging.getLogger(__name__)


def carregar_base_mtd(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO ATUALIZAÇÃO INCREMENTAL: MTD ({marca}) ===")
    data_alvo = obter_data_alvo()
    arquivo_mtd = obter_caminho_base(marca, "MTD", data_alvo)
    arquivo_ftd = _obter_caminho_download(marca, f"FTD - {marca}.xlsx")

    if not arquivo_mtd.exists() or not arquivo_ftd.exists():
        logger.warning("Arquivos para atualização do MTD não encontrados. Pulando etapa.")
        return

    logger.info("Analisando MTD para encontrar o último registro (âncora)...")
    try:
        df_mtd = pd.read_excel(arquivo_mtd, sheet_name="Base", usecols=['UserProfileID'])
        if df_mtd.empty:
            logger.warning("A aba Base do MTD está vazia. Nenhuma âncora encontrada.")
            ultimo_id_bruto = None
        else:
            ultimo_id_bruto = df_mtd['UserProfileID'].dropna().iloc[-1]
            logger.info(f"Último UserProfileID encontrado no MTD: {ultimo_id_bruto}")
    except Exception as e:
        logger.error(f"Erro ao ler a âncora do MTD: {e}")
        return

    logger.info("Lendo e blindando dados novos de FTD...")
    df_ftd = blindar_dados(pd.read_excel(arquivo_ftd))
    coluna_id_ftd = 'UserProfileId' 
    
    id_ancora_str = str(ultimo_id_bruto).replace('.0', '')
    ftd_ids_str = df_ftd[coluna_id_ftd].astype(str).str.replace('.0', '', regex=False)
    
    if ultimo_id_bruto is not None and id_ancora_str in ftd_ids_str.values:
        idx_corte = ftd_ids_str[ftd_ids_str == id_ancora_str].index[0]
        logger.info(f"Correspondência encontrada na linha {idx_corte} do FTD. Pegando registros a partir da linha de baixo...")
        df_novos = df_ftd.iloc[idx_corte + 1:]
    else:
        logger.warning("ID âncora não encontrado no FTD (ou Base MTD vazia). Pegando todos os dados...")
        df_novos = df_ftd

    dados_a_inserir = df_novos.iloc[:, :21].values.tolist()
    total_linhas_novas = len(dados_a_inserir)
    
    if total_linhas_novas == 0:
        logger.info("Não há novos registros no FTD para adicionar ao MTD. Finalizando módulo.")
        return

    wb, excel = None, None
    try:
        _fazer_backup(arquivo_mtd)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False 
        
        logger.info("Abrindo o arquivo oficial de MTD...")
        wb = excel.Workbooks.Open(str(arquivo_mtd), UpdateLinks=3)
        ws = wb.Sheets("Base")
        
        ultima_linha_antiga = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        linha_inicio = ultima_linha_antiga + 1
        linha_fim = linha_inicio + total_linhas_novas - 1
        
        logger.info(f"Injetando {total_linhas_novas} novos registros brutos a partir da linha {linha_inicio}...")
        ws.Range(ws.Cells(linha_inicio, 1), ws.Cells(linha_fim, 21)).Value = dados_a_inserir
        
        if ultima_linha_antiga >= 2:
            logger.info("Escrevendo Fórmulas VLOOKUP Dinâmicas em Lote (Super Rápido)...")
            meses_abrev = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
            abreviacao_marca = {"Betfast": "Bet", "TivoBet": "Tivo", "Faz1Bet": "Faz1"}.get(marca, marca)
            
            # --- INJEÇÃO EM LOTE PARA EVITAR O TRAVAMENTO ---
            formulas_lote = []
            for row in range(linha_inicio, linha_fim + 1):
                data_val = ws.Cells(row, 20).Value  # Coluna T = 20 (TransactionDate)
                if data_val:
                    if isinstance(data_val, str):
                        try: data_obj = datetime.strptime(data_val[:10], "%Y-%m-%d")
                        except ValueError:
                            try: data_obj = datetime.strptime(data_val[:10], "%d/%m/%Y")
                            except Exception: data_obj = datetime.now()
                    else:
                        data_obj = data_val
                        
                    mes_nome = meses_abrev[data_obj.month]
                    ano_full = str(data_obj.year)
                    ano_curto = ano_full[-2:]
                    
                    # Rota padronizada para todas as marcas
                    caminho_base_trans = Path(str(arquivo_mtd)).parent.parent / "Histórico Transações" / ano_full
                    nome_arq = f"{data_obj.month:02d}{mes_nome}_Transações_{abreviacao_marca}_{ano_curto}.xlsx"
                    
                    row_formulas = []
                    for col in range(26, 57):
                        col_idx = 2 if col == 26 else col - 24
                        formula = f"=IFERROR(VLOOKUP($A{row},'{caminho_base_trans}\\[{nome_arq}]MTD_Performance'!$A:$AF,{col_idx},0),0)"
                        row_formulas.append(formula)
                    formulas_lote.append(row_formulas)
                else:
                    formulas_lote.append(["0"] * 31)

            # Injeta todas as milhares de fórmulas de uma vez só!
            ws.Range(ws.Cells(linha_inicio, 26), ws.Cells(linha_fim, 56)).Formula = formulas_lote

            logger.info("Arrastando as fórmulas padrões das colunas V até Y...")
            ws.Range(f"V{ultima_linha_antiga}:Y{linha_fim}").FillDown()
        
        logger.info("Sincronizando Tabelas Dinâmicas da aba 'Din'...")
        atualizar_dinamicas(wb, excel)

        ws_din = wb.Sheets("Din")
        mes_str = MESES_PT[data_alvo.month][0].lower()
        aplicar_filtro_dinamica(ws_din, "B1", "(Tudo)")
        aplicar_filtro_dinamica(ws_din, "B2", mes_str)
        aplicar_filtro_dinamica(ws_din, "A5", valor_desejado="(Tudo)", exceto="(blank)")
        aplicar_filtro_dinamica(ws_din, "B4", valor_desejado="(Tudo)", exceto="(blank)")
        
        logger.info("Salvando arquivo MTD (Mãos longe do teclado!)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de MTD concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de MTD:")
        _fechar_excel_seguro(wb, excel)
        raise