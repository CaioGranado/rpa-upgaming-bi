"""
Atualização incremental da Base Completa histórica (todos os usuários já
vistos de uma marca, usados como referência em Transações via BaseCompleta).
"""
import logging

import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import blindar_dados
from utils.excel_utils import _fechar_excel_seguro
from utils.file_utils import _fazer_backup, _obter_caminho_download, obter_caminho_base_completa

logger = logging.getLogger(__name__)


def atualizar_base_completa_historica(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO ATUALIZAÇÃO INCREMENTAL: BASE COMPLETA ({marca}) ===")
    arquivo_base_completa = obter_caminho_base_completa(marca)
    arquivo_nc = _obter_caminho_download(marca, f"NC - {marca}.xlsx")

    if not arquivo_base_completa.exists() or not arquivo_nc.exists():
        logger.warning("Arquivos para atualização da Base Completa não encontrados. Pulando etapa.")
        return

    logger.info("Analisando a Base Completa para encontrar o último registro (Âncora)...")
    df_base = pd.read_excel(arquivo_base_completa, usecols=['UserProfileID'])
    
    if df_base.empty:
        logger.warning("A Base Completa está vazia. Não há referência para incremental.")
        return
        
    ultimo_id_bruto = df_base['UserProfileID'].dropna().iloc[-1]
    
    id_ancora_str = str(ultimo_id_bruto).replace('.0', '')
    logger.info(f"Último UserProfileID encontrado: {id_ancora_str}")

    logger.info("Lendo dados novos e blindando (NC)...")
    df_nc = blindar_dados(pd.read_excel(arquivo_nc))
    nc_ids_str = df_nc['UserProfileID'].astype(str).str.replace('.0', '', regex=False)
    
    if 'ParentUserName' in df_nc.columns:
        logger.info("Limpando texto 'Orgânicos' da coluna ParentUserName para injeção...")
        df_nc['ParentUserName'] = df_nc['ParentUserName'].replace("Orgânicos", "")

    if id_ancora_str in nc_ids_str.values:
        idx_corte = nc_ids_str[nc_ids_str == id_ancora_str].index[0]
        logger.info(f"Correspondência encontrada! Fatiando novos dados a partir da linha {idx_corte} do arquivo novo...")
        df_novos = df_nc.iloc[idx_corte:]
        achou_ancora = True
    else:
        marcas_arquivo_novo = df_nc['BrandName'].dropna().unique()

        if len(marcas_arquivo_novo) == 1 and marcas_arquivo_novo[0].strip().upper() == marca.strip().upper():
            logger.warning(f"Âncora não encontrada, mas o arquivo pertence exclusivamente à marca {marca}. Injetando carga total por segurança.")
            df_novos = df_nc
            achou_ancora = False
        else:
            logger.error(f"FALHA CRÍTICA: Arquivo NC da marca {marca} contém dados inválidos ou misturados: {marcas_arquivo_novo}. Abortando a injeção de novos dados!")
            return
        
    dados_a_inserir = df_novos.values.tolist()
    total_linhas_novas = len(dados_a_inserir)
    total_colunas = len(df_novos.columns)
    
    if total_linhas_novas == 0:
        logger.info("Não há novos registros para adicionar na Base Completa.")
        return

    wb, excel = None, None
    try:
        _fazer_backup(arquivo_base_completa)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info("Abrindo a Base Completa Histórica para Injeção...")
        wb = excel.Workbooks.Open(str(arquivo_base_completa))
        ws = wb.Sheets(1)
        
        ultima_linha_antiga = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        linha_inicio = ultima_linha_antiga if achou_ancora else (ultima_linha_antiga + 1)
        linha_fim = linha_inicio + total_linhas_novas - 1
        
        logger.info(f"Injetando {total_linhas_novas} registros a partir da linha {linha_inicio}...")
        ws.Range(ws.Cells(linha_inicio, 1), ws.Cells(linha_fim, total_colunas)).Value = dados_a_inserir
        
        logger.info("Salvando Base Completa (Pode demorar, arquivo pesado!)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Base Completa atualizada com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico ao atualizar Base Completa:")
        _fechar_excel_seguro(wb, excel)