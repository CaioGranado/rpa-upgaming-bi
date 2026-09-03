import calendar
import json
import logging
import shutil
from datetime import datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd
import pywintypes
import win32com.client as win32

# Importações da nossa arquitetura
from transformers.data_cleaner import (
    aplicar_corte_datas_futuras,
    blindar_dados,
    garantir_continuidade_temporal,
)
from utils.date_utils import obter_data_alvo
from utils.file_utils import (
    MESES_PT,
    obter_caminho_base,
    obter_caminho_base_completa,
    obter_pasta_download_diario,
    obter_pasta_ugs_diario,
)

logger = logging.getLogger(__name__)

# ==============================================================================
# FUNÇÕES AUXILIARES DE EXCEL E SEGURANÇA
# ==============================================================================
def formatar_brl(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_int(valor):
    return f"{int(valor):,}".replace(",", ".")

def formatar_num(valor):
    if pd.isna(valor): return "0"
    v = float(valor)
    if v.is_integer():
        return f"{int(v):,}".replace(",", ".")
    else:
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def atualizar_dinamicas(wb):
    logger.info("Sincronizando Tabelas Dinâmicas (RefreshAll)...")
    for cache in wb.PivotCaches():
        try: 
            cache.BackgroundQuery = False
        except pywintypes.com_error as e_cache: 
            logger.debug(f"Propriedade BackgroundQuery ignorada neste cache: {e_cache}")
    wb.RefreshAll()

def _fechar_excel_seguro(wb, excel):
    try:
        if wb: wb.Close(SaveChanges=False)
        if excel: excel.Quit()
    except pywintypes.com_error as e_quit:
        logger.debug(f"Erro na interface COM ao forçar fechamento seguro: {e_quit}")
    except AttributeError as e_attr:
        logger.debug(f"Objeto inexistente durante o fechamento seguro: {e_attr}")

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

def _obter_caminho_download(marca: str, nome_arquivo: str) -> Path:
    pasta = obter_pasta_download_diario(marca)
    caminho = pasta / nome_arquivo
    if not caminho.exists():
        logger.warning(f"Arquivo não encontrado no download diário: {caminho}")
    return caminho

# ==============================================================================
# MÓDULOS DE HISTÓRICO E CARGA
# ==============================================================================
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

def carregar_base_nc(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO CARREGAMENTO: NOVAS CONTAS (NC) ({marca}) ===")
    arquivo_origem = _obter_caminho_download(marca, f"NC - {marca}.xlsx")

    data_alvo = obter_data_alvo()
    arquivo_base_oficial = obter_caminho_base(marca, "NC", data_alvo)
    
    if not arquivo_origem.exists() or not arquivo_base_oficial.exists():
        logger.warning("Arquivos de NC não encontrados. Pulando etapa.")
        return

    logger.info("Lendo dados de NC...")
    df_final = pd.read_excel(arquivo_origem)
    
    if 'RegistrationDate' in df_final.columns:
        logger.info("Aplicando corte dinâmico (NC)")
        df_final = aplicar_corte_datas_futuras(df_final, 'RegistrationDate', data_alvo)
        logger.info(f"Restaram {len(df_final)} registros após o corte.")
    
    df_final = blindar_dados(df_final)
    dados_a_inserir = df_final.iloc[:, :35].values.tolist()
    ultima_linha_destino = 1 + len(dados_a_inserir)
    
    if 'Mobile' in df_final.columns:
        ddd_series = df_final['Mobile'].astype(str).str.replace(r'\D', '', regex=True).str[:2]
    else:
        ddd_series = pd.Series([""] * len(dados_a_inserir))
    dados_ddd = [[val] for val in ddd_series.tolist()]

    wb, excel = None, None
    try:
        _fazer_backup(arquivo_base_oficial)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(str(arquivo_base_oficial))
        ws = wb.Sheets("vw_CustomerProfileFull")
        
        ultima_linha_antiga = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        if ultima_linha_antiga >= 3:
            ws.Range(f"A3:AR{ultima_linha_antiga}").ClearContents()
            
        ws.Range("A2:AI2").ClearContents()
        if hasattr(ws, "Range"): ws.Range("AP2:AP2").ClearContents()

        ws.Range(ws.Cells(2, 1), ws.Cells(ultima_linha_destino, 35)).Value = dados_a_inserir
        ws.Range(ws.Cells(2, 42), ws.Cells(ultima_linha_destino, 42)).Value = dados_ddd
        
        if ultima_linha_destino > 2:
            ws.Range(f"AJ2:AO{ultima_linha_destino}").FillDown()
            ws.Range(f"AQ2:AR{ultima_linha_destino}").FillDown()
        
        atualizar_dinamicas(wb)
        excel.CalculateUntilAsyncQueriesDone()
        
        logger.info("Salvando arquivo de NC (Por favor, aguarde. Pode levar alguns segundos)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de NC concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de NC:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_ftd(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO CARREGAMENTO: FTD ({marca}) ===")
    arquivo_origem = _obter_caminho_download(marca, f"FTD - {marca}.xlsx")
    arquivo_base_oficial = obter_caminho_base(marca, "FTD", obter_data_alvo())
    
    if not arquivo_origem.exists() or not arquivo_base_oficial.exists():
        logger.warning("Arquivos de FTD não encontrados. Pulando etapa.")
        return

    logger.info("Lendo e blindando dados de FTD...")
    df_final = blindar_dados(pd.read_excel(arquivo_origem))
    
    dados_a_inserir = df_final.iloc[:, :21].values.tolist()
    ultima_linha_destino = 1 + len(dados_a_inserir)

    wb, excel = None, None
    try:
        _fazer_backup(arquivo_base_oficial)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(str(arquivo_base_oficial))
        ws = wb.Sheets("Sheet 1")
        
        ultima_linha_antiga = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        if ultima_linha_antiga >= 3:
            ws.Range(f"A3:AD{ultima_linha_antiga}").ClearContents()
        ws.Range("A2:U2").ClearContents()

        ws.Range(ws.Cells(2, 1), ws.Cells(ultima_linha_destino, 21)).Value = dados_a_inserir
        
        if ultima_linha_destino > 2:
            ws.Range(f"V2:AD{ultima_linha_destino}").FillDown()
        
        atualizar_dinamicas(wb)
        excel.CalculateUntilAsyncQueriesDone()
        
        logger.info("Salvando arquivo de FTD (Por favor, aguarde)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de FTD concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de FTD:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_transacoes(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO CARREGAMENTO: TRANSAÇÕES E BASE COMPLETA ({marca}) ===")
    arquivo_origem_base = obter_caminho_base_completa(marca)
    arquivo_origem_transacoes = _obter_caminho_download(marca, f"Transações - {marca}.xlsx")
    arquivo_base_oficial = obter_caminho_base(marca, "Transacoes", obter_data_alvo())
    
    if not arquivo_base_oficial.exists():
        logger.warning("Arquivo oficial de Transações não encontrado.")
        return

    if arquivo_origem_base.exists():
        logger.info(f"Lendo dados consolidados de: {arquivo_origem_base.name} para copiar 3 colunas...")
        df_base = blindar_dados(pd.read_excel(arquivo_origem_base))
        dados_base = df_base[['UserProfileID', 'Pin', 'ParentUserName']].values.tolist()
        ultima_linha_base = 1 + len(dados_base)
    else:
        logger.error(f"Arquivo Base Completa não encontrado: {arquivo_origem_base.name}")
        return

    if arquivo_origem_transacoes.exists():
        logger.info(f"Lendo dados de: {arquivo_origem_transacoes.name}...")
        df_transacoes = blindar_dados(pd.read_excel(arquivo_origem_transacoes))
        dados_transacoes = df_transacoes.iloc[:, :28].values.tolist()
        ultima_linha_transacoes = 1 + len(dados_transacoes)
    else:
        logger.error(f"Arquivo de Transações não encontrado: {arquivo_origem_transacoes.name}")
        return

    wb, excel = None, None
    try:
        _fazer_backup(arquivo_base_oficial)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(str(arquivo_base_oficial))
        
        logger.info("Alimentando Aba BaseCompleta...")
        ws_base = wb.Sheets("BaseCompleta")
        ultima_linha_antiga_base = ws_base.Cells(ws_base.Rows.Count, 1).End(-4162).Row
        if ultima_linha_antiga_base >= 2:
            ws_base.Range(f"A2:C{ultima_linha_antiga_base}").ClearContents()
        ws_base.Range(ws_base.Cells(2, 1), ws_base.Cells(ultima_linha_base, 3)).Value = dados_base
        
        logger.info("Alimentando Aba Sheet 1 (Transações)...")
        ws_sheet1 = wb.Sheets("Sheet 1")
        ultima_linha_antiga_transacoes = ws_sheet1.Cells(ws_sheet1.Rows.Count, 1).End(-4162).Row
        
        if ultima_linha_antiga_transacoes >= 3:
            ws_sheet1.Range(f"A3:AJ{ultima_linha_antiga_transacoes}").ClearContents()
        ws_sheet1.Range("A2:AB2").ClearContents()
            
        ws_sheet1.Range(ws_sheet1.Cells(2, 1), ws_sheet1.Cells(ultima_linha_transacoes, 28)).Value = dados_transacoes
        
        if ultima_linha_transacoes > 2:
            ws_sheet1.Range(f"AC2:AJ{ultima_linha_transacoes}").FillDown()
        
        atualizar_dinamicas(wb)
        excel.CalculateUntilAsyncQueriesDone()
        
        logger.info("Salvando arquivo de Transações (Isto pode demorar. Mãos longe do teclado!)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de Transações concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de Transações:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_ugs(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO CARREGAMENTO: UGS ({marca}) ===")
    arquivo_base_oficial = obter_caminho_base(marca, "UGS", obter_data_alvo())
    
    if not arquivo_base_oficial.exists():
        logger.warning("Arquivo oficial de UGS não encontrado.")
        return

    mapeamento_ugs = {
        "Completo": f"{marca} - UGS Completo.xlsx",
        "Slot": f"{marca} - UGS ST.xlsx",
        "LiveCassino": f"{marca} - UGS LC.xlsx",
        "Sportsbook": f"{marca} - UGS SB.xlsx",
        "MiniGames": f"{marca} - UGS MG.xlsx"
    }

    dados_para_injetar = {}
    
    for aba, nome_arquivo in mapeamento_ugs.items():
        caminho_origem = _obter_caminho_download(marca, nome_arquivo)
        if caminho_origem.exists():
            df_temp = blindar_dados(pd.read_excel(caminho_origem))
            dados = df_temp.values.tolist()
            dados_para_injetar[aba] = {
                "dados": dados,
                "total_cols": len(df_temp.columns),
                "total_linhas": len(dados)
            }
        else:
            logger.error(f"Arquivo não encontrado: {nome_arquivo}.")

    if not dados_para_injetar:
        return

    wb, excel = None, None
    try:
        _fazer_backup(arquivo_base_oficial)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(str(arquivo_base_oficial))
        
        for aba, info in dados_para_injetar.items():
            logger.info(f"Injetando {info['total_linhas']} registros na aba '{aba}'...")
            try:
                ws = wb.Sheets(aba)
            except Exception:
                logger.error(f"Aba '{aba}' não encontrada no arquivo.")
                continue
                
            ultima_linha_antiga = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            if ultima_linha_antiga >= 2:
                ws.Range(ws.Cells(2, 1), ws.Cells(ultima_linha_antiga, info["total_cols"])).ClearContents()

            ultima_linha_destino = 1 + info["total_linhas"]
            if info["total_linhas"] > 0:
                ws.Range(ws.Cells(2, 1), ws.Cells(ultima_linha_destino, info["total_cols"])).Value = info["dados"]

        atualizar_dinamicas(wb)
        excel.CalculateUntilAsyncQueriesDone()
        
        logger.info("Salvando arquivo de UGS (Paciência, quase lá!)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de UGS concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de UGS:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_kyc(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO CARREGAMENTO: KYC ({marca}) ===")
    arquivo_nc = _obter_caminho_download(marca, f"NC - {marca}.xlsx")
    arquivo_ftd = _obter_caminho_download(marca, f"FTD - {marca}.xlsx")
    arquivo_base_oficial = obter_caminho_base(marca, "KYC", obter_data_alvo())
    
    if not arquivo_base_oficial.exists():
        logger.warning("Arquivo oficial de KYC não encontrado. Pulando etapa.")
        return

    if arquivo_nc.exists():
        logger.info(f"Lendo dados de: {arquivo_nc.name}...")
        df_nc = pd.read_excel(arquivo_nc)
        
        if 'RegistrationDate' in df_nc.columns:
            logger.info("Aplicando corte rigoroso (NC): Mantendo registros até Hoje às 00:00...")
            data_alvo = obter_data_alvo()
            df_nc = aplicar_corte_datas_futuras(df_nc, 'RegistrationDate', data_alvo)
            logger.info(f"Restaram {len(df_nc)} registros após o corte.")
            
        df_nc = blindar_dados(df_nc)
        dados_nc = df_nc.iloc[:, :35].values.tolist()
        ultima_linha_nc = 1 + len(dados_nc)
    else:
        logger.error(f"Arquivo não encontrado para a aba NC: {arquivo_nc.name}")
        return

    if arquivo_ftd.exists():
        logger.info(f"Lendo e blindando dados de: {arquivo_ftd.name}...")
        df_ftd = blindar_dados(pd.read_excel(arquivo_ftd))
        dados_ftd = df_ftd.iloc[:, :21].values.tolist()
        ultima_linha_ftd = 1 + len(dados_ftd)
    else:
        logger.error(f"Arquivo não encontrado para a aba FTD: {arquivo_ftd.name}")
        return

    wb, excel = None, None
    try:
        _fazer_backup(arquivo_base_oficial)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info("Abrindo o arquivo oficial de KYC...")
        wb = excel.Workbooks.Open(str(arquivo_base_oficial))
        
        logger.info("Injetando dados e fórmulas na aba 'NC'...")
        try:
            ws_nc = wb.Sheets("NC")
            ws_nc.Columns("V:V").NumberFormat = "0"
            ultima_linha_antiga_nc = ws_nc.Cells(ws_nc.Rows.Count, 1).End(-4162).Row
            if ultima_linha_antiga_nc >= 3:
                ws_nc.Range(f"A3:AL{ultima_linha_antiga_nc}").ClearContents()
            ws_nc.Range("A2:AI2").ClearContents()
                
            ws_nc.Range(ws_nc.Cells(2, 1), ws_nc.Cells(ultima_linha_nc, 35)).Value = dados_nc
            
            if ultima_linha_nc > 2:
                ws_nc.Range(f"AJ2:AL{ultima_linha_nc}").FillDown()
                
                logger.info("Aplicando Estilo Zebrado nas linhas novas do NC...")
                ws_nc.Range("A2:AL3").Copy()
                ws_nc.Range(f"A2:AL{ultima_linha_nc}").PasteSpecial(Paste=-4122)
                excel.CutCopyMode = False
        except Exception as e:
            logger.error(f"Erro ao processar aba 'NC': {e}")
        
        logger.info("Injetando dados e fórmulas na aba 'FTD'...")
        try:
            ws_ftd = wb.Sheets("FTD")
            ultima_linha_antiga_ftd = ws_ftd.Cells(ws_ftd.Rows.Count, 1).End(-4162).Row
            if ultima_linha_antiga_ftd >= 3:
                ws_ftd.Range(f"A3:W{ultima_linha_antiga_ftd}").ClearContents()
            ws_ftd.Range("A2:U2").ClearContents()
                
            ws_ftd.Range(ws_ftd.Cells(2, 1), ws_ftd.Cells(ultima_linha_ftd, 21)).Value = dados_ftd
            
            if ultima_linha_ftd > 2:
                ws_ftd.Range(f"V2:W{ultima_linha_ftd}").FillDown()
                
                logger.info("Aplicando Estilo Zebrado nas linhas novas do FTD...")
                ws_ftd.Range("A2:W3").Copy()
                ws_ftd.Range(f"A2:W{ultima_linha_ftd}").PasteSpecial(Paste=-4122)
                excel.CutCopyMode = False
        except Exception as e:
            logger.error(f"Erro ao processar aba 'FTD': {e}")
        
        logger.info("Sincronizando Tabelas Dinâmicas da aba 'DIN'...")
        atualizar_dinamicas(wb)
        excel.CalculateUntilAsyncQueriesDone()
        
        logger.info("Salvando arquivo de KYC...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de KYC concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de KYC:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_mtd(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO ATUALIZAÇÃO INCREMENTAL: MTD ({marca}) ===")
    arquivo_mtd = obter_caminho_base(marca, "MTD", obter_data_alvo())
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
        atualizar_dinamicas(wb)
        excel.CalculateUntilAsyncQueriesDone()
        
        logger.info("Salvando arquivo MTD (Mãos longe do teclado!)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de MTD concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de MTD:")
        _fechar_excel_seguro(wb, excel)

# ==============================================================================
# BASE DE PERFORMANCE
# ==============================================================================
def carregar_base_performance_step1(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 1) ({marca}) ===")
    arquivo_transacoes = obter_caminho_base(marca, "Transacoes", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_transacoes.exists() or not arquivo_performance.exists():
        logger.error("Arquivos necessários para a Base Performance não encontrados.")
        return

    logger.info("Lendo Tabela Dinâmica 'Din_Diario' no arquivo de Transações...")
    df_din = pd.read_excel(arquivo_transacoes, sheet_name="Din_Diario", header=None)
    
    idx_header = df_din[df_din.apply(lambda r: r.astype(str).str.contains('Rótulos de Linha', case=False).any(), axis=1)].index
    
    if idx_header.empty:
        logger.error("Não foi possível encontrar 'Rótulos de Linha' na aba Din_Diario.")
        return
        
    linha_cabecalho = idx_header[0]
    df_din.columns = df_din.iloc[linha_cabecalho]
    df_din = df_din.iloc[linha_cabecalho + 1:].reset_index(drop=True)
    
    df_din = df_din[~df_din['Rótulos de Linha'].astype(str).str.contains('Total|Vazio|NaN|nan', case=False, na=False)]
    
    col_deposito = [c for c in df_din.columns if 'deposit' in str(c).lower()][0]
    col_saque = [c for c in df_din.columns if 'withdraw' in str(c).lower()][0]
    
    df_din['Dia'] = pd.to_numeric(df_din['Rótulos de Linha'], errors='coerce').fillna(0).astype(int)
    df_din[col_deposito] = pd.to_numeric(df_din[col_deposito], errors='coerce').fillna(0)
    df_din[col_saque] = pd.to_numeric(df_din[col_saque], errors='coerce').fillna(0)
    
    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    
    df_din = df_din[(df_din['Dia'] > 0) & (df_din['Dia'] <= limite_dia)].copy()
    
    df_din = garantir_continuidade_temporal(df_din, limite_dia)

    logger.info("Lendo histórico da 'BaseGeral' para auditoria...")
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,J,K", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')
    
    wb, excel = None, None
    try:
        _fazer_backup(arquivo_performance)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        # --- LÓGICA DE ABERTURA DE MÊS AUTOMÁTICA ---
        if not (datas_base == primeiro_dia_mes).any():
            logger.info("⚠️ MÊS NOVO DETECTADO! Iniciando abertura automática do mês na BaseGeral...")
            ultima_linha_preenchida = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            ref_anterior = int(ws.Cells(ultima_linha_preenchida, 1).Value)
            nova_ref = ref_anterior + 1
            
            dias_no_mes = calendar.monthrange(data_alvo.year, data_alvo.month)[1]
            meses_pt = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
            nome_mes = meses_pt[data_alvo.month]
            
            linha_inicio_nova = ultima_linha_preenchida + 1
            linha_fim_nova = linha_inicio_nova + dias_no_mes - 1
            
            logger.info(f" -> Criando {dias_no_mes} novas linhas para {nome_mes}...")
            for offset in range(dias_no_mes):
                r = linha_inicio_nova + offset
                ws.Cells(r, 1).Value = nova_ref
                ws.Cells(r, 2).Value = f"{offset+1:02d}/{data_alvo.month:02d}/{data_alvo.year}"
                ws.Cells(r, 3).Value = data_alvo.year
                ws.Cells(r, 4).Value = nome_mes
                ws.Cells(r, 5).Value = offset + 1
            
            logger.info(" -> Arrastando fórmulas das Colunas F até I...")
            ws.Range(f"F{ultima_linha_preenchida}:I{linha_fim_nova}").FillDown()
            
            logger.info(" -> Aplicando Estilo Zebrado no novo mês...")
            tint_antigo = ws.Cells(ultima_linha_preenchida, 1).Interior.TintAndShade
            
            # Lê a cor da última linha para inverter:
            novo_tint = -0.0499893185216834 if tint_antigo < -0.07 else -0.1499984740745262
            
            ws.Range(ws.Cells(linha_inicio_nova, 1), ws.Cells(linha_fim_nova, 65)).Interior.ThemeColor = 1
            ws.Range(ws.Cells(linha_inicio_nova, 1), ws.Cells(linha_fim_nova, 65)).Interior.TintAndShade = novo_tint

            wb.Save()
            logger.info("Mês aberto e formatado com sucesso! Prosseguindo com a injeção diária...")
            
            df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,J,K", header=0)
            datas_base = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')
            idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
            linha_excel_inicio = idx_inicio_pandas + 2 
        else:
            idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
            linha_excel_inicio = idx_inicio_pandas + 2
            
        dados_para_injetar = []
        for i, row in df_din.iterrows():
            dia = int(row['Dia'])
            novo_deposito = float(row[col_deposito])
            novo_saque = float(row[col_saque])
            dados_para_injetar.append([novo_deposito, novo_saque])
            
            idx_alvo = idx_inicio_pandas + (dia - 1)
            
            if idx_alvo < len(df_base):
                velho_deposito = pd.to_numeric(df_base.iloc[idx_alvo, 1], errors='coerce')
                velho_saque = pd.to_numeric(df_base.iloc[idx_alvo, 2], errors='coerce')
                
                if not np.isnan(velho_deposito):
                    if abs(velho_deposito - novo_deposito) > 0.01:
                        logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | Depósito: Histórico R$ {formatar_brl(velho_deposito)} -> Novo R$ {formatar_brl(novo_deposito)}")
                    else:
                        logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Depósito: Validado perfeitamente (R$ {formatar_brl(novo_deposito)})")
                
                if not np.isnan(velho_saque):
                    if abs(velho_saque - novo_saque) > 0.01:
                        logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | Saque: Histórico R$ {formatar_brl(velho_saque)} -> Novo R$ {formatar_brl(novo_saque)}")
                    else:
                        logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Saque: Validado perfeitamente (R$ {formatar_brl(novo_saque)})")

        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar) - 1
        logger.info(f"Sobrescrevendo colunas J (Depósitos) e K (Saques) das linhas {linha_excel_inicio} até {linha_excel_fim}...")
        ws.Range(ws.Cells(linha_excel_inicio, 10), ws.Cells(linha_excel_fim, 11)).Value = dados_para_injetar
        
        logger.info("Salvando Base Performance...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 1 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico ao gravar a Base Performance:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_performance_step2(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 2 - NC) ({marca}) ===")
    arquivo_nc = obter_caminho_base(marca, "NC", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_nc.exists() or not arquivo_performance.exists():
        logger.error("Arquivos necessários para a Base Performance (Step 2) não encontrados.")
        return

    logger.info("Lendo Tabelas Dinâmicas 'Din_Diario' no arquivo de Novas Contas...")
    df_din_full = pd.read_excel(arquivo_nc, sheet_name="Din_Diario", header=None)
    
    def extrair_dinamica(df_pedaco, nome_coluna_valor):
        mask = df_pedaco.apply(lambda r: r.astype(str).str.contains('Rótulos de Linha', case=False).any(), axis=1)
        if not mask.any(): return pd.DataFrame()
        
        idx = df_pedaco[mask].index[0]
        df = df_pedaco.iloc[idx+1:].copy()
        df.columns = ['Dia_raw', 'Valor_raw']
        
        df['Dia'] = pd.to_numeric(df['Dia_raw'], errors='coerce')
        df = df.dropna(subset=['Dia'])
        df['Dia'] = df['Dia'].astype(int)
        df = df[df['Dia'] > 0] 
        
        df[nome_coluna_valor] = pd.to_numeric(df['Valor_raw'], errors='coerce').fillna(0).astype(int)
        return df[['Dia', nome_coluna_valor]].copy()

    df_p1 = extrair_dinamica(df_din_full.iloc[:, [0, 1]], 'NC_Total')
    df_p2 = extrair_dinamica(df_din_full.iloc[:, [4, 5]], 'NC_Organicos')

    if df_p1.empty or df_p2.empty:
        logger.error("Não foi possível ler as duas tabelas dinâmicas corretamente.")
        return

    df_mesclado = pd.merge(df_p1, df_p2, on='Dia', how='outer').fillna(0).sort_values('Dia').reset_index(drop=True)
    
    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    
    df_mesclado = df_mesclado[(df_mesclado['Dia'] > 0) & (df_mesclado['Dia'] <= limite_dia)].copy()
    
    if df_mesclado.empty:
        logger.error("Nenhum dia válido encontrado nas tabelas dinâmicas até o dia anterior.")
        return
        
    df_mesclado = garantir_continuidade_temporal(df_mesclado, limite_dia)

    logger.info("Lendo histórico da 'BaseGeral' para auditoria...")
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,N,O", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')
    
    if not (datas_base == primeiro_dia_mes).any():
        logger.error(f"A data {primeiro_dia_mes.strftime('%d/%m/%Y')} não foi encontrada na coluna B da BaseGeral!")
        return
        
    idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio_pandas + 2 
    
    logger.info(f"Data {primeiro_dia_mes.strftime('%d/%m/%Y')} encontrada na linha {linha_excel_inicio} do Excel. Iniciando reconciliação...")
    
    dados_para_injetar = []
    
    for i, row in df_mesclado.iterrows():
        dia = int(row['Dia'])
        novo_organico = int(row['NC_Organicos'])
        novo_total = int(row['NC_Total'])
        
        dados_para_injetar.append([novo_organico, novo_total])
        
        idx_alvo = idx_inicio_pandas + (dia - 1)
        
        if idx_alvo < len(df_base):
            velho_organico = pd.to_numeric(df_base.iloc[idx_alvo, 1], errors='coerce')
            velho_total = pd.to_numeric(df_base.iloc[idx_alvo, 2], errors='coerce')
            
            if not np.isnan(velho_organico):
                if int(velho_organico) != novo_organico:
                    logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | NC Orgânicos: Histórico {formatar_int(velho_organico)} -> Novo {formatar_int(novo_organico)}")
                else:
                    logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | NC Orgânicos: Validado ({formatar_int(novo_organico)})")
            
            if not np.isnan(velho_total):
                if int(velho_total) != novo_total:
                    logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | NC Total: Histórico {formatar_int(velho_total)} -> Novo {formatar_int(novo_total)}")
                else:
                    logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | NC Total: Validado ({formatar_int(novo_total)})")

    wb, excel = None, None
    try:
        #_fazer_backup(arquivo_performance)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar) - 1
        
        logger.info(f"Sobrescrevendo colunas N (Orgânicos) e O (Total) das linhas {linha_excel_inicio} até {linha_excel_fim}...")
        ws.Range(ws.Cells(linha_excel_inicio, 14), ws.Cells(linha_excel_fim, 15)).Value = dados_para_injetar
        
        logger.info("Salvando Base Performance...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 2 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico ao gravar a Base Performance:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_performance_step3(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 3 - FTD) ({marca}) ===")
    arquivo_ftd = obter_caminho_base(marca, "FTD", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_ftd.exists() or not arquivo_performance.exists():
        logger.error("Arquivos necessários para a Base Performance (Step 3) não encontrados.")
        return

    def extrair_dinamica_bloco(df_source, col_idx_dia, col_idx_valores, col_names):
        mask = df_source.iloc[:, col_idx_dia].astype(str).str.contains('Rótulos de Linha', case=False)
        if not mask.any(): return pd.DataFrame()
        
        idx = df_source[mask].index[0]
        df = df_source.iloc[idx+1:].copy()
        
        df = df.iloc[:, [col_idx_dia] + col_idx_valores]
        df.columns = ['Dia_raw'] + col_names
        
        df['Dia'] = pd.to_numeric(df['Dia_raw'], errors='coerce')
        df = df.dropna(subset=['Dia'])
        df['Dia'] = df['Dia'].astype(int)
        df = df[df['Dia'] > 0]
        
        for col in col_names:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
        return df[['Dia'] + col_names].copy()

    logger.info("Lendo Tabela Dinâmica 'Din_Afiliados' (Coluna X)...")
    df_afil_full = pd.read_excel(arquivo_ftd, sheet_name="Din_Afiliados", header=None)
    df_x = extrair_dinamica_bloco(df_afil_full, 3, [4], ['FTD_Org'])

    logger.info("Lendo Tabela Dinâmica 'Din_Diario' (Colunas P a U)...")
    df_diario_full = pd.read_excel(arquivo_ftd, sheet_name="Din_Diario", header=None)
    nomes_colunas_pu = ['Col_P', 'Col_Q', 'Col_R', 'Col_S', 'Col_T', 'Col_U']
    df_pu = extrair_dinamica_bloco(df_diario_full, 0, [1, 2, 3, 4, 5, 6], nomes_colunas_pu)

    if df_x.empty or df_pu.empty:
        logger.error("Falha ao ler as tabelas dinâmicas do FTD. Verifique o layout do arquivo.")
        return

    df_mesclado = pd.merge(df_pu, df_x, on='Dia', how='outer').fillna(0).sort_values('Dia').reset_index(drop=True)
    
    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    
    df_mesclado = df_mesclado[(df_mesclado['Dia'] > 0) & (df_mesclado['Dia'] <= limite_dia)].copy()

    if df_mesclado.empty:
        logger.error("Nenhum dia válido encontrado no FTD até o dia anterior.")
        return
        
    df_mesclado = garantir_continuidade_temporal(df_mesclado, limite_dia)

    logger.info("Lendo histórico da 'BaseGeral' para auditoria...")
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,P:U,X", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')
    
    if not (datas_base == primeiro_dia_mes).any():
        logger.error(f"A data {primeiro_dia_mes.strftime('%d/%m/%Y')} não foi encontrada na coluna B da BaseGeral!")
        return
        
    idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio_pandas + 2 
    
    logger.info(f"Data {primeiro_dia_mes.strftime('%d/%m/%Y')} encontrada na linha {linha_excel_inicio} do Excel. Iniciando reconciliação...")
    
    dados_para_injetar_PU = []
    dados_para_injetar_X = []
    
    letras_pu = ['P', 'Q', 'R', 'S', 'T', 'U']
    
    for i, row in df_mesclado.iterrows():
        dia = int(row['Dia'])
        novos_pu = [row[col] for col in nomes_colunas_pu]
        novo_x = row['FTD_Org']
        
        dados_para_injetar_PU.append(novos_pu)
        dados_para_injetar_X.append([novo_x])
        
        idx_alvo = idx_inicio_pandas + (dia - 1)
        
        if idx_alvo < len(df_base):
            for j in range(6):
                velho_val = pd.to_numeric(df_base.iloc[idx_alvo, j+1], errors='coerce')
                novo_val = novos_pu[j]
                letra = letras_pu[j]
                
                if not np.isnan(velho_val):
                    if abs(velho_val - novo_val) > 0.01:
                        logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | Col {letra}: Histórico {formatar_num(velho_val)} -> Novo {formatar_num(novo_val)}")
                    else:
                        logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Col {letra}: Validado ({formatar_num(novo_val)})")
                        
            velho_x = pd.to_numeric(df_base.iloc[idx_alvo, 7], errors='coerce')
            if not np.isnan(velho_x):
                if abs(velho_x - novo_x) > 0.01:
                    logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | Col X (Orgânicos): Histórico {formatar_num(velho_x)} -> Novo {formatar_num(novo_x)}")
                else:
                    logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Col X (Orgânicos): Validado ({formatar_num(novo_x)})")

    wb, excel = None, None
    try:
        #_fazer_backup(arquivo_performance)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar_PU) - 1
        
        logger.info(f"Sobrescrevendo colunas P até U (linhas {linha_excel_inicio} até {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 16), ws.Cells(linha_excel_fim, 21)).Value = dados_para_injetar_PU
        
        logger.info(f"Sobrescrevendo coluna X (linhas {linha_excel_inicio} até {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 24), ws.Cells(linha_excel_fim, 24)).Value = dados_para_injetar_X
        
        logger.info("Salvando Base Performance...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 3 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico ao gravar a Base Performance:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_performance_step4(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 4 - MTD) ({marca}) ===")
    arquivo_mtd = obter_caminho_base(marca, "MTD", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_mtd.exists() or not arquivo_performance.exists():
        logger.error("Arquivos necessários para a Base Performance (Step 4) não encontrados.")
        return

    logger.info("Lendo aba 'Din' no arquivo MTD (Colunas I a O)...")
    df_mtd = pd.read_excel(arquivo_mtd, sheet_name="Din", header=4, usecols="I:O")
    
    nomes_colunas_at_ay = ['Col_AT', 'Col_AU', 'Col_AV', 'Col_AW', 'Col_AX', 'Col_AY']
    df_mtd.columns = ['Dia_raw'] + nomes_colunas_at_ay
    
    df_mtd['Dia'] = pd.to_numeric(df_mtd['Dia_raw'], errors='coerce')
    df_mtd = df_mtd.dropna(subset=['Dia'])
    df_mtd['Dia'] = df_mtd['Dia'].astype(int)
    
    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    
    logger.info(f"Aplicando filtro de data: Coletando dados até o dia de ontem ({limite_dia:02d})...")
    df_mtd = df_mtd[(df_mtd['Dia'] > 0) & (df_mtd['Dia'] <= limite_dia)].copy()
    df_mtd = df_mtd.sort_values('Dia').reset_index(drop=True)

    if df_mtd.empty:
        logger.error("Nenhum dia válido encontrado no MTD após aplicar o filtro de data.")
        return
        
    df_mtd = garantir_continuidade_temporal(df_mtd, limite_dia)

    logger.info("Lendo histórico da 'BaseGeral' para auditoria...")
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,AT:AY", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')
    
    if not (datas_base == primeiro_dia_mes).any():
        logger.error(f"A data {primeiro_dia_mes.strftime('%d/%m/%Y')} não foi encontrada na coluna B da BaseGeral!")
        return
        
    idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio_pandas + 2 
    
    logger.info(f"Data {primeiro_dia_mes.strftime('%d/%m/%Y')} encontrada na linha {linha_excel_inicio}. Iniciando reconciliação das colunas AT até AY...")
    
    dados_para_injetar = []
    letras_at_ay = ['AT', 'AU', 'AV', 'AW', 'AX', 'AY']
    
    for i, row in df_mtd.iterrows():
        dia = int(row['Dia'])
        novos_valores = [int(row[col]) for col in nomes_colunas_at_ay]
        dados_para_injetar.append(novos_valores)
        
        idx_alvo = idx_inicio_pandas + (dia - 1)
        if idx_alvo < len(df_base):
            for j in range(6):
                velho_val_raw = pd.to_numeric(df_base.iloc[idx_alvo, j+1], errors='coerce')
                if not np.isnan(velho_val_raw):
                    velho_val = int(velho_val_raw)
                    novo_val = novos_valores[j]
                    letra = letras_at_ay[j]
                    
                    if letra == 'AT':
                        if velho_val != novo_val:
                            logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | Col {letra}: Histórico {formatar_int(velho_val)} -> Novo {formatar_int(novo_val)}")
                        else:
                            logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Col {letra}: Validado ({formatar_int(novo_val)})")
                    else:
                        if velho_val > novo_val: 
                            logger.warning(f"[AUDITORIA - REDUÇÃO CRÍTICA] Dia {dia:02d} | Col {letra}: Valor REDUZIU! {formatar_int(velho_val)} -> {formatar_int(novo_val)}")
                        elif novo_val > velho_val: 
                            logger.info(f"[AUDITORIA - AUMENTO OK]    Dia {dia:02d} | Col {letra}: Cresceu de {formatar_int(velho_val)} para {formatar_int(novo_val)}")
                        else: 
                            logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Col {letra}: Validado ({formatar_int(novo_val)})")

    wb, excel = None, None
    try:
        #_fazer_backup(arquivo_performance)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar) - 1
        
        logger.info(f"Sobrescrevendo colunas AT até AY (linhas {linha_excel_inicio} até {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 46), ws.Cells(linha_excel_fim, 51)).Value = dados_para_injetar
        
        logger.info("Salvando Base Performance...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 4 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico ao gravar a Base Performance:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_performance_step5(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 5 - KYC) ({marca}) ===")
    arquivo_kyc = obter_caminho_base(marca, "KYC", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_kyc.exists() or not arquivo_performance.exists():
        logger.error("Arquivos necessários para a Base Performance (Step 5) não encontrados.")
        return

    logger.info("Lendo aba 'DIN' no arquivo KYC...")
    df_din_full = pd.read_excel(arquivo_kyc, sheet_name="DIN", header=None)
    
    def extrair_tabela(df_source, col_idx_dia, col_idx_valores, col_names):
        mask = df_source.iloc[:, col_idx_dia].astype(str).str.contains('Rótulos|Dia', case=False)
        if not mask.any(): return pd.DataFrame()
        
        idx = df_source[mask].index[0]
        df = df_source.iloc[idx+1:].copy()
        
        df = df.iloc[:, [col_idx_dia] + col_idx_valores]
        df.columns = ['Dia_raw'] + col_names
        
        df['Dia'] = pd.to_numeric(df['Dia_raw'], errors='coerce')
        df = df.dropna(subset=['Dia'])
        df['Dia'] = df['Dia'].astype(int)
        df = df[df['Dia'] > 0]
        
        for col in col_names:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            
        return df[['Dia'] + col_names].copy()

    df_p1 = extrair_tabela(df_din_full, 0, [1], ['KYC_True'])
    df_p2 = extrair_tabela(df_din_full, 4, [5], ['KYC_FTD_True'])
    df_p3 = extrair_tabela(df_din_full, 8, [9, 10, 11], ['Col_BK', 'Col_BL', 'Col_BM'])

    if df_p1.empty or df_p2.empty or df_p3.empty:
        logger.error("Falha ao ler as tabelas dinâmicas do KYC. Verifique o layout do arquivo.")
        return

    df_mesclado = pd.merge(df_p1, df_p2, on='Dia', how='outer')
    df_mesclado = pd.merge(df_mesclado, df_p3, on='Dia', how='outer').fillna(0).sort_values('Dia').reset_index(drop=True)
    
    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    
    logger.info(f"Aplicando filtro de data: Coletando dados até o dia de ontem ({limite_dia:02d})...")
    df_mesclado = df_mesclado[(df_mesclado['Dia'] > 0) & (df_mesclado['Dia'] <= limite_dia)].copy()

    if df_mesclado.empty:
        logger.error("Nenhum dia válido encontrado no KYC após aplicar o filtro de data.")
        return
        
    df_mesclado = garantir_continuidade_temporal(df_mesclado, limite_dia)

    logger.info("Lendo histórico da 'BaseGeral' para auditoria...")
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,BF,BH,BK:BM", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')
    
    if not (datas_base == primeiro_dia_mes).any():
        logger.error(f"A data {primeiro_dia_mes.strftime('%d/%m/%Y')} não foi encontrada na coluna B da BaseGeral!")
        return
        
    idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio_pandas + 2 
    
    logger.info(f"Data {primeiro_dia_mes.strftime('%d/%m/%Y')} encontrada na linha {linha_excel_inicio}. Iniciando reconciliação...")
    
    dados_BF = []
    dados_BH = []
    dados_BK_BM = []
    
    for i, row in df_mesclado.iterrows():
        dia = int(row['Dia'])
        novo_bf = int(row['KYC_True'])
        novo_bh = int(row['KYC_FTD_True'])
        novos_bk_bm = [int(row['Col_BK']), int(row['Col_BL']), int(row['Col_BM'])]
        
        dados_BF.append([novo_bf])
        dados_BH.append([novo_bh])
        dados_BK_BM.append(novos_bk_bm)
        
        idx_alvo = idx_inicio_pandas + (dia - 1)
        if idx_alvo < len(df_base):
            velho_bf = pd.to_numeric(df_base.iloc[idx_alvo, 1], errors='coerce')
            velho_bh = pd.to_numeric(df_base.iloc[idx_alvo, 2], errors='coerce')
            velhos_bk_bm = [pd.to_numeric(df_base.iloc[idx_alvo, j], errors='coerce') for j in [3, 4, 5]]
            
            for velho_val_raw, novo_val, nome_col in zip([velho_bf, velho_bh], [novo_bf, novo_bh], ['BF', 'BH']):
                if not np.isnan(velho_val_raw):
                    velho_val = int(velho_val_raw)
                    if velho_val > novo_val:  
                        logger.warning(f"[AUDITORIA - REDUÇÃO CRÍTICA] Dia {dia:02d} | Col {nome_col}: Valor REDUZIU! {formatar_int(velho_val)} -> {formatar_int(novo_val)}")
                    elif novo_val > velho_val: 
                        logger.info(f"[AUDITORIA - AUMENTO OK]    Dia {dia:02d} | Col {nome_col}: Cresceu de {formatar_int(velho_val)} para {formatar_int(novo_val)}")
                    else: 
                        logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Col {nome_col}: Validado ({formatar_int(novo_val)})")

            letras_bk_bm = ['BK', 'BL', 'BM']
            for j in range(3):
                velho_val_raw = velhos_bk_bm[j]
                novo_val = novos_bk_bm[j]
                letra = letras_bk_bm[j]
                if not np.isnan(velho_val_raw):
                    velho_val = int(velho_val_raw)
                    if velho_val != novo_val:
                        logger.warning(f"[AUDITORIA - DIVERGÊNCIA] Dia {dia:02d} | Col {letra}: Histórico {formatar_int(velho_val)} -> Novo {formatar_int(novo_val)}")
                    else:
                        logger.info(f"[AUDITORIA - CONFERIDO]   Dia {dia:02d} | Col {letra}: Validado ({formatar_int(novo_val)})")

    wb, excel = None, None
    try:
        #_fazer_backup(arquivo_performance)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_BF) - 1
        
        logger.info(f"Sobrescrevendo coluna BF (linha {linha_excel_inicio} a {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 58), ws.Cells(linha_excel_fim, 58)).Value = dados_BF
        
        logger.info(f"Sobrescrevendo coluna BH (linha {linha_excel_inicio} a {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 60), ws.Cells(linha_excel_fim, 60)).Value = dados_BH
        
        logger.info(f"Sobrescrevendo colunas BK até BM (linha {linha_excel_inicio} a {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 63), ws.Cells(linha_excel_fim, 65)).Value = dados_BK_BM
        
        logger.info("Salvando Base Performance...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 5 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico ao gravar a Base Performance:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_performance_step6(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 6 - GEN STATS) ({marca}) ===")
    arquivo_json = _obter_caminho_download(marca, f"GeneralStats - {marca}.json")
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_json or not arquivo_json.exists() or not arquivo_performance.exists():
        logger.error("Arquivos necessários para a Base Performance (Step 6) não encontrados.")
        return

    mapeamento_upgaming = {
        "Sportsbook": {"bet": 0, "win": 1, "users": 12},  
        "LiveCasino": {"bet": 3, "win": 4, "users": 13},  
        "Slot":       {"bet": 6, "win": 7, "users": 14},  
        "MiniGames":  {"bet": 9, "win": 10, "users": 15}  
    }

    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    
    logger.info("Lendo dados extraídos do arquivo JSON local...")
    with open(arquivo_json, 'r', encoding='utf-8') as f:
        dados_json_mensal = json.load(f)
        
    dados_por_dia = {item["Dia"]: item["dados"] for item in dados_json_mensal}
    dados_para_injetar = []
    
    logger.info(f"Processando dados do dia 01 até {limite_dia:02d}...")
    for dia in range(1, limite_dia + 1):
        json_do_dia = dados_por_dia.get(dia, [])
        if not json_do_dia:
            logger.warning(f"⚠️ ALERTA: Dados do dia {dia:02d} não encontrados no JSON. Preenchendo com zeros.")
            
        linha_excel = [0.0] * 16 
        for jogo in json_do_dia:
            tipo = jogo.get("gameType")
            if tipo in mapeamento_upgaming:
                idx_bet = mapeamento_upgaming[tipo]["bet"]
                idx_win = mapeamento_upgaming[tipo]["win"]
                idx_users = mapeamento_upgaming[tipo]["users"]
                linha_excel[idx_bet] = float(jogo.get("betAmount", 0))
                linha_excel[idx_win] = float(jogo.get("winAmount", 0))
                linha_excel[idx_users] = int(jogo.get("userCount", 0))
                
        dados_para_injetar.append(linha_excel)

    logger.info("Estruturação concluída! Iniciando Auditoria Estrita no Excel...")
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,AC:AR", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    idx_inicio = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')[lambda x: x == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio + 2 
    
    colunas_auditar = {
        0: ('AC', 'Bet_Sb'), 1: ('AD', 'Win_Sb'), 12: ('AO', 'Users_Sb'),
        3: ('AF', 'Bet_Lc'), 4: ('AG', 'Win_Lc'), 13: ('AP', 'Users_Lc'),
        6: ('AI', 'Bet_St'), 7: ('AJ', 'Win_St'), 14: ('AQ', 'Users_St'),
        9: ('AL', 'Bet_Mg'), 10: ('AM', 'Win_Mg'), 15: ('AR', 'Users_Mg')
    }

    for dia_idx, linha_nova in enumerate(dados_para_injetar):
        dia_real = dia_idx + 1
        idx_alvo = idx_inicio + dia_idx
        
        if idx_alvo < len(df_base):
            for col_index, (letra, nome) in colunas_auditar.items():
                velho_val_raw = pd.to_numeric(df_base.iloc[idx_alvo, col_index + 1], errors='coerce') 
                novo_val = linha_nova[col_index]
                
                if not np.isnan(velho_val_raw):
                    if "Users" in nome:
                        v_int = int(velho_val_raw); n_int = int(novo_val)
                        if v_int != n_int: 
                            logger.warning(f"[DIVERGÊNCIA CRÍTICA] Dia {dia_real:02d} | {letra} ({nome}): {formatar_int(v_int)} mudou para {formatar_int(n_int)}")
                        else:
                            logger.info(f"[CONFERIDO]   Dia {dia_real:02d} | {letra} ({nome}): Intacto ({formatar_int(n_int)})")
                    else:
                        if abs(velho_val_raw - novo_val) > 0.01:
                            logger.warning(f"[DIVERGÊNCIA CRÍTICA] Dia {dia_real:02d} | {letra} ({nome}): R$ {formatar_brl(velho_val_raw)} mudou para R$ {formatar_brl(novo_val)}")
                        else:
                            logger.info(f"[CONFERIDO]   Dia {dia_real:02d} | {letra} ({nome}): Intacto (R$ {formatar_brl(novo_val)})")

    wb, excel = None, None
    try:
        #_fazer_backup(arquivo_performance)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True; excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar) - 1
        logger.info(f"Sobrescrevendo bloco AC até AR (linhas {linha_excel_inicio} a {linha_excel_fim})...")
        
        ws.Range(ws.Cells(linha_excel_inicio, 29), ws.Cells(linha_excel_fim, 44)).Value = dados_para_injetar
        
        if linha_excel_inicio >= 3:
            logger.info("Arrastando as fórmulas das colunas AB, AE, AH, AK e AN...")
            linha_referencia = linha_excel_inicio - 1
            for col_f in ['AB', 'AE', 'AH', 'AK', 'AN']:
                ws.Range(f"{col_f}{linha_referencia}:{col_f}{linha_excel_fim}").FillDown()
        
        wb.Save(); wb.Close(); excel.Quit()
        logger.info("-> Step 6 (General Statistics) Concluído com Sucesso!")
    except Exception:
        logger.exception("Erro crítico na gravação do Step 6:")
        _fechar_excel_seguro(wb, excel)

def carregar_base_performance_step7(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 7 - USUÁRIOS ÚNICOS) ({marca}) ===")
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_performance.exists():
        logger.error("Base Performance não encontrada para o Step 7.")
        return

    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    mes_atual = data_alvo.month
    
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,AS", header=0)
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    idx_inicio = pd.to_datetime(df_base.iloc[:, 0], errors='coerce')[lambda x: x == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio + 2 

    dados_AS = []
    pasta_ugs = obter_pasta_ugs_diario(marca, data_alvo.year, data_alvo.month)
    logger.info(f"Lendo e contando os arquivos diários do UGS (01 até {limite_dia:02d})...")
    
    for dia in range(1, limite_dia + 1):
        nome_arquivo = f"{dia:02d}-{mes_atual:02d}.xlsx"
        caminho_ugs = pasta_ugs / nome_arquivo
        
        qtd_usuarios = 0
        if caminho_ugs.exists():
            try:
                df_ugs = pd.read_excel(caminho_ugs, engine='calamine')
                qtd_usuarios = len(df_ugs)
            except Exception as e:
                logger.error(f"Erro ao ler {nome_arquivo}: {e}")
        else:
            logger.warning(f"⚠️ ALERTA: Arquivo diário {nome_arquivo} não encontrado! Preenchendo com 0.")

        dados_AS.append([qtd_usuarios])
        
        idx_alvo = idx_inicio + (dia - 1)
        if idx_alvo < len(df_base):
            velho_val = pd.to_numeric(df_base.iloc[idx_alvo, 1], errors='coerce')
            if not np.isnan(velho_val):
                v_int = int(velho_val)
                if v_int != qtd_usuarios:
                    logger.warning(f"[DIVERGÊNCIA] Dia {dia:02d} | Col AS (Users Únicos): Histórico {v_int} -> Novo {qtd_usuarios}")
                else:
                    logger.info(f"[CONFERIDO]   Dia {dia:02d} | Col AS (Users Únicos): Validado ({qtd_usuarios})")

    wb, excel = None, None
    try:
        #_fazer_backup(arquivo_performance)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True; excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_AS) - 1
        logger.info(f"Sobrescrevendo coluna AS (linhas {linha_excel_inicio} a {linha_excel_fim})...")
        
        ws.Range(ws.Cells(linha_excel_inicio, 45), ws.Cells(linha_excel_fim, 45)).Value = dados_AS
        
        wb.Save(); wb.Close(); excel.Quit()
        logger.info("-> Step 7 (Usuários Únicos) Concluído com Sucesso!")
    except Exception:
        logger.exception("Erro crítico na gravação do Step 7:")
        _fechar_excel_seguro(wb, excel)