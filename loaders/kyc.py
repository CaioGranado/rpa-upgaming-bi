"""
Carregamento do relatório de KYC (abas NC e FTD) na base oficial da marca.
"""
import logging

import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import aplicar_corte_datas_futuras, blindar_dados
from utils.date_utils import obter_data_alvo
from utils.excel_utils import aplicar_filtro_dinamica, atualizar_dinamicas, _fechar_excel_seguro
from utils.file_utils import _fazer_backup, _obter_caminho_download, obter_caminho_base

logger = logging.getLogger(__name__)


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
        
        logger.info("Injetando dados e fórmulas na aba 'FTD'...")
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
        
        logger.info("Sincronizando Tabelas Dinâmicas da aba 'DIN'...")
        atualizar_dinamicas(wb)
        excel.CalculateUntilAsyncQueriesDone()

        ws_din_kyc = wb.Sheets("DIN")
        aplicar_filtro_dinamica(ws_din_kyc, "B2", "True")
        aplicar_filtro_dinamica(ws_din_kyc, "F1", "Sim")
        aplicar_filtro_dinamica(ws_din_kyc, "F2", "True")
        aplicar_filtro_dinamica(ws_din_kyc, "A4", valor_desejado="(Tudo)", exceto="(blank)")
        aplicar_filtro_dinamica(ws_din_kyc, "E4", valor_desejado="(Tudo)", exceto="(blank)")
        aplicar_filtro_dinamica(ws_din_kyc, "I4", valor_desejado="(Tudo)", exceto="(blank)")
        aplicar_filtro_dinamica(ws_din_kyc, "J4", valor_desejado="(Tudo)", exceto="(blank)")

        logger.info("Salvando arquivo de KYC...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de KYC concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de KYC:")
        _fechar_excel_seguro(wb, excel)