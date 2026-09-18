"""
Carregamento do relatório de System Transactions na base oficial da marca,
incluindo a atualização da aba BaseCompleta (colunas de referência de usuário).
"""
import logging

import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import blindar_dados
from utils.date_utils import obter_data_alvo
from utils.excel_utils import aplicar_filtro_dinamica, atualizar_dinamicas, _fechar_excel_seguro
from utils.file_utils import MESES_PT, _fazer_backup, _obter_caminho_download, obter_caminho_base, obter_caminho_base_completa

logger = logging.getLogger(__name__)


def carregar_base_transacoes(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO CARREGAMENTO: TRANSAÇÕES E BASE COMPLETA ({marca}) ===")
    arquivo_origem_base = obter_caminho_base_completa(marca)
    arquivo_origem_transacoes = _obter_caminho_download(marca, f"Transações - {marca}.xlsx")
    data_alvo = obter_data_alvo()
    arquivo_base_oficial = obter_caminho_base(marca, "Transacoes", data_alvo)
    
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

        mes_str = MESES_PT[data_alvo.month][0].lower()

        for aba in ["Din_Type", "Din_Afiliados", "MTD_ID", "Din_Diario"]:
            ws_aba = wb.Sheets(aba)
            aplicar_filtro_dinamica(ws_aba, "B2", "Success")
            aplicar_filtro_dinamica(ws_aba, "B3", mes_str)

        for aba in ["Din_Afiliados", "MTD_ID", "Din_Diario"]:
            ws_alvo = wb.Sheets(aba)
            aplicar_filtro_dinamica(ws_alvo, "A7", valor_desejado="(Tudo)")
            aplicar_filtro_dinamica(ws_alvo, "B5", valor_desejado=["Deposit", "Withdraw", "Bonus Activation"])


        ws_cashback = wb.Sheets("Afiliados_Cashback")
        # Tabela 1
        aplicar_filtro_dinamica(ws_cashback, "A8", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_cashback, "B3", "Sim")
        aplicar_filtro_dinamica(ws_cashback, "B4", mes_str)
        aplicar_filtro_dinamica(ws_cashback, "B5", "Success")
        aplicar_filtro_dinamica(ws_cashback, "B6", "Cash Bonus")
        
        # Tabela 2
        aplicar_filtro_dinamica(ws_cashback, "D8", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_cashback, "E3", "Sim")
        aplicar_filtro_dinamica(ws_cashback, "E4", mes_str)
        aplicar_filtro_dinamica(ws_cashback, "E5", "Success")
        aplicar_filtro_dinamica(ws_cashback, "E6", "LeaderBoard Cash Deposit" )
        
        # Tabela 3
        aplicar_filtro_dinamica(ws_cashback, "G8", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_cashback, "H3", "Sim")
        aplicar_filtro_dinamica(ws_cashback, "H4", mes_str)
        aplicar_filtro_dinamica(ws_cashback, "H5", "Success")
        aplicar_filtro_dinamica(ws_cashback, "H6", "(blank)")
        
        # Tabela 4
        aplicar_filtro_dinamica(ws_cashback, "K8", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_cashback, "L3", "Sim")
        aplicar_filtro_dinamica(ws_cashback, "L4", mes_str)
        aplicar_filtro_dinamica(ws_cashback, "L5", "Success")
        aplicar_filtro_dinamica(ws_cashback, "L6", "(Tudo)")

        ws_mtd_perf = wb.Sheets("MTD_Performance")
        aplicar_filtro_dinamica(ws_mtd_perf, "A5", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_mtd_perf, "B1", "Deposit")
        aplicar_filtro_dinamica(ws_mtd_perf, "B2", "Success")

        ws_fortune = wb.Sheets("fortune")
        aplicar_filtro_dinamica(ws_fortune, "A5", valor_desejado="(Tudo)", exceto="(blank)")
        aplicar_filtro_dinamica(ws_fortune, "B1", "danielfortune" )
        aplicar_filtro_dinamica(ws_fortune, "B2", "Success")
        aplicar_filtro_dinamica(ws_fortune, "B3", "Deposit")

        logger.info("Salvando arquivo de Transações (Isto pode demorar. Mãos longe do teclado!)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de Transações concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de Transações:")
        _fechar_excel_seguro(wb, excel)