"""
Carregamento do relatório de Primeiros Depositantes (FTD) na base oficial da marca.
"""
import logging

import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import blindar_dados
from utils.date_utils import obter_data_alvo
from utils.excel_utils import (
    _fechar_excel_seguro,
    aplicar_filtro_dinamica,
    atualizar_dinamicas,
)
from utils.file_utils import _fazer_backup, _obter_caminho_download, obter_caminho_base

logger = logging.getLogger(__name__)


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
            ws.Range(f"A3:AD{ultima_linha_antiga}").Delete(Shift=-4162)
        ws.Range("A2:U2").ClearContents()

        ws.Range(ws.Cells(2, 1), ws.Cells(ultima_linha_destino, 21)).Value = dados_a_inserir
        
        if ultima_linha_destino > 2:
            ws.Range(f"V2:AD{ultima_linha_destino}").FillDown()
        
        atualizar_dinamicas(wb, excel)

        ws_din_afiliado = wb.Sheets("Din_Afiliados")
        aplicar_filtro_dinamica(ws_din_afiliado, "B2", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_din_afiliado, "E2", valor_desejado="Orgânicos")

        ws_fortune = wb.Sheets("Fortune")
        aplicar_filtro_dinamica(ws_fortune, "B2", valor_desejado="danielfortune")

        ws_din_diario = wb.Sheets("Din_Diario")
        aplicar_filtro_dinamica(ws_din_diario, "A4", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_din_diario, "B3", valor_desejado=["De R$0 até R$10", "De R$10 até R$20", "De R$20 até R$30", "De R$30 até R$50", "Superior a R$50"])

        ws_din_faixa = wb.Sheets("Din_FaixaPagamento")
        aplicar_filtro_dinamica(ws_din_faixa, "A3", valor_desejado="(Tudo)")

        logger.info("Salvando arquivo de FTD (Por favor, aguarde)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de FTD concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de FTD:")
        _fechar_excel_seguro(wb, excel)
        raise