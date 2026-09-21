"""
Carregamento do relatório de Novas Contas (NC) na base oficial da marca.
"""
import logging

import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import aplicar_corte_datas_futuras, blindar_dados
from utils.date_utils import obter_data_alvo
from utils.excel_utils import (
    _fechar_excel_seguro,
    aplicar_filtro_dinamica,
    atualizar_dinamicas,
)
from utils.file_utils import _fazer_backup, _obter_caminho_download, obter_caminho_base

logger = logging.getLogger(__name__)


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
            ws.Range(f"A3:AR{ultima_linha_antiga}").Delete(Shift=-4162)
            
        ws.Range("A2:AI2").ClearContents()
        if hasattr(ws, "Range"): ws.Range("AP2:AP2").ClearContents()

        ws.Range(ws.Cells(2, 1), ws.Cells(ultima_linha_destino, 35)).Value = dados_a_inserir
        ws.Range(ws.Cells(2, 42), ws.Cells(ultima_linha_destino, 42)).Value = dados_ddd
        
        if ultima_linha_destino > 2:
            ws.Range(f"AJ2:AO{ultima_linha_destino}").FillDown()
            ws.Range(f"AQ2:AR{ultima_linha_destino}").FillDown()
        
        atualizar_dinamicas(wb, excel)

        ws_din = wb.Sheets("Din_Diario")
        aplicar_filtro_dinamica(ws_din, "B1", valor_desejado="(Tudo)")
        aplicar_filtro_dinamica(ws_din, "F1", valor_desejado="Orgânicos")
        aplicar_filtro_dinamica(ws_din, "A3", valor_desejado="(Tudo)", exceto="(blank)")
        aplicar_filtro_dinamica(ws_din, "E3", valor_desejado="(Tudo)", exceto="(blank)")

        ws_din_regiao = wb.Sheets("Din_Regiao")
        aplicar_filtro_dinamica(ws_din_regiao, "B2", valor_desejado="Orgânicos")
        aplicar_filtro_dinamica(ws_din_regiao, "F2", valor_desejado="(Tudo)", exceto="Orgânicos")

        ws_fortune = wb.Sheets("Fortune")
        aplicar_filtro_dinamica(ws_fortune, "B2", valor_desejado="danielfortune")
        aplicar_filtro_dinamica(ws_fortune, "F1", valor_desejado="danielfortune")
        aplicar_filtro_dinamica(ws_fortune, "F2", valor_desejado="True")
        aplicar_filtro_dinamica(ws_fortune, "A4", valor_desejado="(Tudo)", exceto="(blank)")
        aplicar_filtro_dinamica(ws_fortune, "E4", valor_desejado="(Tudo)", exceto="(blank)")
        
        logger.info("Salvando arquivo de NC (Por favor, aguarde. Pode levar alguns segundos)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de NC concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de NC:")
        _fechar_excel_seguro(wb, excel)
        raise