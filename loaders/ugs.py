"""
Carregamento do relatório de User Game Statistics (UGS Acumulado, 5 abas:
Completo, Slot, LiveCassino, Sportsbook, MiniGames) na base oficial da marca.
"""
import logging

import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import blindar_dados
from utils.date_utils import calcular_limite_seguro, obter_data_alvo
from utils.excel_utils import (
    _fechar_excel_seguro,
    aplicar_filtro_dinamica,
    atualizar_dinamicas,
)
from utils.file_utils import _fazer_backup, _obter_caminho_download, obter_caminho_base

logger = logging.getLogger(__name__)


def carregar_base_ugs(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO CARREGAMENTO: UGS ({marca}) ===")
    data_alvo = obter_data_alvo()
    arquivo_base_oficial = obter_caminho_base(marca, "UGS", data_alvo)
    
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

        atualizar_dinamicas(wb, excel)

        ws_din_completo = wb.Sheets("Din_Completo")
        aplicar_filtro_dinamica(ws_din_completo, "A3", valor_desejado="(Tudo)", exceto="(blank)")

        data_inicio = data_alvo.replace(day=1)
        str_carimbo = f"dados alimentados de {data_inicio.strftime('%d/%m/%Y 00:00')} até {calcular_limite_seguro(data_alvo).strftime('%d/%m/%Y %H:%M')}"

        for ws_aba in wb.Sheets:
                if ws_aba.Name not in ["Din_Completo", "|"]:
                    ultima_linha =  ws_aba.Cells(ws_aba.Rows.Count, "A").End(-4162).Row
                    linha_carimbo = ultima_linha + 2
                    cel = ws_aba.Cells(linha_carimbo, 1)
                    cel.Value = str_carimbo
                    cel.Font.Italic = True
        
        logger.info("Salvando arquivo de UGS (Paciência, quase lá!)...")
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Carregamento de UGS concluído com sucesso!")
        
    except Exception:
        logger.exception("Erro crítico no carregamento de UGS:")
        _fechar_excel_seguro(wb, excel)
        raise