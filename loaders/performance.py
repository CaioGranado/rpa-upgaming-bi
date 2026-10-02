"""
Auditoria e injeção da Base de Performance — 7 Steps que escrevem, cada um,
num bloco diferente de colunas da mesma aba 'BaseGeral' do mesmo arquivo
(arquivo_performance). Mantidos juntos neste único módulo porque não são 7
responsabilidades diferentes: são 7 facetas da mesma responsabilidade
(popular a BaseGeral), cada uma lendo de uma fonte de dados diferente
(Transações, NC, FTD, MTD, KYC, GeneralStats, UGS Diário) mas escrevendo
no mesmo destino. Separar em 7 arquivos esconderia a duplicação de padrão
entre os Steps (todos calculam idx_inicio/linha_excel_inicio do mesmo jeito),
tornando mais fácil corrigir um e esquecer os outros 6.

Nota sobre backup: só o Step 1 chama _fazer_backup(arquivo_performance) —
decisão deliberada da equipe (evitar 7 backups redundantes do mesmo arquivo
por execução; o backup do Step 1 já cobre o estado "antes de qualquer Step
rodar no dia").

Nota sobre DadosNaoConfiaveisError (Step 7): quando um dia não tem
registro na fonte (arquivo diário de UGS ausente/ilegível), a exceção é
levantada em vez de preencher com zero silenciosamente — "não sei" não
pode virar "é zero" sem confirmação real. O Step 6 (General Stats) não
valida mais nada aqui — a confiabilidade dos dados é responsabilidade
exclusiva da extração (extractors/web_scraper.py), que já garante, via
retry e utils/generalstats_utils.dia_generalstats_e_confiavel(), que só
dados confiáveis chegam a ser salvos no JSON. Um dia não confiável que
sobrevive à extração marca a marca inteira como incompleta, e o main.py
nem chega a chamar este Step 6 para ela.
"""
import calendar
import json
import logging
import re
from datetime import datetime
from typing import Final

import numpy as np
import pandas as pd
import win32com.client as win32

from transformers.data_cleaner import garantir_continuidade_temporal
from utils.date_utils import obter_data_alvo
from utils.excel_utils import _fechar_excel_seguro
from utils.exceptions_utils import DadosNaoConfiaveisError
from utils.file_utils import (
    MESES_PT,
    _fazer_backup,
    _obter_caminho_download,
    obter_caminho_base,
    obter_pasta_ugs_diario,
)
from utils.formatacao_utils import formatar_brl, formatar_int, formatar_num

logger = logging.getLogger(__name__)

# Textos e identificadores repetidos nos 7 steps (uma definição só, em vez de copiar o literal).
_ROTULO_LINHA: Final[str] = "Rótulos de Linha"
_PROGID_EXCEL: Final[str] = "Excel.Application"
_MSG_LENDO_HISTORICO: Final[str] = "Lendo histórico da 'BaseGeral' para auditoria..."
_MSG_SALVANDO: Final[str] = "Salvando Base Performance..."
_MSG_ERRO_GRAVACAO: Final[str] = "Erro crítico ao gravar a Base Performance:"


# Colunas com FÓRMULA da BaseGeral, arrastadas para o mês inteiro na abertura. Lista confirmada
# pela equipe: F:I, L:M, V:W, Y:AB, AE, AH, AK, AN, AZ:BE, BG, BI:BJ. N, AX e AY ficam FORA de
# propósito: são valores colados (N pelo Step 2, que grava N:O; AX e AY pelo Step 4, que grava AT:AY),
# não fórmulas. Qualquer coluna nova de fórmula
# precisa ser incluída aqui; a conferência da abertura (_conferir_formulas_do_mes) avisa se faltar.
_BLOCOS_FORMULAS_ABERTURA = ["F:I", "L:M", "V:W", "Y:AB", "AE:AE", "AH:AH", "AK:AK", "AN:AN", "AZ:BE", "BG:BG", "BI:BJ"]

# Última coluna conhecida da BaseGeral (BM): a mesma usada no zebrado da abertura de mês.
_ULTIMA_COLUNA_BASEGERAL = 65
# Colunas A:E (referência, data, ano, mês e dia) são escritas como VALOR na abertura: nunca são arrastadas.
_ULTIMA_COLUNA_FIXA = 5


def _numero_da_coluna(letras: str) -> int:
    numero = 0
    for letra in letras.upper():
        numero = numero * 26 + (ord(letra) - ord("A") + 1)
    return numero


def _letra_da_coluna(numero: int) -> str:
    letras = ""
    while numero:
        numero, resto = divmod(numero - 1, 26)
        letras = chr(ord("A") + resto) + letras
    return letras


def _colunas_do_bloco(bloco: str) -> set[int]:
    """'F:I' -> {6, 7, 8, 9}."""
    inicio, fim = bloco.split(":")
    return set(range(_numero_da_coluna(inicio), _numero_da_coluna(fim) + 1))


def _conferir_formulas_do_mes(ws, linha_referencia, linha_fim, colunas_arrastadas):
    """
    Confere, logo depois do FillDown dos blocos fixos, se o mês aberto ficou como
    esperado. NÃO arrasta nada por conta própria: a lista de colunas
    (_BLOCOS_FORMULAS_ABERTURA) é decisão da equipe; esta função só confere e avisa.

    1. Coluna com FÓRMULA na linha de referência (último dia do mês anterior) que não
       está na lista: só avisada ([VERIFICAR]). Se deveria ser arrastada, é preciso
       incluí-la na lista. Vale também para colunas depois da BM.
    2. Coluna da lista com VALOR (não fórmula) na referência: avisada ([VERIFICAR]);
       o FillDown copiou esse valor para o mês todo.
    3. Coluna da lista que tinha fórmula na referência e não tem no último dia do mês:
       levanta DadosNaoConfiaveisError antes do wb.Save(), então o arquivo não é salvo
       com o mês aberto pela metade.

    Devolve {"fora_da_lista": [...], "valor_na_lista": [...]} com as letras das colunas.
    """
    usado = ws.UsedRange
    ultima_coluna = max(_ULTIMA_COLUNA_BASEGERAL, usado.Column + usado.Columns.Count - 1)

    formulas_na_referencia = {
        coluna for coluna in range(_ULTIMA_COLUNA_FIXA + 1, ultima_coluna + 1)
        if ws.Cells(linha_referencia, coluna).HasFormula
    }

    fora_da_lista = sorted(formulas_na_referencia - colunas_arrastadas)
    valor_na_lista = sorted(colunas_arrastadas - formulas_na_referencia)
    esperadas = colunas_arrastadas & formulas_na_referencia
    sem_formula_no_fim = sorted(c for c in esperadas if not ws.Cells(linha_fim, c).HasFormula)

    def letras(colunas):
        return [_letra_da_coluna(c) for c in colunas]

    logger.info(
        f" -> Fórmulas do mês: {len(esperadas)} coluna(s) da lista confirmadas na linha de referência "
        f"({linha_referencia}) e presentes na linha {linha_fim}."
        if not sem_formula_no_fim else
        f" -> Fórmulas do mês: {len(esperadas)} coluna(s) da lista com fórmula na linha de referência ({linha_referencia})."
    )
    if fora_da_lista:
        logger.warning(
            f"[VERIFICAR] {letras(fora_da_lista)} têm FÓRMULA na linha {linha_referencia} mas NÃO estão em "
            f"blocos_formulas, então não foram arrastadas. Se deveriam, inclua-as na lista."
        )
    if valor_na_lista:
        logger.warning(
            f"[VERIFICAR] {letras(valor_na_lista)} estão em blocos_formulas, mas na linha {linha_referencia} têm "
            f"VALOR (não fórmula): o FillDown copiou esse valor para o mês todo. Confirme se é intencional."
        )
    if sem_formula_no_fim:
        raise DadosNaoConfiaveisError(
            f"Abertura de mês: as colunas {letras(sem_formula_no_fim)} deveriam ter fórmula na linha {linha_fim} "
            f"(último dia do mês) e não têm. O arquivo não foi salvo."
        )
    return {"fora_da_lista": letras(fora_da_lista), "valor_na_lista": letras(valor_na_lista)}


def carregar_base_performance_step1(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 1) ({marca}) ===")
    arquivo_transacoes = obter_caminho_base(marca, "Transacoes", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_transacoes.exists() or not arquivo_performance.exists():
        raise DadosNaoConfiaveisError("Arquivos necessários para a Base Performance não encontrados.")

    logger.info("Lendo Tabela Dinâmica 'Din_Diario' no arquivo de Transações...")
    df_din = pd.read_excel(arquivo_transacoes, sheet_name="Din_Diario", header=None)
    
    idx_header = df_din[df_din.apply(lambda r: r.astype(str).str.contains(_ROTULO_LINHA, case=False).any(), axis=1)].index
    
    if idx_header.empty:
        logger.error("Não foi possível encontrar 'Rótulos de Linha' na aba Din_Diario.")
        return
        
    linha_cabecalho = idx_header[0]
    df_din.columns = df_din.iloc[linha_cabecalho]
    df_din = df_din.iloc[linha_cabecalho + 1:].reset_index(drop=True)
    
    df_din = df_din[~df_din[_ROTULO_LINHA].astype(str).str.contains('Total|Vazio|NaN|nan', case=False, na=False)]
    
    col_deposito = next((c for c in df_din.columns if 'deposit' in str(c).lower()), None)
    col_saque = next((c for c in df_din.columns if 'withdraw' in str(c).lower()), None)
    if col_deposito is None or col_saque is None:
        raise DadosNaoConfiaveisError(
            f"Colunas de Deposit/Withdraw não encontradas na Tabela Dinâmica 'Din_Diario' ({marca})."
        )
    
    df_din['Dia'] = pd.to_numeric(df_din[_ROTULO_LINHA], errors='coerce').fillna(0).astype(int)
    df_din[col_deposito] = pd.to_numeric(df_din[col_deposito], errors='coerce').fillna(0)
    df_din[col_saque] = pd.to_numeric(df_din[col_saque], errors='coerce').fillna(0)
    
    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    
    df_din = df_din[(df_din['Dia'] > 0) & (df_din['Dia'] <= limite_dia)].copy()
    
    df_din = garantir_continuidade_temporal(df_din, limite_dia)

    logger.info(_MSG_LENDO_HISTORICO)
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,J,K", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], dayfirst=True, errors='coerce').dt.normalize()
    
    wb, excel = None, None
    try:
        _fazer_backup(arquivo_performance)
        excel = win32.DispatchEx(_PROGID_EXCEL)
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
# --- LÓGICA DE ABERTURA DE MÊS AUTOMÁTICA ---
        if not (datas_base == primeiro_dia_mes).any():
            logger.info("⚠️ MÊS NOVO DETECTADO! Iniciando abertura automática do mês na BaseGeral...")
            ultima_linha_preenchida = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            
            # A referência (Coluna A) agora ganha +1 e se mantém para o mês todo
            ref_anterior = int(ws.Cells(ultima_linha_preenchida, 1).Value)
            nova_ref_mes = ref_anterior + 1 
            
            dias_no_mes = calendar.monthrange(data_alvo.year, data_alvo.month)[1]
            nome_mes = MESES_PT[data_alvo.month][0]
            
            linha_inicio_nova = ultima_linha_preenchida + 1
            linha_fim_nova = linha_inicio_nova + dias_no_mes - 1
            
            logger.info(" -> Clonando formatação visual da última linha...")
            ws.Range(f"{ultima_linha_preenchida}:{ultima_linha_preenchida}").Copy()
            ws.Range(f"{linha_inicio_nova}:{linha_fim_nova}").PasteSpecial(Paste=-4122) # Apenas formato
            excel.CutCopyMode = False
            
            logger.info(f" -> Criando {dias_no_mes} novas linhas para {nome_mes}...")
            for offset in range(dias_no_mes):
                r = linha_inicio_nova + offset
                dia_numero = offset + 1
                
                ws.Cells(r, 1).Value = nova_ref_mes
                # Injeção de data nativa e segura
                ws.Cells(r, 2).Value = datetime(data_alvo.year, data_alvo.month, dia_numero) # noqa: DTZ001
                ws.Cells(r, 3).Value = data_alvo.year
                ws.Cells(r, 4).Value = nome_mes
                ws.Cells(r, 5).Value = dia_numero
            
            logger.info(" -> Arrastando todos os blocos de fórmulas estendidas...")
            # Lista modular: o FillDown traz a fórmula, mas sobrescreve a cor
            blocos_formulas = _BLOCOS_FORMULAS_ABERTURA
            colunas_arrastadas = set()
            for bloco in blocos_formulas:
                col_inicio, col_fim = bloco.split(':')
                ws.Range(f"{col_inicio}{ultima_linha_preenchida}:{col_fim}{linha_fim_nova}").FillDown()
                colunas_arrastadas |= _colunas_do_bloco(bloco)

            # Confere o resultado: fórmulas fora da lista, valores na lista e fórmula no fim do mês
            _conferir_formulas_do_mes(ws, ultima_linha_preenchida, linha_fim_nova, colunas_arrastadas)

            logger.info(" -> Alternando a cor de fundo (Zebrado Mensal) sobre as fórmulas...")
            cor_antiga = ws.Cells(ultima_linha_preenchida, 1).Interior.ColorIndex
            # Se o mês anterior era branco (2) ou sem cor (-4142), o novo vira cinza claro (15). Senão, tira a cor.
            nova_cor = 15 if cor_antiga in [2, -4142] else -4142
            ws.Range(f"A{linha_inicio_nova}:BM{linha_fim_nova}").Interior.ColorIndex = nova_cor

            wb.Save()
            logger.info("Mês aberto e formatado com sucesso! Prosseguindo com a injeção diária...")
            
            # Atalho inteligente
            idx_inicio_pandas = linha_inicio_nova - 2
            linha_excel_inicio = linha_inicio_nova 
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
        
        logger.info(_MSG_SALVANDO)
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 1 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception(_MSG_ERRO_GRAVACAO)
        _fechar_excel_seguro(wb, excel)
        raise

def carregar_base_performance_step2(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 2 - NC) ({marca}) ===")
    arquivo_nc = obter_caminho_base(marca, "NC", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_nc.exists() or not arquivo_performance.exists():
        raise DadosNaoConfiaveisError("Arquivos necessários para a Base Performance (Step 2) não encontrados.")

    logger.info("Lendo Tabelas Dinâmicas 'Din_Diario' no arquivo de Novas Contas...")
    df_din_full = pd.read_excel(arquivo_nc, sheet_name="Din_Diario", header=None)
    
    def extrair_dinamica(df_pedaco, nome_coluna_valor):
        mask = df_pedaco.apply(lambda r: r.astype(str).str.contains(_ROTULO_LINHA, case=False).any(), axis=1)
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
        raise DadosNaoConfiaveisError("Nenhum dia válido encontrado nas tabelas dinâmicas até o dia anterior.")
        
    df_mesclado = garantir_continuidade_temporal(df_mesclado, limite_dia)

    logger.info(_MSG_LENDO_HISTORICO)
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,N,O", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], dayfirst=True, errors='coerce').dt.normalize()
    
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
        excel = win32.DispatchEx(_PROGID_EXCEL)
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar) - 1
        
        logger.info(f"Sobrescrevendo colunas N (Orgânicos) e O (Total) das linhas {linha_excel_inicio} até {linha_excel_fim}...")
        ws.Range(ws.Cells(linha_excel_inicio, 14), ws.Cells(linha_excel_fim, 15)).Value = dados_para_injetar
        
        logger.info(_MSG_SALVANDO)
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 2 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception(_MSG_ERRO_GRAVACAO)
        _fechar_excel_seguro(wb, excel)
        raise

def carregar_base_performance_step3(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 3 - FTD) ({marca}) ===")
    arquivo_ftd = obter_caminho_base(marca, "FTD", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_ftd.exists() or not arquivo_performance.exists():
        raise DadosNaoConfiaveisError("Arquivos necessários para a Base Performance (Step 3) não encontrados.")

    def extrair_dinamica_bloco(df_source, col_idx_dia, col_idx_valores, col_names):
        mask = df_source.iloc[:, col_idx_dia].astype(str).str.contains(_ROTULO_LINHA, case=False)
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
        raise DadosNaoConfiaveisError("Nenhum dia válido encontrado no FTD até o dia anterior.")
        
    df_mesclado = garantir_continuidade_temporal(df_mesclado, limite_dia)

    logger.info(_MSG_LENDO_HISTORICO)
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,P:U,X", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], dayfirst=True, errors='coerce').dt.normalize()
    
    if not (datas_base == primeiro_dia_mes).any():
        logger.error(f"A data {primeiro_dia_mes.strftime('%d/%m/%Y')} não foi encontrada na coluna B da BaseGeral!")
        return
        
    idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio_pandas + 2 
    
    logger.info(f"Data {primeiro_dia_mes.strftime('%d/%m/%Y')} encontrada na linha {linha_excel_inicio} do Excel. Iniciando reconciliação...")
    
    dados_para_injetar_pu = []
    dados_para_injetar_x = []
    
    letras_pu = ['P', 'Q', 'R', 'S', 'T', 'U']
    
    for i, row in df_mesclado.iterrows():
        dia = int(row['Dia'])
        novos_pu = [row[col] for col in nomes_colunas_pu]
        novo_x = row['FTD_Org']
        
        dados_para_injetar_pu.append(novos_pu)
        dados_para_injetar_x.append([novo_x])
        
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
        excel = win32.DispatchEx(_PROGID_EXCEL)
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar_pu) - 1
        
        logger.info(f"Sobrescrevendo colunas P até U (linhas {linha_excel_inicio} até {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 16), ws.Cells(linha_excel_fim, 21)).Value = dados_para_injetar_pu
        
        logger.info(f"Sobrescrevendo coluna X (linhas {linha_excel_inicio} até {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 24), ws.Cells(linha_excel_fim, 24)).Value = dados_para_injetar_x
        
        logger.info(_MSG_SALVANDO)
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 3 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception(_MSG_ERRO_GRAVACAO)
        _fechar_excel_seguro(wb, excel)
        raise

def carregar_base_performance_step4(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 4 - MTD) ({marca}) ===")
    arquivo_mtd = obter_caminho_base(marca, "MTD", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_mtd.exists() or not arquivo_performance.exists():
        raise DadosNaoConfiaveisError("Arquivos necessários para a Base Performance (Step 4) não encontrados.")

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
        raise DadosNaoConfiaveisError("Nenhum dia válido encontrado no MTD após aplicar o filtro de data.")
        
    df_mtd = garantir_continuidade_temporal(df_mtd, limite_dia)

    logger.info(_MSG_LENDO_HISTORICO)
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,AT:AY", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], dayfirst=True, errors='coerce').dt.normalize()
    
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
        excel = win32.DispatchEx(_PROGID_EXCEL)
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_para_injetar) - 1
        
        logger.info(f"Sobrescrevendo colunas AT até AY (linhas {linha_excel_inicio} até {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 46), ws.Cells(linha_excel_fim, 51)).Value = dados_para_injetar
        
        logger.info(_MSG_SALVANDO)
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 4 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception(_MSG_ERRO_GRAVACAO)
        _fechar_excel_seguro(wb, excel)
        raise

# Faixas de tempo entre cadastro e FTD (3ª dinâmica da aba DIN do KYC) -> coluna da BaseGeral.
# Faixa 1 -> BK, Faixa 2 -> BL, Faixa 3 -> BM.
_FAIXAS_FTD_KYC: Final[dict[int, str]] = {1: 'Col_BK', 2: 'Col_BL', 3: 'Col_BM'}


def _extrair_faixas_ftd_kyc(df_source: pd.DataFrame, col_idx_dia: int) -> pd.DataFrame:
    """
    Lê a dinâmica Dia x Faixa de FTD (aba DIN do KYC) identificando cada coluna pelo
    NOME do cabeçalho ('Faixa N ...'), nunca pela posição.

    A dinâmica só cria coluna para a faixa que tem dado: sem ninguém na Faixa 2, vêm
    apenas Faixa 1 e Faixa 3, lado a lado. Ler por posição quebraria (IndexError) ou,
    pior, jogaria o valor da Faixa 3 na coluna da Faixa 2. Regra de negócio confirmada:
    faixa ausente = 0 (nenhum FTD naquela faixa).

    Levanta DadosNaoConfiaveisError se aparecer um cabeçalho que não seja 'Faixa 1/2/3'
    (nem 'Total Geral'), ou se a mesma faixa aparecer duas vezes: o layout mudou, e
    adivinhar a coluna poderia gravar número no lugar errado.
    """
    mask = df_source.iloc[:, col_idx_dia].astype(str).str.contains('Rótulos|Dia', case=False)
    if not mask.any():
        return pd.DataFrame()

    idx_cabecalho = df_source[mask].index[0]
    colunas_faixa = {}  # número da faixa -> posição da coluna no df

    for pos in range(col_idx_dia + 1, df_source.shape[1]):
        texto = df_source.iloc[idx_cabecalho, pos]
        if pd.isna(texto) or not str(texto).strip():
            break  # coluna vazia = fim da dinâmica
        texto = str(texto).strip()
        if texto.lower().startswith('total geral'):
            break
        m = re.match(r'Faixa\s*([123])\b', texto, re.IGNORECASE)
        if not m or int(m.group(1)) in colunas_faixa:
            raise DadosNaoConfiaveisError(
                f"Dinâmica de faixas de FTD (KYC/DIN) com cabeçalho inesperado na coluna nº "
                f"{pos + 1}: '{texto}'. Esperado 'Faixa 1', 'Faixa 2' ou 'Faixa 3', sem repetição. "
                f"O layout mudou e não é seguro mapear para BK/BL/BM."
            )
        colunas_faixa[int(m.group(1))] = pos

    corpo = df_source.iloc[idx_cabecalho + 1:]
    df = pd.DataFrame({'Dia': pd.to_numeric(corpo.iloc[:, col_idx_dia], errors='coerce')})

    for faixa, nome_col in _FAIXAS_FTD_KYC.items():
        if faixa in colunas_faixa:
            df[nome_col] = pd.to_numeric(corpo.iloc[:, colunas_faixa[faixa]], errors='coerce')
        else:
            logger.info(
                f"Faixa {faixa} ausente na dinâmica do KYC (nenhum FTD nessa faixa no período). "
                f"Preenchida com 0 ({nome_col})."
            )
            df[nome_col] = 0

    df = df.dropna(subset=['Dia'])
    df['Dia'] = df['Dia'].astype(int)
    df = df[df['Dia'] > 0].copy()

    for nome_col in _FAIXAS_FTD_KYC.values():
        df[nome_col] = df[nome_col].fillna(0).astype(int)

    return df[['Dia', *_FAIXAS_FTD_KYC.values()]]


def carregar_base_performance_step5(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 5 - KYC) ({marca}) ===")
    arquivo_kyc = obter_caminho_base(marca, "KYC", obter_data_alvo())
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_kyc.exists() or not arquivo_performance.exists():
        raise DadosNaoConfiaveisError("Arquivos necessários para a Base Performance (Step 5) não encontrados.")

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
    df_p3 = _extrair_faixas_ftd_kyc(df_din_full, 8)  # faixas lidas por nome de cabeçalho, ausente = 0

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
        raise DadosNaoConfiaveisError("Nenhum dia válido encontrado no KYC após aplicar o filtro de data.")
        
    df_mesclado = garantir_continuidade_temporal(df_mesclado, limite_dia)

    logger.info(_MSG_LENDO_HISTORICO)
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,BF,BH,BK:BM", header=0)
    
    data_alvo = obter_data_alvo()
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], dayfirst=True, errors='coerce').dt.normalize()
    
    if not (datas_base == primeiro_dia_mes).any():
        logger.error(f"A data {primeiro_dia_mes.strftime('%d/%m/%Y')} não foi encontrada na coluna B da BaseGeral!")
        return
        
    idx_inicio_pandas = datas_base[datas_base == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio_pandas + 2 
    
    logger.info(f"Data {primeiro_dia_mes.strftime('%d/%m/%Y')} encontrada na linha {linha_excel_inicio}. Iniciando reconciliação...")
    
    dados_bf = []
    dados_bh = []
    dados_bk_bm = []
    
    for i, row in df_mesclado.iterrows():
        dia = int(row['Dia'])
        novo_bf = int(row['KYC_True'])
        novo_bh = int(row['KYC_FTD_True'])
        novos_bk_bm = [int(row['Col_BK']), int(row['Col_BL']), int(row['Col_BM'])]
        
        dados_bf.append([novo_bf])
        dados_bh.append([novo_bh])
        dados_bk_bm.append(novos_bk_bm)
        
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
        excel = win32.DispatchEx(_PROGID_EXCEL)
        excel.Visible = True
        excel.DisplayAlerts = False
        
        logger.info(f"Abrindo {arquivo_performance.name} para injetar os dados validados...")
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_bf) - 1
        
        logger.info(f"Sobrescrevendo coluna BF (linha {linha_excel_inicio} a {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 58), ws.Cells(linha_excel_fim, 58)).Value = dados_bf
        
        logger.info(f"Sobrescrevendo coluna BH (linha {linha_excel_inicio} a {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 60), ws.Cells(linha_excel_fim, 60)).Value = dados_bh
        
        logger.info(f"Sobrescrevendo colunas BK até BM (linha {linha_excel_inicio} a {linha_excel_fim})...")
        ws.Range(ws.Cells(linha_excel_inicio, 63), ws.Cells(linha_excel_fim, 65)).Value = dados_bk_bm
        
        logger.info(_MSG_SALVANDO)
        wb.Save()
        wb.Close()
        excel.Quit()
        logger.info("-> Step 5 da Base Performance concluído com sucesso!")
        
    except Exception:
        logger.exception(_MSG_ERRO_GRAVACAO)
        _fechar_excel_seguro(wb, excel)
        raise

def carregar_base_performance_step6(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 6 - GEN STATS) ({marca}) ===")
    arquivo_json = _obter_caminho_download(marca, f"GeneralStats - {marca}.json")
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_json or not arquivo_json.exists() or not arquivo_performance.exists():
        raise DadosNaoConfiaveisError("Arquivos necessários para a Base Performance (Step 6) não encontrados.")

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
    
    # Nota: este Step não valida a confiabilidade dos dados — isso é
    # responsabilidade exclusiva da extração (extractors/web_scraper.py),
    # que só grava um dia no JSON depois de confirmar (com retry) que ele
    # passa em utils/generalstats_utils.dia_generalstats_e_confiavel(). Se
    # um dia não confiável sobrevive à extração, a marca inteira já foi
    # marcada como incompleta e o main.py nem chega a chamar este Step.
    # Loader é loader: confia no que recebe, só injeta.
    logger.info(f"Processando dados do dia 01 até {limite_dia:02d}...")
    for dia in range(1, limite_dia + 1):
        json_do_dia = dados_por_dia.get(dia, [])

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
    datas_base = pd.to_datetime(df_base.iloc[:, 0], dayfirst=True, errors='coerce').dt.normalize()
    idx_inicio = datas_base[datas_base == primeiro_dia_mes].index[0]
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
        excel = win32.DispatchEx(_PROGID_EXCEL)
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
        raise

def carregar_base_performance_step7(marca, *args, **kwargs):
    logger.info(f"=== INICIANDO AUDITORIA E INJEÇÃO: BASE PERFORMANCE (STEP 7 - USUÁRIOS ÚNICOS) ({marca}) ===")
    arquivo_performance = obter_caminho_base(marca, "Performance", obter_data_alvo())

    if not arquivo_performance.exists():
        raise DadosNaoConfiaveisError("Base Performance não encontrada para o Step 7.")

    data_alvo = obter_data_alvo()
    limite_dia = data_alvo.day
    mes_atual = data_alvo.month
    
    df_base = pd.read_excel(arquivo_performance, sheet_name="BaseGeral", usecols="B,AS", header=0)
    primeiro_dia_mes = pd.Timestamp(year=data_alvo.year, month=data_alvo.month, day=1)
    datas_base = pd.to_datetime(df_base.iloc[:, 0], dayfirst=True, errors='coerce').dt.normalize()
    idx_inicio = datas_base[datas_base == primeiro_dia_mes].index[0]
    linha_excel_inicio = idx_inicio + 2 

    dados_as = []
    pasta_ugs = obter_pasta_ugs_diario(marca, data_alvo.year, data_alvo.month)
    logger.info(f"Lendo e contando os arquivos diários do UGS (01 até {limite_dia:02d})...")
    
    for dia in range(1, limite_dia + 1):
        nome_arquivo = f"{dia:02d}-{mes_atual:02d}.xlsx"
        caminho_ugs = pasta_ugs / nome_arquivo
        
        if not caminho_ugs.exists():
            raise DadosNaoConfiaveisError(
                f"Arquivo diário {nome_arquivo} não encontrado. Não é seguro "
                f"preencher Usuários Únicos com 0 sem confirmar a ausência real do dado."
            )

        try:
            df_ugs = pd.read_excel(caminho_ugs, engine='calamine')
        except Exception as e:
            raise DadosNaoConfiaveisError(
                f"Arquivo diário {nome_arquivo} não pôde ser lido ({e}). Não é seguro "
                f"preencher Usuários Únicos com 0 sem confirmar o valor real."
            ) from e

        qtd_usuarios = len(df_ugs)

        dados_as.append([qtd_usuarios])
        
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
        excel = win32.DispatchEx(_PROGID_EXCEL)
        excel.Visible = True; excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(arquivo_performance), UpdateLinks=0)
        ws = wb.Sheets("BaseGeral")
        
        linha_excel_fim = linha_excel_inicio + len(dados_as) - 1
        logger.info(f"Sobrescrevendo coluna AS (linhas {linha_excel_inicio} a {linha_excel_fim})...")
        
        ws.Range(ws.Cells(linha_excel_inicio, 45), ws.Cells(linha_excel_fim, 45)).Value = dados_as
        
        wb.Save(); wb.Close(); excel.Quit()
        logger.info("-> Step 7 (Usuários Únicos) Concluído com Sucesso!")
    except Exception:
        logger.exception("Erro crítico na gravação do Step 7:")
        _fechar_excel_seguro(wb, excel)
        raise