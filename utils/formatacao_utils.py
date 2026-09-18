"""
Funções puras de formatação de números para exibição em log/auditoria
(sem nenhuma dependência de COM/Excel — só Python padrão + pandas).
"""
import pandas as pd


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