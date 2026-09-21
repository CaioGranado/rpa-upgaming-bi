import logging
from datetime import datetime, timedelta, timezone

logger = logging.getLogger(__name__)

# Variável global para guardar a decisão do usuário durante a execução
_DATA_ALVO_CACHE = None

def obter_data_alvo():
    """Retorna a data alvo (D-1). Pergunta sobre fechamento se for início do mês."""
    global _DATA_ALVO_CACHE
    if _DATA_ALVO_CACHE:
        return _DATA_ALVO_CACHE

    hoje = datetime.now(timezone.utc).astimezone()
    ontem = hoje - timedelta(days=1)
    
    # Se estamos nos primeiros 5 dias do mês, pergunta sobre o fechamento!
    if hoje.day <= 5:
        print(f"\n[ATENÇÃO] Hoje é dia {hoje.strftime('%d/%m/%Y')}.")
        resposta = input("Deseja rodar o FECHAMENTO do mês anterior? (S/N): ").strip().upper()
        
        if resposta == 'S':
            # Pega o primeiro dia do mês atual e volta 1 dia = Último dia do mês anterior
            ultimo_dia_mes_anterior = hoje.replace(day=1) - timedelta(days=1)
            _DATA_ALVO_CACHE = ultimo_dia_mes_anterior
            logger.info(f"Modo Fechamento Ativado! Data Alvo configurada para: {_DATA_ALVO_CACHE.strftime('%d/%m/%Y')}")
            return _DATA_ALVO_CACHE

    # Fluxo normal (D-1)
    _DATA_ALVO_CACHE = ontem
    return _DATA_ALVO_CACHE

def calcular_limite_seguro(data_referencia: datetime) -> datetime:
    """
    Retorna o início do dia SEGUINTE à data de referência (00:00:00).

    Motivo: os filtros de data do BackOffice têm granularidade de minuto
    (dd-mm-aaaa hh:mm) e cortam no segundo exato. Usar "23:59" do próprio
    dia como limite superior descarta qualquer registro entre 23:59:01 e
    23:59:59. Usar o início do dia seguinte como limite (respeitando que
    o próprio filtro do BO corta exatamente em HH:MM:00) garante a
    captura de 100% do dia de referência, sem sobrepor dados do dia
    seguinte.

    Não usar para o relatório de Novas Contas: a janela de NC (data_fim_nc)
    já é aberta para o futuro (usa 'hoje', sempre além do momento real de
    execução) e o corte de fechamento do dia é feito depois, via datetime
    completo, em aplicar_corte_datas_futuras() — não há truncamento a corrigir ali.
    """
    return data_referencia.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)

def obter_periodo_extracao():
    alvo = obter_data_alvo()
    hoje_real = datetime.now(timezone.utc).astimezone()
    data_inicio = alvo.replace(day=1, hour=0, minute=0, second=0, microsecond=0).strftime("%d-%m-%Y %H:%M")
    # Limite seguro (dia seguinte às 00:00) para não perder o último minuto do dia alvo.
    # Usado por: System Transactions, UGS Acumulado e FTD. NC continua de fora (ver docstring acima).
    data_fim = calcular_limite_seguro(alvo).strftime("%d-%m-%Y %H:%M")
    data_fim_nc = hoje_real.replace(hour=23, minute=59, second=0, microsecond=0).strftime("%d-%m-%Y %H:%M")
    
    return data_inicio, data_fim, data_fim_nc