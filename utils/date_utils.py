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

def obter_periodo_extracao():
    alvo = obter_data_alvo()
    hoje_real = datetime.now(timezone.utc).astimezone()
    data_inicio = alvo.replace(day=1).strftime("%d-%m-%Y 00:00")
    data_fim = alvo.strftime("%d-%m-%Y 23:59")
    data_fim_nc = hoje_real.strftime("%d-%m-%Y 23:59")
    
    return data_inicio, data_fim, data_fim_nc