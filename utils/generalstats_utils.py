"""
Regra de negócio para validar a confiabilidade dos dados de General
Statistics retornados pela API do BackOffice, por dia.

Compartilhado entre dois consumidores com propósitos diferentes:
- extractors/web_scraper.py: decide, durante a extração (com o navegador
  ainda autenticado), se vale tentar o dia de novo antes de desistir.
- loaders/performance.py (Step 6): última linha de defesa antes de
  escrever valores financeiros no Excel, independente de quando/como o
  JSON foi produzido (inclusive se rodado via testar_injecao.py, sem
  passar pela extração de hoje).

Centralizado aqui para que as duas pontas nunca divirjam sobre o que
conta como "dia confiável".
"""
from typing import Final

# Verticais que, por regra de negócio confirmada (2 anos de dados
# ininterruptos nas 3 marcas, exceto queda de servidor ou parada
# programada), NUNCA ficam com aposta zerada num dia normal. Se vier
# ausente ou zerada, é sinal de falha silenciosa na extração — não de
# atividade real zero — mesmo que outras verticais do mesmo dia pareçam
# normais.
VERTICAIS_SEMPRE_ATIVAS: Final[set[str]] = {"Slot"}


def _valor_numerico(jogo: dict, campo: str) -> float:
    """Converte um campo do JSON (às vezes vem como string) para float, com fallback seguro para 0."""
    try:
        return float(jogo.get(campo, 0) or 0)
    except (TypeError, ValueError):
        return 0.0


def dia_generalstats_e_confiavel(json_do_dia: list) -> tuple[bool, str]:
    """
    Avalia se os dados de um dia de General Statistics são confiáveis.

    Retorna (confiavel, motivo) — motivo é "" quando confiavel=True, ou
    uma mensagem explicando a falha caso contrário (pronta para uso
    direto em log ou exceção).

    Duas camadas de checagem:
    1. Falha ampla: nenhum valor real (aposta/ganho/usuários) em nenhuma
       vertical, mesmo com registros presentes.
    2. Falha cirúrgica: uma vertical que SEMPRE tem atividade (ex: Slot)
       veio ausente ou com aposta zerada, mesmo que outras verticais do
       mesmo dia pareçam normais (por isso separada da checagem 1 — uma
       falha parcial não necessariamente zera o dia inteiro).
    """
    if not json_do_dia:
        return False, "nenhum registro retornado pela API"

    valores_bet_por_vertical = {}
    algum_valor_real = False
    for jogo in json_do_dia:
        tipo = jogo.get("gameType")
        bet = _valor_numerico(jogo, "betAmount")
        win = _valor_numerico(jogo, "winAmount")
        users = _valor_numerico(jogo, "userCount")
        valores_bet_por_vertical[tipo] = bet
        if bet or win or users:
            algum_valor_real = True

    if not algum_valor_real:
        return False, (
            f"{len(json_do_dia)} registro(s) retornados, mas aposta, ganho e "
            f"usuários estão zerados em todos os tipos de jogo"
        )

    for vertical in VERTICAIS_SEMPRE_ATIVAS:
        if not valores_bet_por_vertical.get(vertical, 0):
            return False, (
                f"vertical '{vertical}' veio ausente ou com aposta zerada "
                f"(essa vertical é esperada em toda marca, todo dia)"
            )

    return True, ""