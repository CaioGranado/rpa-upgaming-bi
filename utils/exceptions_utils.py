class FormatoInvalidoError(Exception):
    """
    Levantada quando um arquivo baixado do BackOffice não corresponde ao
    formato esperado (xlsx) — por exemplo, quando o endpoint de export
    devolve JSON paginado em vez do relatório completo.

    Tratada de forma especial no loop de extração: interrompe o pipeline
    da marca afetada sem impedir a execução das demais marcas.
    """


class DadosNaoConfiaveisError(Exception):
    """
    Levantada quando um dado necessário para auditoria/injeção não pode ser
    confirmado como real (arquivo ausente, ilegível, ou dia sem registro na
    fonte) — em vez de preencher silenciosamente com zero, o que tornaria
    um "não sei" indistinguível de uma medição real de valor zero.

    Deixada subir naturalmente até o try/except por marca no main.py, que
    interrompe os passos restantes dessa marca sem afetar as demais.
    """