class FormatoInvalidoError(Exception):
    """
    Levantada quando um arquivo baixado do BackOffice não corresponde ao
    formato esperado (xlsx) — por exemplo, quando o endpoint de export
    devolve JSON paginado em vez do relatório completo.

    Tratada de forma especial no loop de extração: interrompe o pipeline
    da marca afetada sem impedir a execução das demais marcas.
    """