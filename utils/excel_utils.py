import logging

import pywintypes

logger = logging.getLogger(__name__)

def atualizar_dinamicas(wb, excel):
    logger.info("Sincronizando Tabelas Dinâmicas (RefreshAll)...")
    for cache in wb.PivotCaches():
        try:
            cache.BackgroundQuery = False
        except pywintypes.com_error as e:
            logger.warning(f"Propriedade BackgroundQuery não suportada neste cache: {e}")

        try:
            cache.MissingItemsLimit = 0
        except pywintypes.com_error as e:
            # Normal: MissingItemsLimit só existe para caches baseados em OLAP.
            # Tabelas dinâmicas comuns (baseadas em intervalo de dados) sempre
            # levantam esse erro aqui — não é sinal de dado não confiável.
            logger.warning(f"Propriedade MissingItemsLimit não suportada neste cache (normal para caches não-OLAP): {e}")

    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()
    logger.info("Tabelas Dinâmicas sincronizadas com sucesso.")

def _fechar_excel_seguro(wb, excel):
    try:
        if wb: wb.Close(SaveChanges=False)
        if excel: excel.Quit()
    except pywintypes.com_error as e_quit:
        logger.debug(f"Erro na interface COM ao forçar fechamento seguro: {e_quit}")
    except AttributeError as e_attr:
        logger.debug(f"Objeto inexistente durante o fechamento seguro: {e_attr}")

def aplicar_filtro_dinamica(sheet, celula, valor_desejado, fallback=None, exceto=None):
    try:
        rng = sheet.Range(celula)
        try:
            pt = rng.PivotTable
        except Exception:
            logger.error(f"Nenhuma Tabela Dinâmica associada à célula {celula} em {sheet.Name}.")
            return

        pivot_field = None
        
        # Estratégia 1: O campo está na própria célula (comum em Rótulos de Linha/Coluna)
        try:
            pivot_field = rng.PivotField
            logger.debug(f"[{celula}] Campo encontrado via Estratégia 1 (PivotField direto): '{pivot_field.Name}'")
        except Exception:
            logger.debug(f"[{celula}] Estratégia 1 não aplicável (célula não é PivotField direto), tentando Estratégia 2...")
            
        # Estratégia 2: Filtro de Página clássico (Nome do campo fica na célula à esquerda)
        if pivot_field is None and rng.Column > 1:
            try:
                nome_campo = str(sheet.Cells(rng.Row, rng.Column - 1).Value).strip()
                pivot_field = pt.PivotFields(nome_campo)
                logger.debug(f"[{celula}] Campo encontrado via Estratégia 2 (célula à esquerda): '{nome_campo}'")
            except Exception:
                logger.debug(f"[{celula}] Estratégia 2 não aplicável (célula à esquerda não mapeou campo), tentando Estratégia 3...")
                
        # Estratégia 3: Fallback para "Rótulos de Linha" compactos (pega o campo de linha principal)
        if pivot_field is None:
            valor_texto = str(rng.Value).strip()
            if "Rótulo" in valor_texto or "Row" in valor_texto:
                try:
                    pivot_field = pt.RowFields(1)
                    logger.debug(f"[{celula}] Campo encontrado via Estratégia 3 (RowFields — valor célula: '{valor_texto}')")
                except Exception:
                    logger.debug(f"[{celula}] Estratégia 3 (RowFields) falhou — valor célula: '{valor_texto}'")
            elif "Coluna" in valor_texto or "Column" in valor_texto:
                try:
                    pivot_field = pt.ColumnFields(1)
                    logger.debug(f"[{celula}] Campo encontrado via Estratégia 3 (ColumnFields — valor célula: '{valor_texto}')")
                except Exception:
                    logger.debug(f"[{celula}] Estratégia 3 (ColumnFields) falhou — valor célula: '{valor_texto}'")

        # Se depois de tudo não achar o campo, aborta com elegância
        if pivot_field is None:
            logger.error(f"Não foi possível mapear o Campo da Dinâmica na célula {celula}.")
            return

        # 1. Cenário: Exclusão (Selecionar Tudo, exceto X)
        if exceto:
            pivot_field.ClearAllFilters()
            try: pivot_field.EnableMultiplePageItems = True
            except: pass

            item_encontrado = False
            for item in pivot_field.PivotItems():
                if str(item.Name).strip().lower() == str(exceto).strip().lower():
                    try:
                        item.Visible = False
                        item_encontrado = True
                    except Exception as e:
                        logger.debug(f"Falha interna ao ocultar '{item.Name}': {e}")
            
            if not item_encontrado:
                logger.debug(f"Item para exclusão '{exceto}' não encontrado. (Tudo) mantido.")
                logger.info(
                    f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                    f"Modo='Exclusão' Item_solicitado='{exceto}' Resultado='Não encontrado, (Tudo) mantido'"
                )
            else:
                logger.info(
                    f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                    f"Modo='Exclusão' Item_ocultado='{exceto}'"
                )
            return

        # 2. Cenário: Múltipla Seleção (Lista de itens)
        if isinstance(valor_desejado, list):
            pivot_field.ClearAllFilters()
            try: pivot_field.EnableMultiplePageItems = True
            except: pass

            itens_ocultados = []
            for item in pivot_field.PivotItems():
                if str(item.Name).strip() not in valor_desejado:
                    try:
                        item.Visible = False
                        itens_ocultados.append(str(item.Name).strip())
                    except Exception as e:
                        logger.debug(f"Falha ao ocultar item não desejado '{item.Name}': {e}")

            logger.info(
                f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                f"Modo='Múltipla Seleção' Itens_mantidos={valor_desejado} Itens_ocultados={itens_ocultados}"
            )
            return

        # 3. Cenário: Resetar (Tudo)
        if valor_desejado == "(Tudo)":
            pivot_field.ClearAllFilters()
            logger.info(
                f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                f"Modo='Reset' Resultado='(Tudo) selecionado'"
            )
            return
        
        # 4. Cenário Padrão: Seleção Única com Fallback
        pivot_field.ClearAllFilters()
        try:
            pivot_field.CurrentPage = valor_desejado
            logger.info(
                f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                f"Modo='Seleção Única' Valor_aplicado='{valor_desejado}' Origem='Solicitado'"
            )
        except Exception:
            if fallback:
                logger.warning(f"Filtro '{valor_desejado}' não encontrado em {sheet.Name}!{celula}")
                if fallback == "(Tudo)":
                    pivot_field.ClearAllFilters()
                    logger.info(
                        f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                        f"Modo='Seleção Única' Valor_solicitado='{valor_desejado}' Valor_aplicado='(Tudo)' Origem='Fallback'"
                    )
                else:
                    try:
                        pivot_field.CurrentPage = fallback
                        logger.info(
                            f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                            f"Modo='Seleção Única' Valor_solicitado='{valor_desejado}' Valor_aplicado='{fallback}' Origem='Fallback'"
                        )
                    except Exception:
                        logger.warning(f"Fallback '{fallback}' também falhou. Mantendo aba sem filtros.")
                        pivot_field.ClearAllFilters()
                        logger.info(
                            f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                            f"Modo='Seleção Única' Valor_solicitado='{valor_desejado}' Fallback_solicitado='{fallback}' "
                            f"Valor_aplicado='(Tudo)' Origem='Fallback falhou, reset aplicado'"
                        )
            else:
                logger.warning(f"Filtro '{valor_desejado}' ausente em {sheet.Name}!{celula}. Nenhum fallback definido, ignorando")
                pivot_field.ClearAllFilters()
                logger.info(
                    f"[FILTRO APLICADO] Aba='{sheet.Name}' Célula='{celula}' Campo='{pivot_field.Name}' "
                    f"Modo='Seleção Única' Valor_solicitado='{valor_desejado}' Valor_aplicado='(Tudo)' "
                    f"Origem='Sem fallback definido, reset aplicado'"
                )

    except Exception as e:
        nome_aba = getattr(sheet, "Name", "Desconhecido")
        logger.error(f"Erro estrutural ao manipular filtro na aba {nome_aba}, célula {celula}: {e}")