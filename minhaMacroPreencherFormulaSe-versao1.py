#!/usr/bin/env python3

import uno

def main():
    preencher_formulas_condicionais()
    return

def preencher_formulas_condicionais():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    for i in range(65, 91):  # ASCII de 'A' a 'Z'
        celula_destino = planilha.getCellByPosition(i - 65, 3)  # Linha 4, coluna correspondente
        formula = '=IF('+chr(i)+'3=0;"É MULTIPLO DE 10";"Valor não é múltiplo de 10")'  # Exemplo: '=SE(A3=0;"É MULTIPLO DE 10";"Valor não é múltiplo de 10")'
        celula_destino.Formula = formula

    for i in range(65, 91):  # Continuação das letras maiúsculas
        for j in range(65, 91):  # ASCII de 'A' a 'Z'
            coluna_atual = (i - 65 + 1) * 26 + (j - 65)  # Calcula a coluna atual
            if coluna_atual > 99:  # Limite de colunas
                break
            celula_destino = planilha.getCellByPosition(coluna_atual, 3)  # Linha 4, coluna correspondente
            formula = '=IF('+chr(i)+chr(j)+'3=0;"É MULTIPLO DE 10";"Valor não é múltiplo de 10")'  # Exemplo: '=SE(AA3=0;"É MULTIPLO DE 10";"Valor não é múltiplo de 10")'
            celula_destino.Formula = formula

    return
