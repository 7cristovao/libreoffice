#!/usr/bin/env python3

import uno

def main():
    preencher_formulas()
    return

def preencher_formulas():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    for i in range(65, 91):  # ASCII de 'A' a 'Z'
        celula_destino = planilha.getCellByPosition(i - 65, 2)  # Linha 3, coluna correspondente
        formula = "=MOD("+chr(i)+"1;10)"  # Exemplo: "=MOD(A1;10)"
        celula_destino.Formula = formula

    for i in range(65, 91):  # Continuação das letras maiúsculas
        for j in range(65, 91):  # ASCII de 'A' a 'Z'
            coluna_atual = (i - 65 + 1) * 26 + (j - 65)  # Calcula a coluna atual
            if coluna_atual > 99:  # Limite de colunas
                break
            celula_destino = planilha.getCellByPosition(coluna_atual, 2)  # Linha 3, coluna correspondente
            formula = "=MOD("+chr(i)+chr(j)+"1;10)"  # Exemplo: "=MOD(AA1;10)"
            celula_destino.Formula = formula

    return
