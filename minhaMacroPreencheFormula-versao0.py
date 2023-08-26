#!/usr/bin/env python3

import uno

def main():
    preencher_formulas()
    return

def preencher_formulas():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    for i in range(65, 91):  # ASCII de 'A' a 'Z'
        celula_destino = planilha.getCellByPosition(i - 65, 1)  # Linha 2, coluna correspondente
        formula = "="+chr(i)+"1"  # Exemplo: "=A1"
        celula_destino.Formula = formula

    for i in range(65, 91):  # Segunda parte das letras maiúsculas
        for j in range(65, 91):  # ASCII de 'A' a 'Z'
            coluna_atual = (i - 65 + 1) * 26 + (j - 65)  # Calcula a coluna atual
            celula_destino = planilha.getCellByPosition(coluna_atual - 1, 1)  # Linha 2, coluna correspondente
            formula = "="+chr(i)+chr(j)+"1"  # Exemplo: "=AA1"
            celula_destino.Formula = formula

    return
