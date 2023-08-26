#!/usr/bin/env python3

import uno

def main():
    preencher_cabecalhos()
    return

def preencher_cabecalhos():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    coluna_atual = 0
    for i in range(65, 91):  # ASCII de 'A' a 'Z'
        celula = planilha.getCellByPosition(coluna_atual, 0)  # Linha 1
        celula.String = chr(i)
        coluna_atual += 1

    for i in range(65, 91):  # Segunda parte das letras maiúsculas
        for j in range(65, 91):  # ASCII de 'A' a 'Z'
            celula = planilha.getCellByPosition(coluna_atual, 0)  # Linha 1
            celula.String = chr(i) + chr(j)
            coluna_atual += 1

    return
