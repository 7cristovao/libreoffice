#!/usr/bin/env python3

import uno

def main():
    preencher_letras()
    return

def preencher_letras():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    for i in range(65, 91):  # ASCII de 'A' a 'Z'
        coluna = i - 65  # Começando da coluna A (ASCII 65)
        celula = planilha.getCellByPosition(coluna, 0)  # Linha 1
        celula.String = chr(i)

    for i in range(97, 123):  # ASCII de 'a' a 'z'
        coluna = i - 97 + 26  # Continuando após as letras maiúsculas
        celula = planilha.getCellByPosition(coluna, 0)  # Linha 1
        celula.String = chr(i)

    return
