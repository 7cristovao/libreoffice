#!/usr/bin/env python3

import uno

def main():
    preencher_numeros()
    return

def preencher_numeros():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    for i in range(1, 101):
        coluna = i - 1  # A coluna começa em 0
        celula = planilha.getCellByPosition(coluna, 0)  # A linha é 0 (linha 1)
        celula.Value = i

    return
