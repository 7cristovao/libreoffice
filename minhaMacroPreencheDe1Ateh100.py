#!/usr/bin/env python3

import uno

def main():
    preencher_numeros()
    return

def preencher_numeros():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    for i in range(1, 101):
        celula = planilha.getCellByPosition(0, i - 1)  # A coluna é 0 (A), a linha começa em 0
        celula.Value = i

    return
