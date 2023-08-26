#!/usr/bin/env python3

import uno

def main():
    soma_valores()
    return

def soma_valores():
    doc = XSCRIPTCONTEXT.getDocument()
    planilha = doc.Sheets.getByName("Planilha1")  # Substitua "Planilha1" pelo nome da sua planilha

    celula_a1 = planilha.getCellRangeByName("A1")
    celula_b1 = planilha.getCellRangeByName("B1")
    celula_c1 = planilha.getCellRangeByName("C1")
    celula_d1 = planilha.getCellRangeByName("D1")

    celula_a1.Value = 1
    celula_b1.Value = 2
    celula_c1.Value = 3

    soma_formula = f"=A1 + B1 + C1"
    celula_d1.Formula = soma_formula

    return
