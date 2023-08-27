#!/usr/bin/env python3

import uno
from com.sun.star.table import CellRangeAddress

def main():
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    
    doc = desktop.getCurrentComponent()
    if not doc:
        print("Nenhum documento aberto")
        return
    
    sheet = doc.Sheets.getByIndex(0)  # Acessa a primeira planilha
    
    # Fórmula na célula B15
    formula_b15 = "=MATCH(F15;C2:C7;0)-1"
    cell_b15 = sheet.getCellRangeByName("B15")
    cell_b15.setFormula(formula_b15)
    
    # Fórmula na célula F15
    formula_f15 = "=VLOOKUP(INDEX(A2:E7;MATCH(A15;E2:E7;0);1);A2:E7;3;0)"
    cell_f15 = sheet.getCellRangeByName("F15")
    cell_f15.setFormula(formula_f15)
    
    return

if __name__ == "__main__":
    main()

