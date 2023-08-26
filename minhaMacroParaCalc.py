#!/usr/bin/env python3

import uno

def main():
    hola_calc()
    return

def hola_calc():
    doc = XSCRIPTCONTEXT.getDocument()
    celda = doc.CurrentController.Selection
    celda.String = 'Hola Mundo'
    return
