#!/usr/bin/env python3

import uno

def main():
    hola_writer()
    return

def hola_writer():
    doc = XSCRIPTCONTEXT.getDocument()
    texto = doc.Text
    texto.String = 'Hola Mundo'
    return
