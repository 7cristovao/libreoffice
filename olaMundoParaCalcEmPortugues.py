import uno

def escrever_ola_na_celula():
    # Obter o contexto do script
    contexto_local = XSCRIPTCONTEXT.getComponentContext()
    
    # Obter o modelo do documento ativo
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    
    if doc is not None:
        # Obter a planilha ativa
        planilha = doc.CurrentController.ActiveSheet
        
        # Definir o valor "Olá Mundo" na célula A1
        faixa_celula = planilha.getCellRangeByName("A1")
        faixa_celula.setString("Olá Mundo")

# Chamar a função
escrever_ola_na_celula()
