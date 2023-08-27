import uno

def write_hello_to_cell():
    # Obter o contexto do script
    local_ctx = XSCRIPTCONTEXT.getComponentContext()
    
    # Obter o modelo do documento ativo
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    
    if doc is not None:
        # Obter a planilha ativa
        sheet = doc.CurrentController.ActiveSheet
        
        # Definir o valor "Olá Mundo" na célula A1
        cell_range = sheet.getCellRangeByName("A1")
        cell_range.setString("Olá Mundo")

# Chamar a função
write_hello_to_cell()
