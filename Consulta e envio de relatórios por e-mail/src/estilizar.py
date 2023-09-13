##################################################################################################
########################### FUNÇÃO PARA ATRIBUIR OS ESTILOS A PLANILHA ###########################
##################################################################################################

from openpyxl.styles import Alignment, Border, Side

def aplicar_estilos(sheet, linha_init):
    #DEFINIÇÃO DE ALINHAMENTOS
    center = Alignment(horizontal='center', vertical='center')

    #ATRIBUIR BORDAS
    border = Border(
        left=Side(style='thin', color='000000'),  # Borda esquerda
        right=Side(style='thin', color='000000'),  # Borda direita
        top=Side(style='thin', color='000000'),  # Borda superior
        bottom=Side(style='thin', color='000000')  # Borda inferior
    )

    for row in sheet.iter_rows(min_row=linha_init, max_row=sheet.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = border
            cell.alignment = center