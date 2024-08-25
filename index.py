from openpyxl import workbook, load_workbook

def criar_plan(Notas):
    wb = workbook
    ws = wb.active
    ws.title = 'Semestre'
    
