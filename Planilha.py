from openpyxl import Workbook
from openpyxl.styles import Alignment, Border,Font,Side
from openpyxl.formatting.rule import CellIsRule
import os

def criar_plan():
    
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'Semestres'
    
    data = [
        ['Disciplina', 'Ap1','Ap2','Ext','Média Final'],
        ['Física',0,0,0,0],
        ['Poo', 0,0,0,0],
        ['Inglês',0,0,0,0],
        ['Extensão',0,0,0,0],
        ['Algoritimos e estrutura de dados',0,0,0,0],
        ['Fundamentos de Fabricação', 0,0,0,0],
        ['Sistemas Hidráulicos e pneumáticos',0,0,0,0]
    ]

    
    for row in data:
        sheet.append(row)
    
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value))> max_length:
                    max_length = len(cell.value)
            except:
                pass
            adjusted_width = max_length + 2 
            sheet.column_dimensions[column].width = adjusted_width
        
    number_format = '0.0'
    
    for row in sheet.iter_rows(min_row=2, max_row=8,min_col=2, max_col=5):
        for cell in row:
            cell.number_format = number_format
            
    wb.save('Notas.xlsx')

criar_plan()

os.startfile('Notas.xlsx')