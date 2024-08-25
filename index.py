from openpyxl import Workbook
from openpyxl.styles import Alignment, Border,Font,Side

def criar_plan():
    
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'Semestres'
    
    data = [
        ['Disciplina', 'Ap1','Ap2','Ext','Média Final'],
        ['Física', '-','-','-','-',],
        ['Poo', '-','-','-','-',],
        ['Inglês', '-','-','-','-',],
        ['Extensão', '-','-','-','-',],
        ['Algoritimos e estrutura de dados', '-','-','-','-',],
        ['Fundamentos de Fabricação', '-','-','-','-',],
        ['Sistemas Hidráulicos e pneumáticos', '-','-','-','-',]
    ]


    for row in data:
        sheet.append(row)
        
    wb.save('Notas.xlsx')

criar_plan()