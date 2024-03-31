import openpyxl
import re
from openpyxl.styles import PatternFill, Alignment, Font

def process_html(html, excel):
    with open(html, 'r', encoding='utf-8') as entry:
        lines = entry.readlines()
    data_diff = []
    data_initial_plus = []
    data_initial_minus = []
    initial_plus = []  
    initial_minus = []  
    for line in lines:
        if re.match(r'^diff', line):
            diff_part = re.sub(r'^diff --git a/(.*) b/(.*)', r'/\1', line.strip())
            if not re.search(r'\.(webp|jpg)$', diff_part):
                data_diff.append(diff_part)
                if initial_plus:
                    data_initial_plus.append(initial_plus)
                initial_plus = [] 
                if initial_minus:
                    data_initial_minus.append(initial_minus)
                initial_minus = []  
        elif re.match(r'^\+', line):
            initial_plus.append(line.strip()[1:]) 
        elif re.match(r'^-', line):
            initial_minus.append(line.strip()[1:]) 
    if initial_plus:
        data_initial_plus.append(initial_plus)
    if initial_minus:
        data_initial_minus.append(initial_minus)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["URL", "Additions", "Deletions"])
    font_diff = Font(color="FFFFFF", size=24)
    fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
    for idx, diff in enumerate(data_diff, start=2):  
        celda = ws.cell(row=idx, column=1, value=diff)
        celda.fill = fill
        celda.font = font_diff
    for fila_inicio, changes in enumerate(data_initial_plus, start=2): 
        text_changes = '\n'.join(changes)
        ws.cell(row=fila_inicio, column=2, value=text_changes)

    for fila_inicio, changes in enumerate(data_initial_minus, start=2): 
        text_changes = '\n'.join(changes)
        ws.cell(row=fila_inicio, column=3, value=text_changes)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for fila in range(1, len(data_diff) + 1):
        ws.cell(row=fila, column=1).font = font_diff

    for row in ws.iter_rows(min_row=1, max_row=len(data_diff) + 1, max_col=3):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.column_dimensions['A'].width = 100  
    ws.column_dimensions['B'].width = 100  
    ws.column_dimensions['C'].width = 100  
    wb.save(excel)

process_html("entry.html", "output_file.xlsx")
