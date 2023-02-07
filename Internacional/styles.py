
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from constants import *

import time
start_time = time.time()

def run_styles(ws, venta):
  thin = Side(border_style="thin", color=white)
  line_blue = Side(border_style="thin", color=blue)

  ws['H1'].font = Font(bold=True, color=white)
  ws['H1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
  ws['H1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
  ws['H1'].fill = PatternFill("solid", fgColor=lightBlue)

  ws['N1'].font = Font(bold=True, color=white)
  ws['N1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
  ws['N1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
  ws['N1'].fill = PatternFill("solid", fgColor=lightBlue)

  ws['R1'].font = Font(bold=True, color=white)
  ws['R1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
  ws['R1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
  ws['R1'].fill = PatternFill("solid", fgColor=lightBlue)

  for i in range(1, ws.max_column + 1):
    ws[f'{get_column_letter(i)}2'].font = Font(bold=True, color=white)
    ws[f'{get_column_letter(i)}2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'{get_column_letter(i)}2'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'{get_column_letter(i)}2'].fill = PatternFill("solid", fgColor=lightBlue)

  # Tamaños
  ws.column_dimensions['B'].width = 23
  ws.column_dimensions['C'].width = 16
  ws.column_dimensions['E'].width = 32

  for i in range(6, ws.max_column + 1):
    ws.column_dimensions[f'{get_column_letter(i)}'].width = 11

  # Merge 
  ws.merge_cells('H1:M1')
  ws.merge_cells('N1:Q1')
  ws.merge_cells('R1:V1')

  # Tamaño
  ws.row_dimensions[1].height =  25

  for i in range(3, ws.max_row + 1):
    ws[f'A{i}'].font = Font(bold=False, color=blue)
    ws[f'B{i}'].font = Font(bold=False, color=blue)
    ws[f'C{i}'].font = Font(bold=False, color=blue)
    ws[f'D{i}'].font = Font(bold=False, color=blue)
    ws[f'E{i}'].font = Font(bold=False, color=blue)
    ws[f'A{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'B{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'C{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'D{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'E{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'A{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'B{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'C{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'D{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'E{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # Linea separadora azul
    if venta == 'local':
      ws[f'F{i}'].border = Border(left=line_blue)
      ws[f'H{i}'].border = Border(left=line_blue)
      ws[f'N{i}'].border = Border(left=line_blue)
      ws[f'R{i}'].border = Border(left=line_blue)
      ws[f'W{i}'].border = Border(left=line_blue)

      # Bold optimista y pesimista
      ws[f'L{i}'].font = Font(bold=True)
      ws[f'M{i}'].font = Font(bold=True)
      ws[f'P{i}'].font = Font(bold=True)
      ws[f'Q{i}'].font = Font(bold=True)
      ws[f'U{i}'].font = Font(bold=True)
      ws[f'V{i}'].font = Font(bold=True)
    
    else:
      ws[f'F{i}'].border = Border(left=line_blue)
      ws[f'H{i}'].border = Border(left=line_blue)
      ws[f'L{i}'].border = Border(left=line_blue)
      ws[f'P{i}'].border = Border(left=line_blue)
      ws[f'U{i}'].border = Border(left=line_blue)

      ws[f'J{i}'].font = Font(bold=True)
      ws[f'K{i}'].font = Font(bold=True)
      ws[f'N{i}'].font = Font(bold=True)
      ws[f'O{i}'].font = Font(bold=True)
      ws[f'S{i}'].font = Font(bold=True)
      ws[f'T{i}'].font = Font(bold=True)


def run_number_format(ws):
  print("--- %s NORMAL 6.1 ---" % (time.time() - start_time))
  for i in range(2, ws.max_row + 1):
    for j in range(6, ws.max_column + 1):
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
  print("--- %s NORMAL 6.2 ---" % (time.time() - start_time))