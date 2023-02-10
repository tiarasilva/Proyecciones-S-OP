
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from constants import *

import time
start_time = time.time()

def run_styles(ws):
  thin = Side(border_style="thin", color=white)
  line_blue = Side(border_style="thin", color=blue)

  for letter in ['I', 'O', 'S']:
    ws[f'{letter}1'].font = Font(bold=True, color=white)
    ws[f'{letter}1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'{letter}1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'{letter}1'].fill = PatternFill("solid", fgColor=lightBlue)

  for i in range(1, ws.max_column + 1):
    ws[f'{get_column_letter(i)}2'].font = Font(bold=True, color=white)
    ws[f'{get_column_letter(i)}2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'{get_column_letter(i)}2'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'{get_column_letter(i)}2'].fill = PatternFill("solid", fgColor=lightBlue)

  # Tamaños
  ws.column_dimensions['B'].width = 16
  ws.column_dimensions['C'].width = 23
  ws.column_dimensions['D'].width = 16
  ws.column_dimensions['F'].width = 32

  for i in range(7, ws.max_column + 1):
    ws.column_dimensions[f'{get_column_letter(i)}'].width = 11

  # Tamaño
  ws.row_dimensions[1].height =  25

  for i in range(3, ws.max_row + 1):
    ws[f'A{i}'].font = Font(bold=False, color=blue)
    ws[f'B{i}'].font = Font(bold=False, color=blue)
    ws[f'C{i}'].font = Font(bold=False, color=blue)
    ws[f'D{i}'].font = Font(bold=False, color=blue)
    ws[f'E{i}'].font = Font(bold=False, color=blue)
    ws[f'F{i}'].font = Font(bold=False, color=blue)
    ws[f'A{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'B{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'C{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'D{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'E{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'F{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'A{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'B{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'C{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'D{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'E{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'F{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # Linea separadora azul
    ws[f'G{i}'].border = Border(left=line_blue)
    ws[f'I{i}'].border = Border(left=line_blue)
    ws[f'O{i}'].border = Border(left=line_blue)
    ws[f'S{i}'].border = Border(left=line_blue)
    ws[f'X{i}'].border = Border(left=line_blue)

    # Bold optimista y pesimista
    ws[f'M{i}'].font = Font(bold=True)
    ws[f'N{i}'].font = Font(bold=True)
    ws[f'Q{i}'].font = Font(bold=True)
    ws[f'R{i}'].font = Font(bold=True)
    ws[f'V{i}'].font = Font(bold=True)
    ws[f'W{i}'].font = Font(bold=True)

    # Merge 
    ws.merge_cells('I1:N1')
    ws.merge_cells('O1:R1')
    ws.merge_cells('S1:W1')

def run_number_format(ws):
  print("--- %s NORMAL 6.1 ---" % (time.time() - start_time))
  for i in range(2, ws.max_row + 1):
    for j in range(6, ws.max_column + 1):
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
  print("--- %s NORMAL 6.2 ---" % (time.time() - start_time))
