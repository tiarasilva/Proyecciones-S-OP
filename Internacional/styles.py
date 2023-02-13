
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from constants import *

import time
start_time = time.time()

def run_styles(ws):
  thin = Side(border_style="thin", color=white)
  line_blue = Side(border_style="thin", color=blue)

  for letter in ['J', 'P', 'T']:
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
  ws.column_dimensions['G'].width = 16

  for i in range(8, ws.max_column + 1):
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
    ws[f'G{i}'].font = Font(bold=False, color=blue)
    ws[f'A{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'B{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'C{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'D{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'E{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'F{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'G{i}'].fill = PatternFill("solid", fgColor=lightlightBlue)
    ws[f'A{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'B{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'C{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'D{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'E{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'F{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'G{i}'].border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # Linea separadora azul
    ws[f'H{i}'].border = Border(left=line_blue)
    ws[f'J{i}'].border = Border(left=line_blue)
    ws[f'P{i}'].border = Border(left=line_blue)
    ws[f'T{i}'].border = Border(left=line_blue)
    ws[f'Y{i}'].border = Border(left=line_blue)

    # Bold optimista y pesimista
    ws[f'N{i}'].font = Font(bold=True)
    ws[f'O{i}'].font = Font(bold=True)
    ws[f'R{i}'].font = Font(bold=True)
    ws[f'S{i}'].font = Font(bold=True)
    ws[f'W{i}'].font = Font(bold=True)
    ws[f'X{i}'].font = Font(bold=True)

    # Merge 
    ws.merge_cells('J1:O1')
    ws.merge_cells('P1:S1')
    ws.merge_cells('T1:X1')

def run_number_format(ws):
  print("--- %s NORMAL 6.1 ---" % (time.time() - start_time))
  for i in range(2, ws.max_row + 1):
    for j in range(8, ws.max_column + 1):
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
      ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
  print("--- %s NORMAL 6.2 ---" % (time.time() - start_time))
