from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from constants import *

import time
start_time = time.time()

def run_styles(ws):
  thin = Side(border_style="thin", color=white)

  for i in range(1, ws.max_column + 1):
    ws[f'{get_column_letter(i)}2'].font = Font(bold=True, color=white)
    ws[f'{get_column_letter(i)}2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'{get_column_letter(i)}2'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'{get_column_letter(i)}2'].fill = PatternFill("solid", fgColor=lightBlue)

  # Tamaños
  ws.column_dimensions['A'].width = 10
  ws.column_dimensions['B'].width = 10
  ws.column_dimensions['C'].width = 19
  ws.column_dimensions['D'].width = 10
  ws.column_dimensions['E'].width = 22
  ws.column_dimensions['F'].width = 12

  ws.column_dimensions['L'].width = 12
  ws.column_dimensions['N'].width = 12
  ws.column_dimensions['P'].width = 12

  ws.column_dimensions['V'].width = 12
  ws.column_dimensions['X'].width = 12
  ws.column_dimensions['Z'].width = 12

  ws.row_dimensions[1].height = 25

  # Merge
  ws.merge_cells('G1:P1')
  ws.merge_cells('Q1:Z1')
  ws['G1'].font = Font(bold=True, color=white)
  ws['G1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
  ws['G1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
  ws['G1'].fill = PatternFill("solid", fgColor=lightBlue)
  
  ws['Q1'].font = Font(bold=True, color=white)
  ws['Q1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
  ws['Q1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
  ws['Q1'].fill = PatternFill("solid", fgColor=lightBlue)

def run_number_format(ws):
  print("--- %s ETA 6.2 ---" % (time.time() - start_time))
  thin = Side(border_style="thin", color=white)
  border_white = Border(top=thin, left=thin, right=thin, bottom=thin)

  line_blue = Side(border_style="thin", color=blue)
  font_blue = Font(bold=False, color=blue)
  border_light_blue = Border(left=line_blue)

  fill_light_blue = PatternFill("solid", fgColor=lightlightBlue)

  print("--- %s ETA 6.3 ---" % (time.time() - start_time))
  for i in range(3, ws.max_row + 1):
    ws[f'G{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'H{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'I{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'J{i}'].number_format = BUILTIN_FORMATS[3]

    ws[f'K{i}'].number_format = BUILTIN_FORMATS[4]
    ws[f'M{i}'].number_format = BUILTIN_FORMATS[4]

    ws[f'Q{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'R{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'S{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'T{i}'].number_format = BUILTIN_FORMATS[3]

    ws[f'U{i}'].number_format = BUILTIN_FORMATS[4]
    ws[f'W{i}'].number_format = BUILTIN_FORMATS[4]

    # for j in range(1, 7):
    ws[f'A{i}'].font = font_blue
    ws[f'A{i}'].fill = fill_light_blue
    ws[f'A{i}'].border = border_white

    ws[f'B{i}'].font = font_blue
    ws[f'B{i}'].fill = fill_light_blue
    ws[f'B{i}'].border = border_white

    ws[f'C{i}'].font = font_blue
    ws[f'C{i}'].fill = fill_light_blue
    ws[f'C{i}'].border = border_white

    ws[f'D{i}'].font = font_blue
    ws[f'D{i}'].fill = fill_light_blue
    ws[f'D{i}'].border = border_white

    ws[f'E{i}'].font = font_blue
    ws[f'E{i}'].fill = fill_light_blue
    ws[f'E{i}'].border = border_white

    ws[f'F{i}'].font = font_blue
    ws[f'F{i}'].fill = fill_light_blue
    ws[f'F{i}'].border = border_white

    ws[f'G{i}'].border = border_light_blue
    ws[f'K{i}'].border = border_light_blue
    ws[f'Q{i}'].border = border_light_blue
    ws[f'U{i}'].border = border_light_blue
    ws[f'AA{i}'].border = border_light_blue

    ws[f'O{i}'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'P{i}'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'Y{i}'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'Z{i}'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
  print("--- %s ETA 6.4 ---" % (time.time() - start_time))