
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from constants import *

def run_styles(ws):
  thin = Side(border_style="thin", color=white)
  for i in range(1, ws.max_column + 1):
    ws[f'{get_column_letter(i)}1'].font = Font(bold=True, color=white)
    ws[f'{get_column_letter(i)}1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'{get_column_letter(i)}1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'{get_column_letter(i)}1'].fill = PatternFill("solid", fgColor=lightBlue)

  # Tamaños
  ws.column_dimensions['B'].width = 17
  ws.column_dimensions['D'].width = 32
  ws.column_dimensions['E'].width = 10
  ws.column_dimensions['F'].width = 10

def run_number_format(ws):
  for i in range(2, ws.max_row + 1):
    ws[f'E{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'F{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'G{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'H{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'I{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'J{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'K{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'L{i}'].number_format = BUILTIN_FORMATS[3]
    ws[f'M{i}'].number_format = BUILTIN_FORMATS[3]