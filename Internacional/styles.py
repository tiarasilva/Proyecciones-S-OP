
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from constants import *

def run_styles(ws):
  thin = Side(border_style="thin", color=white)
  for i in range(1, ws.max_column + 1):
    ws[f'{get_column_letter(i)}1'].font = Font(bold=True, color=white)
    ws[f'{get_column_letter(i)}1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[f'{get_column_letter(i)}1'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
    ws[f'{get_column_letter(i)}1'].fill = PatternFill("solid", fgColor=lightBlue)

  # tamaños
  ws.column_dimensions['C'].width = 25
  ws.column_dimensions['D'].width = 12