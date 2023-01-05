
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from constants import *

def run_styles(ws):
  for i in range(i, ws.max_column):
    ws[f'{get_column_letter(i)}1'].font = Font(bold=True, color=white)