from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from styles import run_styles, run_number_format
from constants import *

import time
from datetime import datetime
start_time = time.time()

# ----- 1. Abrimos el excel de los parametros
wb_parametros = load_workbook(filename_parametros, data_only = True, read_only = True)
ws_parametros_time = wb_parametros['Lead time']
ws_parametros_venta = wb_parametros['Venta']

tipo_venta = ws_parametros_venta['B1'].value
# print(tipo_venta)

venta_iteracion = 'directa'
dict_parametros_venta = {'directa': {}, 'local': {}}
for row in ws_parametros_time.iter_rows(2, ws_parametros_time.max_row, values_only = True):
  if row[1] is None:
    break

  if 'Venta Local' == row[0]:
    venta_iteracion = 'local'
  oficina = row[1]
  tiempo_transito = row[2]
  dict_parametros_venta[venta_iteracion][oficina] = tiempo_transito
wb_parametros.close()

# ----- 2. Creamos el excel de resultados
wb = Workbook()
ws = wb.create_sheet('Rango proyección')
ws.append(['Sector', 'Oficina', 'Material', 'Descripción', 'Venta Actual', 'Plan', 'Stock planta', 'Puerto Chile', 'Centro Agua', 'Puerto Oficina', 'Almacen oficina','Pesimista Proy.', 'Optimista. Proy.'])
run_styles(ws)

del wb['Sheet']

# ----- 3. Leemos Venta actual
wb_venta = load_workbook(filename_venta, data_only=True, read_only=True)
ws_venta = wb_venta['Venta - Plan']

for row in ws_venta.iter_rows(7, ws_venta.max_row, values_only=True):
  sector = row[1]
  material = row[2]
  descripcion = row[3]
  oficina = row[5]
  plan_total = row[6]
  venta_total = row[7]
  if sector is not None and descripcion is not None:
    ws.append({ 1: sector, 
                2: oficina, 
                3: int(material), 
                4: descripcion, 
                5: venta_total or 0, 
                6: plan_total or 0})

# ----- 4. Puerto Chile, Centro Agua, Puerto Oficina, Almacén Oficina
wb_puerto = load_workbook(filename_puerto, data_only=True, read_only=True)
ws_puerto = wb_puerto.active
dict_puerto_chile = {}
dict_centro_agua = {}
dict_puerto_oficina = {}
dict_almacen = {}
dict_puerto = {}

ws_stock_puerto = wb.create_sheet('Stock - Puerto')
ws_stock_puerto.append({ 1: 'Sector', 2: 'Oficina', 3: 'Material', 4: 'Descripción', 5: 'Nivel 2', 6: 'Puerto Chile', 9: 'Fechas', 11: 'Centro Agua', 14: 'Fechas', 16: 'Puerto Oficina', 19: 'Fechas', 21: 'Almacén Oficina', 24: 'Fechas' })
ws_stock_puerto.merge_cells('F1:H1')
ws_stock_puerto.merge_cells('I1:J1')
ws_stock_puerto.merge_cells('K1:M1')
ws_stock_puerto.merge_cells('N1:O1')
ws_stock_puerto.merge_cells('P1:R1')
ws_stock_puerto.merge_cells('S1:T1')
ws_stock_puerto.merge_cells('U1:W1')
ws_stock_puerto.merge_cells('X1:Y1')
ws_stock_puerto.append({ 6: 'Stock liberado', 7: 'Stock no liberado', 8: 'Stock total', 
                          9: 'Fecha inicio', 10: 'Fecha termino', 
                          11: 'Stock liberado', 12: 'Stock no liberado', 13: 'Stock total', 
                          14: 'Fecha inicio', 15: 'Fecha termino', 
                          16: 'Stock liberado', 17: 'Stock no liberado', 18: 'Stock total',
                          19: 'Fecha inicio', 20: 'Fecha termino', 
                          21: 'Stock liberado', 22: 'Stock no liberado', 23: 'Stock total',
                          24: 'Fecha inicio', 25: 'Fecha termino' })
ws_stock_puerto.merge_cells('A2:E2')
run_styles(ws_stock_puerto)
thin = Side(border_style="thin", color=white)

for i in range(1, 26):
  ws_stock_puerto[f'{get_column_letter(i)}2'].font = Font(bold=True, color=white)
  ws_stock_puerto[f'{get_column_letter(i)}2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
  ws_stock_puerto[f'{get_column_letter(i)}2'].border = Border(top=thin, left=thin, right=thin, bottom=thin)
  ws_stock_puerto[f'{get_column_letter(i)}2'].fill = PatternFill("solid", fgColor=lightBlue)
wb.save(filename)

for i in range(6, 26):
  ws_stock_puerto.column_dimensions[f'{get_column_letter(i)}'].width = 10

for row in ws_puerto.iter_rows(8, ws_puerto.max_row, values_only=True):
  oficina = row[0]
  sector = row[1]
  material = row[2]
  descripcion = row[3]
  puerto_liberado = row[7]
  puerto_no_liberado = row[8]
  agua_liberado = row[10]
  agua_no_liberado = row[11]
  oficina_lib = row[13]
  oficina_no_lib = row[14]
  almacen_lib = row[16]
  almacen_no_lib = row[17]

  if material is not None:
    ws_stock_puerto.append({1: sector, 2: oficina, 3: int(material), 4: descripcion, 
      6: puerto_liberado or 0, 
      7: puerto_no_liberado or 0, 
      11: agua_liberado or 0, 
      12: agua_no_liberado or 0, 
      16: oficina_lib or 0, 
      17: oficina_no_lib or 0,
      21: almacen_lib or 0,
      22: almacen_no_lib or 0 })

for i in range(3, ws_stock_puerto.max_row + 1):
  ws_stock_puerto[f'H{i}'] = f'=SUM(F{i}:G{i})'
  ws_stock_puerto[f'M{i}'] = f'=SUM(K{i}:L{i})'
  ws_stock_puerto[f'R{i}'] = f'=SUM(P{i}:Q{i})'
  ws_stock_puerto[f'W{i}'] = f'=SUM(U{i}:V{i})'

  ws_stock_puerto[f'F{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'G{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'H{i}'].number_format = BUILTIN_FORMATS[3]

  ws_stock_puerto[f'I{i}'].number_format = BUILTIN_FORMATS[15]
  ws_stock_puerto[f'J{i}'].number_format = BUILTIN_FORMATS[15]
  
  ws_stock_puerto[f'K{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'L{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'M{i}'].number_format = BUILTIN_FORMATS[3]

  ws_stock_puerto[f'N{i}'].number_format = BUILTIN_FORMATS[15]
  ws_stock_puerto[f'O{i}'].number_format = BUILTIN_FORMATS[15]

  ws_stock_puerto[f'P{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'Q{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'R{i}'].number_format = BUILTIN_FORMATS[3]

  ws_stock_puerto[f'S{i}'].number_format = BUILTIN_FORMATS[15]
  ws_stock_puerto[f'T{i}'].number_format = BUILTIN_FORMATS[15]

  ws_stock_puerto[f'U{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'V{i}'].number_format = BUILTIN_FORMATS[3]
  ws_stock_puerto[f'W{i}'].number_format = BUILTIN_FORMATS[3]

  ws_stock_puerto[f'X{i}'].number_format = BUILTIN_FORMATS[15]
  ws_stock_puerto[f'Y{i}'].number_format = BUILTIN_FORMATS[15]
 
# ----- Tiempo de Stock
date = datetime.now()
print(date)

for row in ws.iter_rows(2, ws.max_row, values_only=True):
  sector = row[0]
  oficina = row[1]
  material = row[2]
  descripcion = row[3]
  puerto_liberado = row[5]
  puerto_no_liberado = row[6]
  inicio_puerto = row[8]
  termino_puerto = row[9]

  agua_liberado = row[10]
  agua_no_liberado = row[11]
  inicio_agua = row[13]
  termino_agua = row[14]

  oficina_lib = row[15]
  oficina_no_lib = row[16]
  inicio_oficina = row[18]
  termino_oficina = row[19]

  almacen_lib = row[20]
  almacen_no_lib = row[17]

  print(row)

# ----- Guardar la información
run_number_format(ws)
wb.save(filename)
wb.close()
print("--- %s seconds ---" % (time.time() - start_time))