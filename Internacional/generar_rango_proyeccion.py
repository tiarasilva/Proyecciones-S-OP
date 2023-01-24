from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from Stock.stock import stock
from ETA.ETA import create_ETA
from MessageBox.MessageBox import messageBox
from styles import run_styles, run_number_format
from constants import *

import time
import calendar
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import holidays
start_time = time.time()

# ----- 1. Abrimos el excel de los parametros
wb_parametros = load_workbook(filename_parametros, data_only = True, read_only = True)
ws_parametros_time = wb_parametros['Lead time']
ws_parametros_venta = wb_parametros['Venta']

# DIAS PARA VENDER
ws_parametros_dias_venta = wb_parametros['Cierre venta']
dict_cierre_venta = {}
ws_dias_max = ws_parametros_dias_venta.max_row

for row in ws_parametros_dias_venta.iter_rows(2, ws_dias_max, values_only = True):
  if row[1] is None:
    break
  oficina = row[1]
  dias_cierre = row[2]
  dict_cierre_venta[oficina.lower()] = dias_cierre

# FECHAS
selected_tipo_venta = ws_parametros_venta['B1'].value
selected_month = ws_parametros_venta['B2'].value
number_selected_month = month_number[selected_month.lower()]
today = datetime.now()
date_selected_month = date(today.year, number_selected_month, 1)
last_day_month = calendar.monthrange(today.year, today.month)[1]

# HOLIDAYS
dict_holidays = {
  'agro america': holidays.US(years=today.year),     # USA
  'agro europa': holidays.IT(years=today.year),      # ITALIA
  'agro mexico': holidays.MX(years=today.year),      # Mexico
  'agrosuper shanghai': holidays.CN(years=today.year),    # China
  'andes asia': holidays.KR(years=today.year)         # Corea del sur
}

# TIPO DE VENTA y LEAD TIME
venta_iteracion = 'directa'
escenario_iteracion = 'optimista'
dict_lead_time = {
  'optimista': {'directa': {}, 'local': {}},
  'pesimista': {'directa': {}, 'local': {}},
}

for row in ws_parametros_time.iter_rows(2, ws_parametros_time.max_row, values_only = True):
  if 'Venta Local' == row[0]:
    venta_iteracion = 'local'
  
  if 'PESIMISTA' == row[0]:
    escenario_iteracion = 'pesimista'
    venta_iteracion = 'directa'

  oficina = row[1]
  planta = row[2]
  puerto = row[3]
  agua = row[4]
  destino = row[5]
  almacen = row[6]

  if oficina is not None:
    if oficina.lower() != 'oficina':
      if 'local' in venta_iteracion.lower():
        dict_lead_time[escenario_iteracion][venta_iteracion][oficina.lower()] = { 'Planta': planta, 'Puerto': puerto, 'Agua': agua, 'Destino': destino, 'Almacen': almacen }
      else:
        dict_lead_time[escenario_iteracion][venta_iteracion][oficina.lower()] = { 'Planta': planta, 'Puerto': puerto }

wb_parametros.close()

# ----. Nombre fechas
month_1 = date_selected_month
month_2 = date_selected_month + relativedelta(months=1)
month_3 = date_selected_month + relativedelta(months=2)

name_month_1 = month_translate_EN_CL[month_1.strftime('%B').lower()]
name_month_2 = month_translate_EN_CL[month_2.strftime('%B').lower()]
name_month_3 = month_translate_EN_CL[month_3.strftime('%B').lower()]

# ----- 2. Creamos el excel de resultados
wb = Workbook()
ws = wb.create_sheet()
ws.title = 'Rango proyección'

ws.append({
  1: 'Sector', 
  2: 'Llave',
  3: 'Oficina', 
  4: 'Material', 
  5: 'Descripción',
  6: 'Venta Actual',
  7: 'Plan',
  8: f'Proyección {name_month_1}',
  14: f'Proyección {name_month_2}',
  19: f'Proyección {name_month_3}'
})

ws.append({
  8: 'ETA Pesimista',
  9: 'ETA Optimista',
  10: 'Puerto Oficina',
  11: 'Almacen oficina',
  12: 'Pesimista Proy.',
  13: 'Optimista. Proy.',
  14: 'ETA Pesimista',
  15: 'ETA Optimista',
  16: 'Pesimista Proy.',
  17: 'Optimista. Proy.',
  18: 'Asignación de venta',
  19: 'ETA Pesimista',
  20: 'ETA Optimista',
  21: 'Pesimista Proy.',
  22: 'Optimista. Proy.'
})
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
    if oficina.lower() in dict_lead_time['optimista'][selected_tipo_venta.lower()].keys():
      ws.append({ 1: sector,
                  2: f'{oficina.lower()}{material}',
                  3: oficina, 
                  4: int(material), 
                  5: descripcion, 
                  6: venta_total or 0, 
                  7: plan_total or 0,
                  8: 0,
                  10: 0,
                  11: 0,
                  14: 0,
                  18: 0,
                  19: 0
                })
wb_venta.close()

# ----- 4. Creamos la sheet Stock - Oficina
ws_stock_oficina = wb.create_sheet('Stock - Oficina')
stock(ws_stock_oficina, dict_lead_time, selected_tipo_venta, selected_month)
wb.save(filename)

# ----- 5. Creamos la sheet Stock - Oficina
print("--- %s 4.ETA inicio ---" % (time.time() - start_time))
ws_stock_ETA = wb.create_sheet('Stock - ETA')
create_ETA(ws_stock_ETA, dict_lead_time, selected_tipo_venta, date_selected_month, dict_cierre_venta)
wb.save(filename)
ws_ETA_max_row = ws_stock_ETA.max_row
print("--- %s 5. ETA final---" % (time.time() - start_time))

# ----- 6. Agregamos Stock Puerto Oficina,	Almacen oficina
dict_stock = {}
dict_stock_all = {}

for row in ws_stock_oficina.iter_rows(4, ws_stock_oficina.max_row, values_only=True):
  sector = row[0]
  oficina = row[1]
  material = row[2]
  descripcion = row[3]
  llave = f'{oficina.lower()}{material}'
  puerto_oficina = row[8] + row[12]
  almacen = row[16] + row[20]
  dict_stock[llave] = { 'Puerto oficina': puerto_oficina, 'Almacen': almacen }
  dict_stock_all[llave] = { 'sector': sector, 'oficina': oficina, 'material': material, 'descripcion': descripcion, 'Puerto oficina': puerto_oficina, 'Almacen': almacen }

print("--- %s 6. ---" % (time.time() - start_time))
dict_leftover_country = {}

for i, row in enumerate(ws.iter_rows(3, ws.max_row, values_only = True), 3):
  llave = row[1]
  oficina = row[2]
  lead_time_opt = dict_lead_time['optimista'][selected_tipo_venta.lower()][oficina.lower()]
  holidays_country = dict_holidays[oficina.lower()]

  if llave in dict_stock:
    ws[f'J{i}'].value = dict_stock[llave]['Puerto oficina'] or 0
    ws[f'K{i}'].value = dict_stock[llave]['Almacen'] or 0
    dict_stock_all.pop(llave, None)

  # -- Feriados
  leftover_days = 0
  if oficina in dict_leftover_country:
    leftover_days = dict_leftover_country[oficina.lower()]
  else:
    for day in range(today.day + 1, last_day_month + 1):
      date_day = date(today.year, number_selected_month, day)
      if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
        leftover_days += 1
    dict_leftover_country[oficina.lower()] = leftover_days
  
  if 'local' in selected_tipo_venta.lower():
    # -- MES N
    ws[f'L{i}'].value = f'=F{i} + K{i} + H{i}'                    # PESIMISTA --> Venta Actual + Almacen oficina
    ws[f'M{i}'].value = f'=F{i} + K{i} + I{i}'                    # OPTIMISTA --> Venta Actual + Almacen oficina

    if leftover_days >= round(lead_time_opt['Destino'], 2):
      ws[f'M{i}'].value = f'=F{i} + K{i} + J{i} + I{i}'           # OPTIMISTA --> + Puerto Oficina
      ws[f'M{i}'].fill = PatternFill("solid", fgColor=yellow)
    
    # -- MES N + 1
    # + no alcance a vender del MES N
    # + producción de este mes
    # + ETAS
    ws[f'P{i}'].value = f'=N{i}'
    ws[f'Q{i}'].value = f'=O{i}'

    # -- MES N + 2
    # + no alcance a vender del MES N + 1
    # + Asignación de venta
    # + ETAS
    ws[f'U{i}'].value = f'= 0.7 * R{i} + S{i}'
    ws[f'V{i}'].value = f'= 0.7 * R{i} + T{i}'

    # ----- Stock planta	Puerto Chile	Centro Agua
  ws[f'I{i}'].value = f"=SUMIFS('Stock - ETA'!$G$3:G{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$5)"
  ws[f'O{i}'].value = f"=SUMIFS('Stock - ETA'!$H$3:H{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$5) + SUMIFS('Stock - ETA'!$G$3:G{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$7)"
  ws[f'T{i}'].value = f"=SUMIFS('Stock - ETA'!$I$3:I{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$5) + SUMIFS('Stock - ETA'!$H$3:H{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$8)"

  ws[f'H{i}'].value = f"=SUMIFS('Stock - ETA'!$Q$3:Q{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$Z$3:Z{ws_ETA_max_row}, 'Rango proyección'!$X$5)"
  ws[f'N{i}'].value = f"=SUMIFS('Stock - ETA'!$R$3:R{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$Z$3:Z{ws_ETA_max_row}, 'Rango proyección'!$X$5) + SUMIFS('Stock - ETA'!$Q$3:Q{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$Z$3:Z{ws_ETA_max_row}, 'Rango proyección'!$X$7)"
  ws[f'S{i}'].value = f"=SUMIFS('Stock - ETA'!$S$3:S{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$Z$3:Z{ws_ETA_max_row}, 'Rango proyección'!$X$5) + SUMIFS('Stock - ETA'!$R$3:R{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$Z$3:Z{ws_ETA_max_row}, 'Rango proyección'!$X$8)"
wb.save(filename)

print("--- %s 7. ---" % (time.time() - start_time))
# ----- Stock sin Venta ni Plan
j = ws.max_row
for key, value in dict_stock_all.items():
  j += 1
  of = dict_stock_all[key]['oficina']
  mat = dict_stock_all[key]['material']
  ws.append({
    1: dict_stock_all[key]['sector'],
    2: f'{of}{mat}',
    3: of,
    4: mat,
    5: dict_stock_all[key]['descripcion'],
    6: 0,
    7: 0,
    8: 0,
    9: f"=SUMIFS('Stock - ETA'!$G$3:G{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$5)",
    10: dict_stock_all[key]['Puerto oficina'],
    11: dict_stock_all[key]['Almacen'],
    12: f'=F{j} + K{j} + H{i}',
    13: f'=F{j} + K{j} + I{i}',
    14: 0,
    15: f"=SUMIFS('Stock - ETA'!$H$3:H{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$5) + SUMIFS('Stock - ETA'!$G$3:G{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$7)",
    16: f'=N{j}',
    17: f'=O{j}',
    18: 0,
    19: 0,
    20: f"=SUMIFS('Stock - ETA'!$I$3:I{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$5) + SUMIFS('Stock - ETA'!$H$3:H{ws_ETA_max_row}, 'Stock - ETA'!$E$3:E{ws_ETA_max_row}, 'Rango proyección'!B{i}, 'Stock - ETA'!$P$3:P{ws_ETA_max_row}, 'Rango proyección'!$X$8)",
    21: f'=0.7 * R{j} + S{j}',
    22: f'=0.7 * R{j} + T{j}',
  })
  
print("--- %s 8. ---" % (time.time() - start_time))

# ----- Guardar la información
run_styles(ws)
run_number_format(ws)

ws['X5'].value = "SI"
ws['X6'].value = f"Mes {month_1.month}"
ws['X7'].value = f"Mes {month_2.month}"
ws['X8'].value = f"Mes {month_3.month}"

ws['X3'].fill = PatternFill("solid", fgColor=yellow)
ws['Y3'].value = 'Se suma los pedidos de puerto oficina'

ws['X4'].fill = PatternFill("solid", fgColor=lightGreen)
ws['Y4'].value = 'Se suma los pedidos que llegan este mes de agua'
print("--- %s 11. ---" % (time.time() - start_time))

wb.save(filename)
wb.close()
print("--- %s seconds ---" % (time.time() - start_time))
messageBox(dict_lead_time, selected_tipo_venta)