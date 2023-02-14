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
# from TD.pivotTable import pivot_table

import sys
import os

start_time = time.time()

# ----- 0. PATH
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
  print('running in a PyInstaller bundle')
  os.chdir(sys._MEIPASS)
else:
  print('running in a normal Python process')

# ----- 1. Abrimos el excel de los parametros
wb_parametros = load_workbook(filename_parametros, data_only = True, read_only = True)
ws_parametros_time = wb_parametros['Lead time']
ws_parametros_venta = wb_parametros['Venta']
ws_porcentaje = wb_parametros['Asignación productivo']

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
selected_year = ws_parametros_venta['B1'].value
selected_month = ws_parametros_venta['B2'].value
selected_week = ws_parametros_venta['B3'].value
number_selected_month = month_number[selected_month.lower()]
today = datetime.now()
date_selected_month = date(int(selected_year), number_selected_month, 1)
last_day_month = calendar.monthrange(today.year, today.month)[1]

# Nombre fechas
month_1 = date_selected_month
month_2 = date_selected_month + relativedelta(months=1)
month_3 = date_selected_month + relativedelta(months=2)

name_month_1 = month_translate_EN_CL[month_1.strftime('%B').lower()]
name_month_2 = month_translate_EN_CL[month_2.strftime('%B').lower()]
name_month_3 = month_translate_EN_CL[month_3.strftime('%B').lower()]

selected_month_year = f'{month_1.strftime("%m")}.{month_1.year}'

# HOLIDAYS
last_year = month_1.year - 1
this_year = month_1.year
next_year = month_1.year + 1

dict_holidays = {
  last_year: {
    'agro america': holidays.US(years=last_year),         # USA
    'agro europa': holidays.IT(years=last_year),          # ITALIA
    'agro mexico': holidays.MX(years=last_year),          # Mexico
    'agrosuper shanghai': holidays.CN(years=last_year),   # China
    'andes asia': holidays.KR(years=last_year),           # Corea del sur
    'chile': holidays.CL(years=last_year)
  },
  today.year: {
    'agro america': holidays.US(years=this_year),         # USA
    'agro europa': holidays.IT(years=this_year),          # ITALIA
    'agro mexico': holidays.MX(years=this_year),          # Mexico
    'agrosuper shanghai': holidays.CN(years=this_year),   # China
    'andes asia': holidays.KR(years=this_year),           # Corea del sur
    'chile': holidays.CL(years=this_year)
  },
  next_year: {
    'agro america': holidays.US(years=next_year),         # USA
    'agro europa': holidays.IT(years=next_year),          # ITALIA
    'agro mexico': holidays.MX(years=next_year),          # Mexico
    'agrosuper shanghai': holidays.CN(years=next_year),   # China
    'andes asia': holidays.KR(years=next_year),           # Corea del sur
    'chile': holidays.CL(years=next_year)
  }
}

# TIPO DE VENTA y LEAD TIME
venta_iteracion = 'Venta Directa'
escenario_iteracion = 'optimista'
dict_lead_time = {
  'optimista': {'Venta Directa': {}, 'Venta Local': {}},
  'pesimista': {'Venta Directa': {}, 'Venta Local': {}},
}

for row in ws_parametros_time.iter_rows(2, ws_parametros_time.max_row, values_only = True):
  if 'Venta Local' == row[0]:
    venta_iteracion = 'Venta Local'
  
  if 'Venta Directa' == row[0]:
    venta_iteracion = 'Venta Directa'
  
  if 'PESIMISTA' == row[0]:
    escenario_iteracion = 'pesimista'

  oficina = row[1]
  planta = row[2]
  puerto = row[3]
  agua = row[4]
  destino = row[5]
  almacen = row[6]

  if oficina is not None:
    if oficina.lower() != 'oficina':
      if 'Venta Local' == venta_iteracion:
        dict_lead_time[escenario_iteracion][venta_iteracion][oficina.lower()] = { 'Planta': planta, 'Puerto': puerto, 'Agua': agua, 'Destino': destino, 'Almacen': almacen }
      else:
        dict_lead_time[escenario_iteracion][venta_iteracion][oficina.lower()] = { 'Planta': planta, 'Puerto': puerto }

# PORCENTAJE DE PRODUCCIÓN
dict_porcentaje_produccion = {}

for row in ws_porcentaje.iter_rows(2, ws_porcentaje.max_row, values_only=True):
  oficina = row[1].lower()
  produccion = row[2]
  dict_porcentaje_produccion[oficina] = produccion

wb_parametros.close()

# ----- 2. Creamos el excel de resultados VENTA LOCAL y VENTA DIRECTA
wb = Workbook()
ws = wb.active
ws.title = sheet_name

ws.append({
  10: f'Proyección {name_month_1} {month_1.year}',
  18: f'Proyección {name_month_2} {month_2.year}',
  22: f'Proyección {name_month_3} {month_3.year}'
})

ws.append({
  1: 'Sector',                # A
  2: 'Canal de Distribución', # B
  3: 'Llave',                 # C
  4: 'Oficina',               # D
  5: 'Material',              # E
  6: 'Descripción',           # F
  7: 'Nivel 2',               # G
  8: 'Venta Actual',          # H
  9: 'Plan',                  # I
  10: 'ETA Pesimista',        # J
  11: 'ETA Optimista',        # K
  12: 'Puerto Chile Pes.',    # L
  13: 'Puerto Chile Opt.',    # M
  14: 'Puerto Oficina',       # N
  15: 'Almacen oficina',      # O
  16: 'Pesimista Proy.',      # P
  17: 'Optimista Proy.',      # Q
  18: 'ETA Pesimista',        # R
  19: 'ETA Optimista',        # S
  20: 'Pesimista Proy.2',     # T
  21: 'Optimista Proy.2',     # U
  22: 'Asignación de venta',  # V
  23: 'ETA Pesimista',        # W
  24: 'ETA Optimista',        # X
  25: 'Pesimista Proy.3',     # Y
  26: 'Optimista Proy.3'      # Z
})

# ----- 3. Leemos Venta actual
wb_venta = load_workbook(filename_venta, data_only=True, read_only=True)
ws_venta = wb_venta['Venta - Plan']

for row in ws_venta.iter_rows(7, ws_venta.max_row, values_only=True):
  month_year = row[0]
  sector = row[1]
  oficina = row[3]
  material = row[4]
  descripcion = row[5]
  nivel_2 = row[6]
  venta_total = row[7]
  plan_total = row[8]

  canal_distribucion = 'Venta Directa'
  
  if oficina is not None:
    if oficina.lower() in dict_lead_time['optimista']['Venta Local'].keys():
      canal_distribucion = 'Venta Local'

    if month_year == selected_month_year:
      ws.append({ 1: sector,
                  2: canal_distribucion,
                  3: f'{oficina.lower()}{material}',
                  4: oficina, 
                  5: int(material),
                  6: descripcion,
                  7: nivel_2, 
                  8: venta_total or 0,                # G
                  9: plan_total or 0,                 # H
                  20: 0
                })

wb_venta.close()

# ----- 4. Creamos la sheet Stock - Oficina
ws_stock_oficina = wb.create_sheet('Stock - Oficina')
stock(ws_stock_oficina, dict_lead_time)
wb.save(filename)

# ----- 5. Creamos la sheet Stock - Oficina
print("--- %s 4.ETA inicio ---" % (time.time() - start_time))
ws_stock_ETA = wb.create_sheet(sheet_stock)
create_ETA(ws_stock_ETA, dict_lead_time, date_selected_month, dict_cierre_venta, dict_holidays)
wb.save(filename)
ETA_maxRow = ws_stock_ETA.max_row

print("--- %s 5. ETA final---" % (time.time() - start_time))

# ----- 6. Agregamos Stock Puerto Oficina,	Almacen oficina
dict_stock = {}

for row in ws_stock_oficina.iter_rows(4, ws_stock_oficina.max_row, values_only=True):
  sector = row[0]
  oficina = row[1]
  material = row[2]
  descripcion = row[3]
  llave = f'{oficina.lower()}{material}'
  puerto_oficina = row[8] + row[12]
  almacen = row[16] + row[20]
  dict_stock[llave] = { 'sector': sector, 'oficina': oficina, 'material': material, 'descripcion': descripcion, 'Puerto oficina': puerto_oficina, 'Almacen': almacen }

print("--- %s 6. ---" % (time.time() - start_time))

# ----- 7. Asignaciones de venta MES N+2
wb_asignaciones = load_workbook(filename_asignaciones, read_only=True, data_only=True)
ws_asignaciones = wb_asignaciones['Asignaciones de venta']
ws_asignaciones_max_row = ws_asignaciones.max_row
dict_asignaciones = {}
month_year = ''
sector = ''
office = ''

for row in ws_asignaciones.iter_rows(3, ws_asignaciones_max_row - 1, values_only=True):
  if row[0] is not None:
    month_year = str(row[0])
  
  if row[1] is not None:
    sector = row[1]

  if row[2] is not None:
    office = row[2]

  material = row[3]
  description = row[4]
  RV_final = row[6]

  key = f'{office.lower()}{material}'
  if month_year in dict_asignaciones:
    if key not in dict_asignaciones[month_year]:
      dict_asignaciones[month_year][key] = { 'sector': sector, 'oficina': oficina, 'material': material,'descripcion': description, 'RV final': RV_final }
  else:
    dict_asignaciones[month_year] = { key: { 'sector': sector, 'oficina': oficina, 'material': material,'descripcion': description, 'RV final': RV_final }}
wb_asignaciones.close()

# ----- 8. Stock Centro Agua
wb_agua = load_workbook(filename_chile, read_only=True, data_only=True)
ws_agua = wb_agua['Stock']
agua_max = ws_agua.max_row
dict_agua = {}

for row in ws_agua.iter_rows(4, agua_max, values_only=True):
  month_year = row[0]
  sector = row[1]
  oficina = row[2]
  material = row[3]
  descripcion = row[4]
  nivel_2 = row[5]
  puerto_chile = row[8]
  key = f'{oficina.lower()}{material}'

  if month_year in dict_agua:
    if key not in dict_agua[month_year]:
      dict_agua[month_year][key] = {
        'fecha': month_year,
        'sector': sector,
        'oficina': oficina,
        'material': material,
        'descripcion': descripcion,
        'nivel 2': nivel_2,
        'puerto chile': puerto_chile
      }
  else:
    dict_agua[month_year] = {
      key: {
        'fecha': month_year,
        'sector': sector,
        'oficina': oficina,
        'material': material,
        'descripcion': descripcion,
        'nivel 2': nivel_2,
        'puerto chile': puerto_chile
      }
    }
wb_agua.close()

# ----- 9.
dict_leftover_country = {}
key_month1_year = f'{month_1.strftime("%m")}.{month_1.year}'
key_month3_year = f'{month_3.strftime("%m")}.{month_3.year}'
max_row = ws.max_row

for i, row in enumerate(ws.iter_rows(3, max_row, values_only = True), 3):
  canal_distribucion = row[1]
  llave = row[2]
  oficina = row[3]
  nivel_2 = row[6]
  lead_time_opt = dict_lead_time['optimista'][canal_distribucion][oficina.lower()]
  lead_time_pes = dict_lead_time['pesimista'][canal_distribucion][oficina.lower()]
  last_stop_LT = 0

  if canal_distribucion == 'Venta Directa':
    last_stop_LT = round(lead_time_opt['Puerto'], 2)
    holidays_country = dict_holidays[month_1.year]['chile']
  else:
    last_stop_LT = round(lead_time_opt['Destino'], 2)
    holidays_country = dict_holidays[month_1.year][oficina.lower()]
  
   # -- 9.1. Feriados
  leftover_days = 0
  if oficina in dict_leftover_country:
    leftover_days = dict_leftover_country[oficina.lower()]
  else:
    for day in range(today.day + 1, last_day_month + 1):
      date_day = date(today.year, number_selected_month, day)
      if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
        leftover_days += 1
    dict_leftover_country[oficina.lower()] = leftover_days
    
  # -- 9.2. Centro agua mes 1
  if key_month1_year in dict_agua:
    ws[f'L{i}'].value = 0
    ws[f'M{i}'].value = 0
    if llave in dict_agua[key_month1_year]:
      stock_agua = dict_agua[key_month1_year][llave]['puerto chile'] or 0
      dict_agua[key_month1_year].pop(llave)
      LT_pes = lead_time_pes['Puerto']
      LT_opt = lead_time_opt['Puerto']

      pct_prod_pes = (leftover_days - LT_pes) / leftover_days
      pct_prod_opt = (leftover_days - LT_opt) / leftover_days

      ws[f'L{i}'].value = stock_agua * pct_prod_pes or 0
      ws[f'M{i}'].value = stock_agua * pct_prod_opt or 0
  
  # -- 9.3. Asignaciones mes 3
  if key_month3_year in dict_asignaciones:
    if llave in dict_asignaciones[key_month3_year]:
      ws[f'V{i}'].value = dict_asignaciones[key_month3_year][llave]['RV final']
      dict_asignaciones[key_month3_year].pop(llave)
  
  # -- 9.4. Stock Puerto Oficina y Almacen
  if canal_distribucion == 'Venta Local':
    if llave in dict_stock:
      ws[f'N{i}'].value = dict_stock[llave]['Puerto oficina'] or 0
      ws[f'O{i}'].value = dict_stock[llave]['Almacen'] or 0
      dict_stock.pop(llave, None)
  
  # -- 9.5. MES N
  ws[f'P{i}'].value = f'=H{i} + O{i} + J{i}'                    # PESIMISTA --> Venta Actual + Almacen oficina + ETA Pesimista n
  ws[f'Q{i}'].value = f'=H{i} + O{i} + K{i}'                    # OPTIMISTA --> Venta Actual + Almacen oficina + ETA Optimista n
  ws[f'Q{i}'].fill = PatternFill("solid", fgColor=yellow)

  if leftover_days >= last_stop_LT:
    ws[f'Q{i}'].value = f'=H{i} + O{i} + K{i} + N{i}'           # OPTIMISTA --> + Puerto Oficina
    ws[f'Q{i}'].fill = PatternFill("solid", fgColor=lightGreen)
  
  # -- 9.6. MES N + 1
  ws[f'T{i}'].value = f'=R{i}'                                  # PESIMISTA --> + ETA Pesimista n+1
  ws[f'U{i}'].value = f'=S{i}'                                  # OPTIMISTA --> + ETA Optimista n+1

  # -- 9.7. MES N + 2
  porcentaje = dict_porcentaje_produccion[oficina.lower()]
  ws[f'Y{i}'].value = f'= {porcentaje} * V{i} + W{i}'           # PESIMISTA --> Asignación de venta + ETA Pesimista n+2
  ws[f'Z{i}'].value = f'= {porcentaje} * V{i} + X{i}'           # OPTIMISTA --> Asignación de venta + ETA Optimista n+2
  
  # ----- Stock planta	Puerto Chile	Centro Agua
  ws[f'J{i}'].value = f"=SUMIFS('{sheet_stock}'!$R$3:R{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$5)"
  ws[f'K{i}'].value = f"=SUMIFS('{sheet_stock}'!$H$3:H{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$5)"

  ws[f'R{i}'].value = f"=SUMIFS('{sheet_stock}'!$S$3:S{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$R$3:R{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$7)"
  ws[f'S{i}'].value = f"=SUMIFS('{sheet_stock}'!$I$3:I{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$H$3:H{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$7)"
  
  ws[f'W{i}'].value = f"=SUMIFS('{sheet_stock}'!$T$3:T{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$S$3:S{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$8)"
  ws[f'X{i}'].value = f"=SUMIFS('{sheet_stock}'!$J$3:J{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$I$3:I{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$8)"

print("--- %s 10. ---" % (time.time() - start_time))
# ----- 10. Stock sin Venta ni Plan
i = ws.max_row

for key, value in dict_stock.items():
  of = value['oficina']
  mat = value['material']
  key = f'{of}{mat}'
  porcentaje = dict_porcentaje_produccion[of.lower()]

  canal_distribucion = 'Venta Directa'
  
  if of.lower() in dict_lead_time['optimista']['Venta Local'].keys():
    canal_distribucion = 'Venta Local'

  ws.append({
    1: value['sector'],
    2: canal_distribucion,
    3: key,
    4: of,
    5: mat,
    6: value['descripcion'],
    7: '',
    8: 0,
    9: 0,
    13: value['Puerto oficina'],     # 
    14: value['Almacen'],            # 
    21: 0,
  })
  
  if of.lower() in dict_lead_time['optimista']['Venta Local']:
    i += 1
    ws[f'J{i}'].value = f"=SUMIFS('{sheet_stock}'!$R$3:R{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$5)"
    ws[f'K{i}'].value = f"=SUMIFS('{sheet_stock}'!$H$3:H{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$5)"

    ws[f'R{i}'].value = f"=SUMIFS('{sheet_stock}'!$S$3:S{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$R$3:R{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$7)"
    ws[f'S{i}'].value = f"=SUMIFS('{sheet_stock}'!$I$3:I{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$H$3:H{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$7)"
    
    ws[f'W{i}'].value = f"=SUMIFS('{sheet_stock}'!$T$3:T{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$S$3:S{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AB$8)"
    ws[f'X{i}'].value = f"=SUMIFS('{sheet_stock}'!$J$3:J{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$5) + SUMIFS('{sheet_stock}'!$I$3:I{ETA_maxRow},'{sheet_stock}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_stock}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AB$8)"

    ws[f'P{i}'].value = f'=H{i} + O{i} + J{i}'
    ws[f'Q{i}'].value = f'=H{i} + O{i} + K{i}'

    ws[f'T{i}'].value = f'=R{i}'
    ws[f'U{i}'].value = f'=S{i}'

    ws[f'Y{i}'].value = f'= {porcentaje} * V{i} + W{i}'
    ws[f'Z{i}'].value = f'= {porcentaje} * V{i} + X{i}'

  if month_year in dict_asignaciones:
    if key in dict_asignaciones:
      ws[f'V{i}'].value = dict_asignaciones[month_year][key]['RV final'] 
    
  
print("--- %s 11. ---" % (time.time() - start_time))

# ----- 11. Guardar la información
run_styles(ws)
run_number_format(ws)

ws['AB5'].value = "SI"
ws['AB6'].value = f"Mes {month_1.month}"
ws['AB7'].value = f"Mes {month_2.month}"
ws['AB8'].value = f"Mes {month_3.month}"

ws['AB3'].fill = PatternFill("solid", fgColor=yellow)
ws['AC3'].value = 'Se suma los pedidos de puerto oficina'

ws['AB4'].fill = PatternFill("solid", fgColor=lightGreen)
ws['AC4'].value = 'Se suma los pedidos que llegan este mes de agua'
print("--- %s 11. ---" % (time.time() - start_time))

wb.save(filename)
print(wb.sheetnames)
print(wb['Rango proyecciones'].max_row)
wb.close()

# ----- Sheet TD
# main()
# pivot_table()

# user_response = input('Desea chequear la proyección?: (Si/No)')

# if user_response in ['Si', 'si', 'SI']:
#   wb_proy = load_workbook()

print("--- %s seconds ---" % (time.time() - start_time))
messageBox(dict_lead_time, 'Venta Local')