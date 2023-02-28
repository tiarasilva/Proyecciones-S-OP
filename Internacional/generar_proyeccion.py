from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_PERCENTAGE, BUILTIN_FORMATS
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from constants import *
from ETA.ETA import create_ETA
from MessageBox.MessageBox import messageBox
from PuertoChile.PuertoChile import create_puerto_chile
# from ProdVenta.MissingProduction import create_missing_production
from Stock.stock import stock
from styles import run_styles

import time
import calendar
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import holidays

import sys, os
from os import path

start_time = time.time()

# ----- 0. PATH
if getattr(sys, 'frozen', False):
  print('\nRunning in a PyInstaller bundle\n')
  bundle_dir = sys._MEIPASS
  filename_parametros = path.abspath(path.join(path.dirname(__file__), filename_parametros))
  filename_venta = path.abspath(path.join(path.dirname(__file__), filename_venta))
  filename_asignaciones = path.abspath(path.join(path.dirname(__file__), filename_asignaciones))
  filename_chile = path.abspath(path.join(path.dirname(__file__), filename_chile))
  filename_dias = path.abspath(path.join(path.dirname(__file__), filename_dias))
  filename_logistica = path.abspath(path.join(path.dirname(__file__), filename_logistica))
  filename_pedidos_confirmados = path.abspath(path.join(path.dirname(__file__), filename_pedidos_confirmados))
  filename = path.abspath(path.join(path.dirname(sys.executable), filename))
  path_img = path.abspath(path.join(path.dirname(__file__), path_img))

else:
  print('\nRunning in a Python environment\n')
  bundle_dir = os.path.dirname(os.path.abspath(__file__))

# # ----- 1. Abrimos el excel de los parametros
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

# LEFTOVER POR OFICINAS
dict_leftover_country = {
  'chile': 0,
  'agro america' : 0,
  'agro europa' : 0,
  'agro mexico' : 0,
  'agrosuper shanghai' : 0,
  'andes asia' : 0,
}

for oficina in dict_leftover_country.keys():
  leftover_days = 0
  holidays_country = dict_holidays[month_1.year][oficina.lower()]

  if oficina == 'chile':
    holidays_country = dict_holidays[month_1.year]['chile']

  for day in range(today.day + 1, last_day_month + 1):
    date_day = date(today.year, number_selected_month, day)
    if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
      leftover_days += 1
  dict_leftover_country[oficina.lower()] = leftover_days

# PRODUCTIVE DAYS IN CHILE
productive_days = 0
holidays_country = dict_holidays[month_1.year]['chile']

for day in range(1, last_day_month + 1):
  date_day = date(today.year, number_selected_month, day)
  if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
    productive_days +=1


# ----- 2. Creamos el excel de resultados VENTA LOCAL y VENTA DIRECTA
wb = Workbook()
ws = wb.active
ws.title = sheet_name

ws.append({
  10: f'Proyección {name_month_1} {month_1.year}',
  22: f'Proyección {name_month_2} {month_2.year}',
  28: f'Proyección {name_month_3} {month_3.year}'
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
  11: 'Plan - Prod. actual',  # K
  12: 'Puerto Chile Pes.',    # L
  13: 'Puerto Oficina',       # M
  14: 'Almacen oficina',      # N
  15: 'Proy. Pesimista',      # O

  16: 'ETA Optimista',        # P
  17: 'Plan - Prod. actual',  # Q
  18: 'Puerto Chile Opt.',    # R
  19: 'Puerto Oficina',       # S
  20: 'Almacen oficina',      # T
  21: 'Proy. Optimista',      # U

  22: 'ETA Pesimista N+1',    # V
  23: 'Inventario mes N',     # W
  24: f'Proy. Pesimista {name_month_2}',  # X

  25: 'ETA Optimista N+1',    # Y
  26: 'Inventario mes N',     # Z
  27: f'Proy. Optimista {name_month_2}',  # AA

  28: 'Asignación de venta',  # AB
  29: 'ETA Pesimista N+2',    # AC
  30: f'Proy. Pesimista {name_month_3}',  # AD
  
  31: 'ETA Optimista N+2',    # AE
  32: f'Proy. Optimista {name_month_3}'   # AF
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
stock(ws_stock_oficina, dict_lead_time, filename_dias)
wb.save(filename)

# ----- 5. Creamos la sheet Stock - ETA
print("--- %s 4.ETA inicio ---" % (time.time() - start_time))
ws_stock_ETA = wb.create_sheet(sheet_name_ETA)
create_ETA(ws_stock_ETA, dict_lead_time, date_selected_month, dict_cierre_venta, dict_holidays, filename_logistica, filename_pedidos_confirmados)
wb.save(filename)
ETA_maxRow = ws_stock_ETA.max_row

print("--- %s 5. ---" % (time.time() - start_time))

# ----- 6. Creamos la sheet Producción Faltante
# print("--- %s 6.ETA inicio ---" % (time.time() - start_time))
# ws_missing_prod = wb.create_sheet(sheet_name_MP)
# create_missing_production(ws_missing_prod, dict_leftover_country)
# wb.save(filename)

# print("--- %s 7. ---" % (time.time() - start_time))

# ----- 7. Agregamos Stock Puerto Oficina,	Almacen oficina
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

# ----- 8. Asignaciones de venta MES N+2
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

# ----- 9. Creamos la sheet Stock - Puerto Chile
ws_puerto_chile = wb.create_sheet(sheet_name_PC)
create_puerto_chile(ws_puerto_chile, filename_chile, dict_lead_time, dict_holidays, month_1, dict_leftover_country)
wb.save(filename)
PC_max_row = ws_puerto_chile.max_row

wb_agua = load_workbook(filename_chile, read_only=True, data_only=True)
ws_agua = wb_agua['Stock']
agua_max = ws_agua.max_row
dict_agua = {}

i = 4
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
  i += 1
wb_agua.close()

# ----- 9.
key_month1_year = f'{month_1.strftime("%m")}.{month_1.year}'
key_month3_year = f'{month_3.strftime("%m")}.{month_3.year}'
max_row = ws.max_row

print("--- %s 9.1 ---" % (time.time() - start_time))
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
    leftover_days = dict_leftover_country['chile']
  else:
    last_stop_LT = round(lead_time_opt['Destino'], 2)
    leftover_days = dict_leftover_country[oficina.lower()]

  # -- 9.1. Plan - Prod. actual
  LT_opt_puerto = lead_time_opt['Puerto']
  LT_pes_puerto = lead_time_pes['Puerto']

  ws[f'K{i}'].value = f"=(I{i} - H{i}) * MAX(({leftover_days} - {LT_pes_puerto})/({LT_pes_puerto}), 0)"
  ws[f'Q{i}'].value = f"=(I{i} - H{i}) * MAX(({leftover_days} - {LT_opt_puerto})/({LT_opt_puerto}), 0)"
    
  # -- 9.2.PUERTO CHILE mes 1
  if key_month1_year in dict_agua:
    if llave in dict_agua[key_month1_year]:
      volumen_puerto_chile = dict_agua[key_month1_year][llave]['puerto chile'] or 0
      ws[f'L{i}'].value = f"={volumen_puerto_chile} * ({leftover_days}/{productive_days})" or 0
      ws[f'R{i}'].value = f"={volumen_puerto_chile} * ({leftover_days}/{productive_days})" or 0
      dict_agua[key_month1_year].pop(llave)
  
  # -- 9.3. Asignaciones mes 3
  if key_month3_year in dict_asignaciones:
    if llave in dict_asignaciones[key_month3_year]:
      ws[f'AB{i}'].value = dict_asignaciones[key_month3_year][llave]['RV final'] or 0
      dict_asignaciones[key_month3_year].pop(llave)
  
  # -- 9.6. MES N + 1
  # + ETAS
  # + inventario --> Proy mes N - Stock
  # - Ventas
  # Producción
  ws[f'X{i}'].value = f'=V{i} + W{i}'                             # PESIMISTA --> + ETA Pesimista n+1 + Inventario N
  ws[f'AA{i}'].value = f'=Y{i} + Z{i}'                            # OPTIMISTA --> + ETA Optimista n+1 + Inventario N
  ws[f'X{i}'].fill = PatternFill("solid", fgColor=yellow)
  ws[f'AA{i}'].fill = PatternFill("solid", fgColor=yellow)

  # -- 9.7. MES N + 2
  porcentaje = dict_porcentaje_produccion[oficina.lower()]
  ws[f'AD{i}'].value = f'= {porcentaje} * AB{i} + AC{i}'           # PESIMISTA --> Asignación de venta + ETA Pesimista n+2
  ws[f'AF{i}'].value = f'= {porcentaje} * AB{i} + AE{i}'           # OPTIMISTA --> Asignación de venta + ETA Optimista n+2
  ws[f'AD{i}'].fill = PatternFill("solid", fgColor=yellow)
  ws[f'AF{i}'].fill = PatternFill("solid", fgColor=yellow)
  
  # VENTA LOCAL
  if canal_distribucion == "Venta Local":
    # -- ETA: Stock planta	Puerto Chile	Centro Agua
    ws[f'J{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$R$3:R{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$5)"
    ws[f'P{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$H$3:H{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$5)"

    ws[f'V{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$S$3:S{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$R$3:R{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$7)"
    ws[f'Y{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$I$3:I{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$H$3:H{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$7)"
    
    ws[f'AC{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$T$3:T{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$S$3:S{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$8)"
    ws[f'AE{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$J$3:J{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$I$3:I{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$8)"

    # -- Stock Puerto Oficina y Almacen
    if llave in dict_stock:
      stock_puerto_oficina = dict_stock[llave]['Puerto oficina'] or 0
      stock_almacen = dict_stock[llave]['Almacen'] or 0
      ws[f'M{i}'].value = f"={stock_puerto_oficina} * ({leftover_days} / {productive_days})"
      ws[f'N{i}'].value = f"={stock_almacen} * ({leftover_days} / {productive_days})"

      ws[f'S{i}'].value = f"={stock_puerto_oficina} * ({leftover_days} / {productive_days})"
      ws[f'T{i}'].value = f"={stock_almacen} * ({leftover_days} / {productive_days})"
      dict_stock.pop(llave, None)

    # -- Proyecciones mes N
    ws[f'O{i}'].value = f'=H{i} + N{i} + J{i}'                    # PESIMISTA --> Venta Actual + Almacen oficina + ETA Pesimista n
    ws[f'O{i}'].fill = PatternFill("solid", fgColor=yellow)
    ws[f'U{i}'].value = f'=H{i} + T{i} + P{i}'                    # OPTIMISTA --> Venta Actual + Almacen oficina + ETA Optimista n
    ws[f'U{i}'].fill = PatternFill("solid", fgColor=yellow)
    
    if leftover_days >= last_stop_LT:
      ws[f'U{i}'].value = f'=H{i} + T{i} + S{i} + P{i}'           # OPTIMISTA --> + Puerto Oficina
      ws[f'U{i}'].fill = PatternFill("solid", fgColor=lightGreen)
  
  # VENTA DIRECTA
  elif canal_distribucion == "Venta Directa": # VENTA EN PUERTO CHILE 
    ws[f'J{i}'].value = f"=SUMIF('{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$R$3:R{ETA_maxRow})"
    ws[f'P{i}'].value = f"=SUMIF('{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$H$3:H{ETA_maxRow})"

    ws[f'V{i}'].value = f"=SUMIF('{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$S$3:S{ETA_maxRow})"
    ws[f'Y{i}'].value = f"=SUMIF('{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$I$3:I{ETA_maxRow})"

    ws[f'AC{i}'].value = f"=SUMIF('{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$T$3:T{ETA_maxRow})"
    ws[f'AE{i}'].value = f"=SUMIF('{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$J$3:J{ETA_maxRow})"

    # -- Proyecciones mes N
    ws[f'O{i}'].value = f'=H{i} + J{i} + K{i} + L{i}'                    # PESIMISTA --> Venta Actual + ETA + Puerto Chile Pes. - (Plan - Prod. actual)
    ws[f'O{i}'].fill = PatternFill("solid", fgColor=yellow)
    ws[f'U{i}'].value = f'=H{i} + P{i} + Q{i} + R{i}'                    # OPTIMISTA --> Venta Actual + ETA + Puerto Chile Opt.
    ws[f'U{i}'].fill = PatternFill("solid", fgColor=yellow)

  # ----- Styles
  thin = Side(border_style="thin", color=white)
  line_blue = Side(border_style="thin", color=blue)

  for j in range(8, ws.max_column + 1):
    ws[f'{get_column_letter(j)}{i}'].number_format = BUILTIN_FORMATS[3]
  
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
  ws[f'V{i}'].border = Border(left=line_blue)
  ws[f'AB{i}'].border = Border(left=line_blue)
  ws[f'AG{i}'].border = Border(left=line_blue)

  # Bold optimista y pesimista
  ws[f'O{i}'].font = Font(bold=True)
  ws[f'U{i}'].font = Font(bold=True)
  ws[f'X{i}'].font = Font(bold=True)
  ws[f'AA{i}'].font = Font(bold=True)
  ws[f'AD{i}'].font = Font(bold=True)
  ws[f'AF{i}'].font = Font(bold=True)

  # Merge 
  ws.merge_cells('J1:U1')
  ws.merge_cells('V1:AA1')
  ws.merge_cells('AB1:AF1')

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
    13: value['Puerto oficina'] or 0,     # 
    14: value['Almacen'] or 0,            # 
    19: value['Puerto oficina'] or 0,     # 
    20: value['Almacen'] or 0,            # 
  })
  
  if of.lower() in dict_lead_time['optimista']['Venta Local']:
    i += 1
    # -- ETA: Stock planta	Puerto Chile	Centro Agua
    ws[f'J{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$R$3:R{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$5)"
    ws[f'P{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$H$3:H{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$5)"

    ws[f'V{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$S$3:S{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$R$3:R{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$7)"
    ws[f'Y{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$I$3:I{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$H$3:H{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$7)"
    
    ws[f'AC{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$T$3:T{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$S$3:S{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$AA$3:AA{ETA_maxRow},'{sheet_name}'!$AH$8)"
    ws[f'AE{i}'].value = f"=SUMIFS('{sheet_name_ETA}'!$J$3:J{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$5) + SUMIFS('{sheet_name_ETA}'!$I$3:I{ETA_maxRow},'{sheet_name_ETA}'!$F$3:F{ETA_maxRow},'{sheet_name}'!C{i},'{sheet_name_ETA}'!$Q$3:Q{ETA_maxRow},'{sheet_name}'!$AH$8)"

    # -- Stock Puerto Oficina y Almacen
    if llave in dict_stock:
      ws[f'M{i}'].value = dict_stock[llave]['Puerto oficina'] or 0
      ws[f'N{i}'].value = dict_stock[llave]['Almacen'] or 0

      ws[f'S{i}'].value = dict_stock[llave]['Puerto oficina'] or 0
      ws[f'T{i}'].value = dict_stock[llave]['Almacen'] or 0
      dict_stock.pop(llave, None)

    # -- Proyecciones mes N
    ws[f'O{i}'].value = f'=H{i} + N{i} + J{i}'                    # PESIMISTA --> Venta Actual + Almacen oficina + ETA Pesimista n
    ws[f'O{i}'].fill = PatternFill("solid", fgColor=yellow)
    ws[f'U{i}'].value = f'=H{i} + T{i} + P{i}'                    # OPTIMISTA --> Venta Actual + Almacen oficina + ETA Optimista n
    ws[f'U{i}'].fill = PatternFill("solid", fgColor=yellow)
    
    # -- Proyecciones mes N + 1
    ws[f'X{i}'].value = f'=V{i} + W{i}'                             # PESIMISTA --> + ETA Pesimista n+1 + Inventario N
    ws[f'AA{i}'].value = f'=Y{i} + Z{i}'                            # OPTIMISTA --> + ETA Optimista n+1 + Inventario N

    # -- Proyecciones mes N + 2
    porcentaje = dict_porcentaje_produccion[oficina.lower()]
    ws[f'AD{i}'].value = f'= {porcentaje} * AB{i} + AC{i}'           # PESIMISTA --> Asignación de venta + ETA Pesimista n+2
    ws[f'AF{i}'].value = f'= {porcentaje} * AB{i} + AE{i}'           # OPTIMISTA --> Asignación de venta + ETA Optimista n+2

  if month_year in dict_asignaciones:
    if key in dict_asignaciones:
      ws[f'AB{i}'].value = dict_asignaciones[month_year][key]['RV final'] 

# ----- 11. Guardar la información
run_styles(ws)

ws['AH5'].value = "SI"
ws['AH6'].value = f"Mes {month_1.month}"
ws['AH7'].value = f"Mes {month_2.month}"
ws['AH8'].value = f"Mes {month_3.month}"

ws['AH3'].fill = PatternFill("solid", fgColor=yellow)
ws['AI3'].value = 'Se suma los pedidos de puerto oficina'

ws['AH4'].fill = PatternFill("solid", fgColor=lightGreen)

ws['AI4'].value = 'Se suma los pedidos que llegan este mes de agua'
print("--- %s 11. ---" % (time.time() - start_time))

wb.save(filename)
wb.close()
print("--- %s seconds ---" % (time.time() - start_time))
# messageBox(dict_lead_time, 'Venta Local', filename_dias, path_img)