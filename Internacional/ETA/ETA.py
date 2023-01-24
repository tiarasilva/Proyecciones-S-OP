from openpyxl import load_workbook
from ETA.styles import *
from constants import *

import time
import calendar
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import holidays
start_time = time.time()

def create_ETA(ws, dict_lead_time, selected_tipo_venta, date_selected_month, dict_cierre_venta):
  # HOLIDAYS
  print("--- %s ETA 1 ---" % (time.time() - start_time))
  today = datetime.now()
  last_year = today.year - 1
  this_year = today.year
  next_year = today.year + 1
  dict_holidays = {
    last_year: {
      'agro america': holidays.US(years=last_year),         # USA
      'agro europa': holidays.IT(years=last_year),          # ITALIA
      'agro mexico': holidays.MX(years=last_year),          # Mexico
      'agrosuper shanghai': holidays.CN(years=last_year),   # China
      'andes asia': holidays.KR(years=last_year)            # Corea del sur
    },
    today.year: {
      'agro america': holidays.US(years=this_year),         # USA
      'agro europa': holidays.IT(years=this_year),          # ITALIA
      'agro mexico': holidays.MX(years=this_year),          # Mexico
      'agrosuper shanghai': holidays.CN(years=this_year),   # China
      'andes asia': holidays.KR(years=this_year)            # Corea del sur
    },
    next_year: {
      'agro america': holidays.US(years=next_year),         # USA
      'agro europa': holidays.IT(years=next_year),          # ITALIA
      'agro mexico': holidays.MX(years=next_year),          # Mexico
      'agrosuper shanghai': holidays.CN(years=next_year),   # China
      'andes asia': holidays.KR(years=next_year)            # Corea del sur
    }
  }

  print("--- %s ETA 2 ---" % (time.time() - start_time))
  # DATA
  ws.append({
    # 6: f'=SUBTOTALES(9;Tabla14[Mes {selected_month}])'
    7: 'OPTIMISTA',
    17: 'PESIMISTA',
  })

  month_1 = date_selected_month
  month_2 = date_selected_month + relativedelta(months=1)
  month_3 = date_selected_month + relativedelta(months=2)
  month_4 = date_selected_month + relativedelta(months=3)

  name_month_1 = month_translate_EN_CL[month_1.strftime('%B').lower()]
  name_month_2 = month_translate_EN_CL[month_2.strftime('%B').lower()]
  name_month_3 = month_translate_EN_CL[month_3.strftime('%B').lower()]
  name_month_4 = month_translate_EN_CL[month_4.strftime('%B').lower()]

  ws.append({
    1: 'Sector', 
    2: 'Pedido',
    3: 'Oficina', 
    4: 'Material', 
    5: 'Llave',
    6: 'ETA',
    7: f'Centro Agua {name_month_1}',
    8: f'Centro Agua {name_month_2}',
    9: f'Centro Agua {name_month_3}',
    10: f'Centro Agua {name_month_4}',
    11: 'Lead time Destino',
    12: 'Fecha final Destino',
    13: 'Lead time Almacen',
    14: 'Fecha final Almacen',
    15: 'Leftover days',
    16: 'Considerar PES.',
    17: f'Centro Agua {name_month_1}',
    18: f'Centro Agua {name_month_2}',
    19: f'Centro Agua {name_month_3}',
    20: f'Centro Agua {name_month_4}',
    21: 'Lead time Destino',
    22: 'Fecha final Destino',
    23: 'Lead time Almacen',
    24: 'Fecha final Almacen',
    25: 'Leftover days',
    26: 'Considerar OPT',
  })

  print("--- %s ETA 3 ---" % (time.time() - start_time))

  # ----- Filtrar información
  wb = load_workbook(filename_ETA, read_only=True, data_only=True)
  ws_ETA = wb['ETA']
  ws_ETA_max_row = ws_ETA.max_row
  dict_status = {
    'despachado': 'Puerto',
    'embarcado': 'Agua',
    'a programar': 'Planta',
    'programado': 'Planta'
  }
  dict_leftover_date_opt = {}
  dict_leftover_date_pes = {}

  i = 3
  print("--- %s ETA 4 ---" % (time.time() - start_time))
  for row in ws_ETA.iter_rows(4, ws_ETA_max_row, values_only=True):
    pedido = row[0]
    eta = row[10]
    material = row[11]
    oficina = row[21]
    n_sector = row[28]
    kilos = row[33]
    eta = datetime.date(eta)

    if oficina.lower() in dict_lead_time['optimista'][selected_tipo_venta.lower()]:
      ws[f'A{i}'].value = dict_sector_numero[n_sector]
      ws[f'B{i}'].value = pedido
      ws[f'C{i}'].value = oficina
      ws[f'D{i}'].value = material
      ws[f'E{i}'].value = f'{oficina.lower()}{material}'
      ws[f'F{i}'].value = eta

      # OPTIMISTA
      lead_time_opt = dict_lead_time['optimista'][selected_tipo_venta.lower()][oficina.lower()]
      LT_agua = lead_time_opt['Destino']
      LT_destino = lead_time_opt['Almacen']
      tiempo_final = eta + timedelta(LT_agua) + timedelta(LT_destino)

      if tiempo_final.month == date_selected_month.month:
        ws[f'G{i}'].value = kilos
      
      elif tiempo_final.month == month_2.month:
        ws[f'H{i}'].value = kilos
      
      elif tiempo_final.month == month_3.month:
        ws[f'I{i}'].value = kilos
      
      else:
        ws[f'J{i}'].value = kilos
        ws[f'N{i}'].fill = PatternFill("solid", fgColor=yellow)

      ws[f'K{i}'].value = LT_agua
      ws[f'L{i}'].value = eta + timedelta(LT_agua)
      ws[f'M{i}'].value = LT_destino
      ws[f'N{i}'].value = tiempo_final

      # PESIMISTA
      lead_time_pes = dict_lead_time['pesimista'][selected_tipo_venta.lower()][oficina.lower()]
      LT_agua = lead_time_pes['Destino']
      LT_destino = lead_time_pes['Almacen']
      tiempo_final_pes = eta + timedelta(LT_agua) + timedelta(LT_destino)

      if tiempo_final_pes.month == date_selected_month.month:
        ws[f'Q{i}'].value = kilos
      
      elif tiempo_final.month == month_2.month:
        ws[f'R{i}'].value = kilos
      
      elif tiempo_final.month == month_3.month:
        ws[f'S{i}'].value = kilos
      
      else:
        ws[f'T{i}'].value = kilos
        ws[f'T{i}'].fill = PatternFill("solid", fgColor=yellow)

      ws[f'U{i}'].value = LT_agua
      ws[f'V{i}'].value = eta + timedelta(LT_agua)
      ws[f'W{i}'].value = LT_destino
      ws[f'X{i}'].value = tiempo_final_pes

    # ---- Calendario leftover days for month
      # OPTIMISTA
      holidays_country = dict_holidays[tiempo_final.year][oficina.lower()]

      leftover_days = 0
      if tiempo_final in dict_leftover_date_opt:
        leftover_days = dict_leftover_date_opt[tiempo_final]
      else:
        last_day_month = calendar.monthrange(tiempo_final.year, tiempo_final.month)[1]
        for day in range(tiempo_final.day + 1, last_day_month + 1):
          date_day = date(tiempo_final.year, tiempo_final.month, day)
          if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
            leftover_days += 1
        dict_leftover_date_opt[tiempo_final] = leftover_days
      
      ws[f'O{i}'].value = leftover_days

      if leftover_days > dict_cierre_venta[oficina.lower()]:
        ws[f'P{i}'].value = 'SI'
      else:
        ws[f'P{i}'].value = f'Mes {tiempo_final.month + 1}'
      
      # PESIMISTA
      leftover_days_pes = 0
      if tiempo_final_pes in dict_leftover_date_pes:
        leftover_days_pes = dict_leftover_date_pes[tiempo_final_pes]
      else:
        last_day_month = calendar.monthrange(tiempo_final.year, tiempo_final.month)[1]
        for day in range(tiempo_final_pes.day + 1, last_day_month + 1):
          date_day = date(tiempo_final.year, tiempo_final.month, day)
          if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
            leftover_days_pes += 1
        dict_leftover_date_pes[tiempo_final_pes] = leftover_days_pes
      
      ws[f'Y{i}'].value = leftover_days_pes

      if leftover_days_pes > dict_cierre_venta[oficina.lower()]:
        ws[f'Z{i}'].value = 'SI'
      else:
        ws[f'Z{i}'].value = f'Mes {tiempo_final_pes.month + 1}'

      i += 1
  print("--- %s ETA 5 ---" % (time.time() - start_time))
# ----- Corremos estilos y cerramos
  wb.close()
  run_styles(ws)
  run_number_format(ws)
  print("--- %s ETA 6 ---" % (time.time() - start_time))