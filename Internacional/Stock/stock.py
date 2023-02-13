from openpyxl import load_workbook
from Stock.styles import *
from constants import *

import calendar
from datetime import datetime, date
import holidays

# ----- Creamos el sheet
def stock(ws, dict_lead_time):
  dict_message = {'Puerto Oficina': {}, 'Almacen': {}}

  ws.append({
    6: 'Puerto Oficina', 
    14: 'Almacén Oficina'
  })

  ws.append({
    6: 'Stock liberado', 10: 'Stock no liberado',    # Puerto Oficina
    14: 'Stock liberado', 18: 'Stock no liberado',   # Almacen Oficina
  })
  
  ws.append({
    1: 'Sector', 
    2: 'Oficina', 
    3: 'Material', 
    4: 'Descripción', 
    5: 'Nivel 2', 
    6: 'Total KG', 7: 'Nº Días de Antigüedad Centro', 8: 'Nº Días de Antigüedad Oficina', 9: 'Status considerado',
    10: 'Total KG', 11: 'Nº Días de Antigüedad Centro', 12: 'Nº Días de Antigüedad Oficina', 13: 'Status considerado',
    14: 'Total KG', 15: 'Nº Días de Antigüedad Centro', 16: 'Nº Días de Antigüedad Oficina', 17: 'Status considerado',
    18: 'Total KG', 19: 'Nº Días de Antigüedad Centro', 20: 'Nº Días de Antigüedad Oficina', 21: 'Status considerado',
  })

  # ----- Días antiguedad Stock
  wb_dias_stock = load_workbook(filename_dias, read_only=True, data_only=True)
  ws_dias_stock = wb_dias_stock['Stock']
  dict_dias_stock = {}
  obj = calendar.Calendar()

  sector = ''
  oficina = ''
  i = 4
  for row in ws_dias_stock.iter_rows(5, ws_dias_stock.max_row - 1, values_only=True):
    if row[3] is None:
      break
    
    if row[1] is not None:
      sector = row[1]
    
    if row[2] is not None:
      oficina = row[2]
    
    month_year = row[0]
    material = row[3]
    descripcion = row[4]

    stock_almacen_lib = row[5] or 0
    dias_almacen_centro_lib = row[6] or 0
    dias_almacen_lib = row[7] or 0

    stock_almacen_no_lib = row[8] or 0
    dias_almacen_centro_no_lib = row[9] or 0
    dias_almacen_no_lib = row[10] or 0

    stock_oficina_lib = row[14] or 0
    dias_oficina_centro_lib = row[15] or 0
    dias_oficina_lib = row[16] or 0

    stock_oficina_no_lib = row[17] or 0
    dias_oficina_centro_no_lib = row[18] or 0
    dias_oficina_no_lib = row[19] or 0

    llave = f'{oficina.lower()}{material}'

    if oficina.lower() in dict_lead_time['optimista']['Venta Local'].keys():
      ws[f'A{i}'].value = sector
      ws[f'B{i}'].value = oficina
      ws[f'C{i}'].value = int(material)
      ws[f'D{i}'].value = descripcion
      ws[f'F{i}'].value = stock_oficina_lib or 0
      ws[f'G{i}'].value = dias_oficina_centro_lib or 0
      ws[f'H{i}'].value = dias_oficina_lib or 0
      ws[f'J{i}'].value = stock_oficina_no_lib or 0
      ws[f'K{i}'].value = dias_oficina_centro_no_lib or 0
      ws[f'L{i}'].value = dias_oficina_no_lib or 0
      ws[f'N{i}'].value = stock_almacen_lib or 0
      ws[f'O{i}'].value = dias_almacen_centro_lib or 0
      ws[f'P{i}'].value = dias_almacen_lib or 0
      ws[f'R{i}'].value = stock_almacen_no_lib or 0
      ws[f'S{i}'].value = dias_almacen_centro_no_lib or 0
      ws[f'T{i}'].value = dias_almacen_no_lib or 0

      today = datetime.now()
      holidays_country = holidays.CL(years=today.year)

      if 'america' in oficina.lower():
        holidays_country = holidays.US(years=today.year)        # USA
      elif 'europa' in oficina.lower():
        holidays_country = holidays.IT(years=today.year)        # Italia
      elif 'mexico' in oficina.lower():
        holidays_country = holidays.MX(years=today.year)        # Mexico
      elif 'shanghai'in oficina.lower():  
        holidays_country = holidays.CN(years=today.year)        # China
      elif 'asia' in oficina.lower():
        holidays_country = holidays.KR(years=today.year)        # Corea del sur

       # Considerar de lunes a sábado sin los feriados
      business_days = 0
      sm = 12
      for day in obj.itermonthdates(today.year, sm):
        if day.month == sm:
          if day not in holidays_country and day.strftime('%A') != 'Sunday':
            business_days += 1
      
      leftover_days = business_days - today.day 
      lead_time = dict_lead_time['optimista']['Venta Local'][oficina.lower()]
    
      # PUERTO OFICINA
      # Stock liberado
      dias_oficina = dias_oficina_centro_lib + leftover_days
      if dias_oficina >= lead_time['Destino']:
        ws[f'I{i}'].value = stock_oficina_lib
        ws[f'I{i}'].font = Font(bold=True, color=green)
        ws[f'I{i}'].fill = PatternFill("solid", fgColor=lightGreen)
      else:
        ws[f'I{i}'].value = 0
        ws[f'I{i}'].font = Font(bold=True, color=darkRed)
        ws[f'I{i}'].fill = PatternFill("solid", fgColor=lightRed)
      
      # Stock no liberado
      dias_oficina = dias_oficina_centro_no_lib + leftover_days
      if dias_oficina >= lead_time['Destino']:
        ws[f'M{i}'].value = stock_oficina_no_lib
        ws[f'M{i}'].font = Font(bold=True, color=green)
        ws[f'M{i}'].fill = PatternFill("solid", fgColor=lightGreen)
      else:
        ws[f'M{i}'].value = 0
        ws[f'M{i}'].font = Font(bold=True, color=darkRed)
        ws[f'M{i}'].fill = PatternFill("solid", fgColor=lightRed)
      
      # ALMACEN OFICINA
      # Stock liberado
      dias_almacen = dias_almacen_lib + leftover_days
      if dias_almacen >= lead_time['Almacen']:
        ws[f'Q{i}'].value = stock_almacen_lib
        ws[f'Q{i}'].font = Font(bold=True, color=green)
        ws[f'Q{i}'].fill = PatternFill("solid", fgColor=lightGreen)
      else:
        ws[f'Q{i}'].value = 0
        ws[f'Q{i}'].font = Font(bold=True, color=darkRed)
        ws[f'Q{i}'].fill = PatternFill("solid", fgColor=lightRed)
      
      dias_almacen = dias_almacen_no_lib + leftover_days
      if dias_almacen >= lead_time['Almacen']:
        ws[f'U{i}'].value = stock_almacen_no_lib
        ws[f'U{i}'].font = Font(bold=True, color=green)
        ws[f'U{i}'].fill = PatternFill("solid", fgColor=lightGreen)
      else:
        ws[f'U{i}'].value = 0
        ws[f'U{i}'].font = Font(bold=True, color=darkRed)
        ws[f'U{i}'].fill = PatternFill("solid", fgColor=lightRed)
      i += 1  

  wb_dias_stock.close()    

  # ----- Cerramos y guardamos
  run_styles(ws)
  run_number_format(ws)