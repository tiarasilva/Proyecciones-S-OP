from openpyxl import load_workbook
from Stock.styles import *
from constants import *

import calendar
from datetime import datetime, date
import holidays

# ----- Creamos el sheet
def stock(ws, dict_lead_time, selected_tipo_venta, selected_month):
  selected_month = month_translate_CL_EN[selected_month.lower()]
  dict_message = {'Puerto Oficina': {}, 'Almacen': {}}

  ws.append({ 1: 'Sector', 
    2: 'Oficina', 
    3: 'Material', 
    4: 'Descripción', 
    5: 'Nivel 2', 
    6: 'Puerto Oficina', 
    14: 'Almacén Oficina'
  })

  ws.append({
    6: 'Stock liberado', 10: 'Stock no liberado',    # Puerto Oficina
    14: 'Stock liberado', 18: 'Stock no liberado',   # Almacen Oficina
  })
  
  ws.append({
    6: 'Total KG', 7: 'Nº Días de Antigüedad Centro', 8: 'Nº Días de Antigüedad Oficina', 9: 'Status considerado',
    10: 'Total KG', 11: 'Nº Días de Antigüedad Centro', 12: 'Nº Días de Antigüedad Oficina', 13: 'Status considerado',
    14: 'Total KG', 15: 'Nº Días de Antigüedad Centro', 16: 'Nº Días de Antigüedad Oficina', 17: 'Status considerado',
    18: 'Total KG', 19: 'Nº Días de Antigüedad Centro', 20: 'Nº Días de Antigüedad Oficina', 21: 'Status considerado',
  })

  # ----- Días antiguedad Stock
  wb_dias_stock = load_workbook(filename_dias, read_only=True, data_only=True)
  ws_dias_stock = wb_dias_stock.active
  dict_dias_stock = {}
  obj = calendar.Calendar()

  for i, row in enumerate(ws_dias_stock.iter_rows(9, ws_dias_stock.max_row, values_only=True), 4):
    if row[1] is None:
      break
    # Stock No liberado
    sector = row[0]
    oficina = row[1]
    material = row[2]
    descripcion = row[3]
    stock_oficina_lib = row[4] or 0
    stock_oficina_no_lib = row[5] or 0
    dias_oficina_centro_lib = row[7] or 0
    dias_oficina_centro_no_lib = row[8] or 0
    oficina_dias_lib = row[10] or 0
    oficina_dias_no_lib = row[11] or 0
    stock_almacen_lib = row[13] or 0
    stock_almacen_no_lib = row[14] or 0
    dias_almacen_centro_lib = row[16] or 0
    dias_almacen_centro_no_lib = row[17] or 0
    dias_almacen_lib = row[19] or 0
    dias_almacen_oficina_no_lib = row[20] or 0
    llave = oficina.lower() + material

    ws.append({
      1: sector,
      2: oficina,
      3: int(material),
      4: descripcion,
      6: stock_oficina_lib or 0,
      7: dias_oficina_centro_lib or 0,
      8: oficina_dias_lib or 0,

      10: stock_oficina_no_lib or 0,
      11: dias_oficina_centro_no_lib or 0,
      12: oficina_dias_no_lib or 0,

      14: stock_almacen_lib or 0,
      15: dias_almacen_centro_lib or 0,
      16: dias_almacen_lib or 0,

      18: stock_almacen_no_lib or 0,
      19: dias_almacen_centro_no_lib or 0,
      20: dias_almacen_oficina_no_lib or 0,
    })

    if oficina.lower() in dict_lead_time['optimista'][selected_tipo_venta.lower()].keys():
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
      lead_time = dict_lead_time['optimista'][selected_tipo_venta.lower()][oficina.lower()]
    
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
      
      dias_almacen = dias_almacen_oficina_no_lib + leftover_days
      if dias_almacen >= lead_time['Almacen']:
        ws[f'U{i}'].value = stock_almacen_no_lib
        ws[f'U{i}'].font = Font(bold=True, color=green)
        ws[f'U{i}'].fill = PatternFill("solid", fgColor=lightGreen)
      else:
        ws[f'U{i}'].value = 0
        ws[f'U{i}'].font = Font(bold=True, color=darkRed)
        ws[f'U{i}'].fill = PatternFill("solid", fgColor=lightRed)

  wb_dias_stock.close()    

  # ----- Cerramos y guardamos
  run_styles(ws)
  run_number_format(ws)