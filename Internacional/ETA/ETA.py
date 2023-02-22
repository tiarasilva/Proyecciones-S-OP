from openpyxl import load_workbook
from ETA.styles import *
from constants import *

import time
import calendar
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import holidays
start_time = time.time()

def create_ETA(ws, dict_lead_time, date_selected_month, dict_cierre_venta, dict_holidays, filename_logistica, filename_pedidos_confirmados):
  # DATA
  ws.append({
    # 6: f'=SUBTOTALES(9;Tabla14[Mes {selected_month}])'
    8: 'OPTIMISTA',
    18: 'PESIMISTA',
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
    1: 'Sector',                          # A
    2: 'Canal de Distribución',           # B
    3: 'Pedido',                          # C
    4: 'Oficina',                         # D
    5: 'Material',                        # E
    6: 'Llave',                           # F
    7: 'ETA',                             # G
    8: f'Centro Agua {name_month_1}',     # H
    9: f'Centro Agua {name_month_2}',     # I
    10: f'Centro Agua {name_month_3}',    # J
    11: f'Centro Agua {name_month_4}',    # K
    12: 'Lead time 1',                    # L
    13: 'Fecha final 1',                  # M 
    14: 'Lead time 2',                    # N
    15: 'Fecha final 2',                  # O
    16: 'Leftover days',                  # P
    17: 'Considerar PES.',                # Q
    18: f'Centro Agua {name_month_1}',    # R
    19: f'Centro Agua {name_month_2}',    # S
    20: f'Centro Agua {name_month_3}',    # T
    21: f'Centro Agua {name_month_4}',    # U
    22: 'Lead time 1',                    # V
    23: 'Fecha final 1',                  # W
    24: 'Lead time 2',                    # X
    25: 'Fecha final 2',                  # Y
    26: 'Leftover days',                  # Z
    27: 'Considerar OPT',                 # AA
  })

  # ----- 1. File Planificiación Industrial - NO USO
  # print("--- %s ETA 5.0  ---" % (time.time() - start_time))
  # wb_PI = load_workbook(f"Inputs/ETA/{filename_PI}", read_only=True, data_only=True)
  # ws_PI = wb_PI['Consolidado']
  # dict_PI = {}

  # print("--- %s ETA 5.1  ---" % (time.time() - start_time))
  # ws_PI_max = ws_PI.max_row
  # for row in ws_PI.iter_rows(2, ws_PI_max, values_only=True):
  #   centro_productivo = row[1]
  #   material = row[2]
  #   kilos = row[3]
  #   fecha_disp = row[6]
  # wb_PI.close()

  # ----- 2. File Distribución II - NO USO
  # print("--- %s ETA 5.1.1  ---" % (time.time() - start_time))
  # wb_distr = load_workbook(f"Inputs/ETA/{filename_distribucion_II}", read_only=True, data_only=True)
  # ws_distr = wb_distr['Terrestres']
  # ws_distr_max = ws_distr.max_row
  
  # i = 2
  # print("--- %s ETA 5.2  ---" % (time.time() - start_time))
  # for row in ws_distr.iter_rows(2, ws_distr_max, values_only=True):
  #   if row[0] is None:
  #     break
  #   oficina = row[0]
  #   pedido = int(row[1])
  #   transporte_terrestre = int(row[2])
  #   material = int(row[3])
  #   descripcion = row[4] #
  #   cantidad_pedida = row[5]
  #   UM = row[6]
  #   estado = row[7]
  #   fecha_programa = row[8]
  #   fecha_despacho = row[9] # Es ETD o Fec- Real D (N) o Ini- Stack (O), F Conf Logista, Fecha Programa
  #   fecha_factura = row[10] #
  #   sector = row[11]
  #   pais_destino = row[12]
  #   puerto_destino = row[13]
  #   cod_solicitante = row[14] #
  #   nombre_solicitante = row[15]
  #   incoterm = row[16]  # Es INCO?
  #   nivel_1 = row[17]
  # wb_distr.close()

  # ----- 3. File Confirmados
  wb_confirmados = load_workbook(filename_pedidos_confirmados, read_only=True, data_only=True)
  ws_confirmados = wb_confirmados["CONF - AP (zarpe mes n y n+1)"]
  ws_confirmados_max = ws_confirmados.max_row
  dict_confirmados = {}
  i = 3

  print("--- %s ETA 5.3  ---" % (time.time() - start_time))
  for row in ws_confirmados.iter_rows(3, ws_confirmados_max, values_only=True):
    if row[0] is None:
      break
    oficina = row[0]
    meanwhile_oficina = oficina.split(' ', 1)
    codigo_oficina = meanwhile_oficina[0]
    oficina = meanwhile_oficina[1]

    sector = row[1]
    pedido = row[2]
    status = row[3]
    pto_destino = row[4]
    tipo_venta = row[5]
    material = row[6]
    descripcion = row[7]
    nivel_2 = row[8]
    fecha_programa = row[10]
    kilos_stock = row[12]
    kilos_proyectado = row[13]
    total_kilos = row[14]
    ETD = row[15]
    ETA = row[16]
    NAVE = row[17]
    fecha_carga = row[17]

    # !!!!! QUE HAGO CON LOS REPETIDOS!!! ---> Se borra el AP Confirmados
    key = f"{oficina.lower()}{pedido}{material}"

    if key in dict_confirmados:
      print(f"Repetidoo {oficina}, pedido: {pedido}, material: {material}\n")
    
    else:
      dict_confirmados[key] = {
        'oficina': oficina,
        'sector': sector,
        'pedido': pedido,
        'tipo venta': tipo_venta,
        'material': material,
        'descripcion': descripcion,
        'nivel 2': nivel_2,
        'kilos stock': kilos_stock,
        'kilos proyectado': kilos_proyectado,
        'total kilos': total_kilos,
        'ETD': ETD,
        'ETA': ETA,
      } 
    
    canal_distribucion = 'Venta Directa'

    if oficina.lower() in dict_lead_time['optimista']['Venta Local']:
      canal_distribucion = 'Venta Local'

    if ETA:
      ws[f'A{i}'].value = dict_sector_numero[int(sector)]
      ws[f'B{i}'].value = canal_distribucion
      ws[f'C{i}'].value = pedido
      ws[f'D{i}'].value = oficina
      ws[f'E{i}'].value = material
      ws[f'F{i}'].value = f'{oficina.lower()}{material}'
      ws[f'G{i}'].value = datetime.date(ETA)
      ws[f'H{i}'].value = total_kilos
      i += 1
    wb_confirmados.close()
  
  # ----- 4. File Logistica
  wb_logistica = load_workbook(filename_logistica, read_only=True, data_only=True)
  ws_logistica = wb_logistica['Pedidos Planta-Puerto-Embarcado']
  ws_logistica_max = ws_logistica.max_row
  dict_logistica = {}
  j = ws.max_row

  print("--- %s ETA 5.4  ---" % (time.time() - start_time))
  for row in ws_logistica.iter_rows(2, ws_logistica_max, values_only=True):
    if row[1] is None:
      break
    oficina = row[1]
    tipo_venta = row[2]
    pedido = row[3]
    status_pedido = row[4]
    material = row[5]
    nave = row[6]
    pto_destino = row[7]
    fecha_despacho_real = row[8]
    ETD = row[9]
    ETA = row[10]
    naviera = row[11]
    kilos = row[12]
    ubicacion = row[13]
    key = f"{oficina.lower()}{pedido}{material}"
    canal_distribucion = 'Venta Directa'

    if oficina.lower() in dict_lead_time['optimista']['Venta Local']:
      canal_distribucion = 'Venta Local'

    if key in dict_confirmados:
      print(f"Repetido Confirmados con Logistica {oficina}, pedido: {pedido}, material: {material}")

    else:
      if ETA and type(ETA) != str:
        ws[f'A{j}'].value = dict_sector_numero[int(sector)]
        ws[f'B{j}'].value = canal_distribucion
        ws[f'C{j}'].value = pedido
        ws[f'D{j}'].value = oficina
        ws[f'E{j}'].value = material
        ws[f'F{j}'].value = f'{oficina.lower()}{material}'
        ws[f'G{j}'].value = datetime.date(ETA)
        ws[f'H{j}'].value = kilos
        j += 1
    
      # !!!!! QUE HAGO CON LOS REPETIDOS!!! ---> Se borra el AP Confirmados
      if key in dict_logistica:
        if ETA != dict_logistica[key]['ETA']:
          print(f"Repetido Logistica {oficina}, pedido: {pedido}, material: {material}")
        else:
          dict_logistica[key]['kilos'] += kilos
          ws[f'H{j}'].value = dict_logistica[key]['kilos']
    
      else:
        dict_logistica[key] = {
          'oficina': oficina,
          'tipo_venta': tipo_venta,
          'pedido': pedido,
          'status_pedido': status_pedido,
          'material': material,
          'nave': nave,
          'pto_destino': pto_destino,
          'fecha_despacho_real': fecha_despacho_real,
          'ETD': ETD,
          'ETA': ETA,
          'kilos': kilos,
          'ubicacion': ubicacion,
        }
  wb_logistica.close()

  # ----- 5. Calculo FECHAS
  dict_leftover_date = {}

  print("--- %s ETA 5.5  ---" % (time.time() - start_time))
  for i, row in enumerate(ws.iter_rows(3, ws.max_row - 1, values_only=True), 3):
    canal_distribucion = row[1]
    oficina = row[3]
    material = row[4]
    ETA = row[6]
    kilos = ws[f'H{i}'].value
    ws[f'H{i}'].value = ''

    # ----- OPTIMISTA -----
    lead_time_opt = dict_lead_time['optimista'][canal_distribucion][oficina.lower()]
    LT_1 = lead_time_opt['Planta']
    LT_2 = lead_time_opt['Puerto']
    tiempo_final_opt = ETA + timedelta(LT_1) + timedelta(LT_2)
    ws[f'L{i}'].value = LT_1
    ws[f'M{i}'].value = ETA + timedelta(LT_1)
    ws[f'N{i}'].value = LT_2

    if canal_distribucion == 'Venta Local':
      LT_3 = lead_time_opt['Agua']
      LT_4 = lead_time_opt['Destino']
      LT_5 = lead_time_opt['Almacen']
      tiempo_final_opt = ETA + timedelta(LT_4) + timedelta(LT_5)
      ws[f'L{i}'].value = LT_4
      ws[f'M{i}'].value = ETA + timedelta(LT_4)
      ws[f'N{i}'].value = LT_5
      
    if tiempo_final_opt.month == month_1.month:
      ws[f'H{i}'].value = kilos
    
    elif tiempo_final_opt.month == month_2.month:
      ws[f'I{i}'].value = kilos
    
    elif tiempo_final_opt.month == month_3.month:
      ws[f'J{i}'].value = kilos
    
    else:
      ws[f'K{i}'].value = kilos
      ws[f'O{i}'].fill = PatternFill("solid", fgColor=yellow)

    ws[f'O{i}'].value = tiempo_final_opt
    
    # ----- PESIMISTA -----
    lead_time_pes = dict_lead_time['pesimista'][canal_distribucion][oficina.lower()]
    LT_1 = lead_time_pes['Planta']
    LT_2 = lead_time_pes['Puerto']
    tiempo_final_pes = ETA + timedelta(LT_1) + timedelta(LT_2)
    ws[f'V{i}'].value = LT_1
    ws[f'W{i}'].value = ETA + timedelta(LT_1)
    ws[f'X{i}'].value = LT_2

    if canal_distribucion == 'Venta Local':
      LT_3 = lead_time_pes['Agua']
      LT_4 = lead_time_pes['Destino']
      LT_5 = lead_time_pes['Almacen']
      ws[f'V{i}'].value = LT_4
      ws[f'W{i}'].value = ETA + timedelta(LT_4)
      ws[f'X{i}'].value = LT_5

    if tiempo_final_pes.month == month_1.month:
      ws[f'R{i}'].value = kilos
    
    elif tiempo_final_pes.month == month_2.month:
      ws[f'S{i}'].value = kilos
    
    elif tiempo_final_pes.month == month_3.month:
      ws[f'T{i}'].value = kilos
    
    else:
      ws[f'U{i}'].value = kilos
      ws[f'U{i}'].fill = PatternFill("solid", fgColor=yellow)

    ws[f'Y{i}'].value = tiempo_final_pes

    # ---- Calendario leftover days for month
    for tiempo_final in [tiempo_final_pes, tiempo_final_opt]:
      holidays_country = dict_holidays[tiempo_final.year]['chile']
      if canal_distribucion == 'Venta Local':
        holidays_country = dict_holidays[tiempo_final.year][oficina.lower()]
      leftover_days = 0

      if oficina.lower() in dict_leftover_date.keys():
        if tiempo_final in dict_leftover_date[oficina.lower()]:
          leftover_days = dict_leftover_date[oficina.lower()][tiempo_final]

        else:
          last_day_month = calendar.monthrange(tiempo_final.year, tiempo_final.month)[1]
          for day in range(tiempo_final.day + 1, last_day_month + 1):
            date_day = date(tiempo_final.year, tiempo_final.month, day)
            if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
              leftover_days += 1
          dict_leftover_date[oficina.lower()][tiempo_final] = leftover_days

      else:
        dict_leftover_date[oficina.lower()] = {}
        last_day_month = calendar.monthrange(tiempo_final.year, tiempo_final.month)[1]
        for day in range(tiempo_final.day + 1, last_day_month + 1):
          date_day = date(tiempo_final.year, tiempo_final.month, day)
          if date_day not in holidays_country and date_day.strftime('%A') != 'Sunday':
            leftover_days += 1
        dict_leftover_date[oficina.lower()][tiempo_final] = leftover_days

    leftover_opt = dict_leftover_date[oficina.lower()][tiempo_final_opt]
    ws[f'P{i}'].value = leftover_opt

    if leftover_opt > dict_cierre_venta[oficina.lower()]:
      ws[f'Q{i}'].value = 'SI'
    else:
      ws[f'Q{i}'].value = f'Mes {tiempo_final_opt.month + 1}'
    
    leftover_pes = dict_leftover_date[oficina.lower()][tiempo_final_pes]
    ws[f'Z{i}'].value = leftover_pes

    if leftover_pes > dict_cierre_venta[oficina.lower()]:
      ws[f'AA{i}'].value = 'SI'
    else:
      ws[f'AA{i}'].value = f'Mes {tiempo_final_pes.month + 1}'

  
  print("--- %s ETA 5 ---" % (time.time() - start_time))
# ----- Corremos estilos y cerramos
  run_styles(ws)
  # run_number_format(ws)
  print("--- %s ETA 6 ---" % (time.time() - start_time))