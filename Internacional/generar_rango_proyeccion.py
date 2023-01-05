from openpyxl import Workbook, load_workbook

from styles import run_styles
from constants import *
import time
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

for office in dict_parametros_venta[tipo_venta.lower()]:
  ws = wb.create_sheet(office)
  print(wb.sheetnames)
  ws.append(['Sector', 'Material', 'Descripción', 'Venta plan', 'Stock planta', 'Puerto Chile', 'Centro Agua', 'Puerto Oficina', 'Almacen oficina','Pesimista Proy.', 'Optimista. Proy.'])
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
  if oficina in wb.sheetnames:
    ws = wb[oficina]
    ws.append({1: sector, 2: material, 3: descripcion, 4: venta_total})


# ----- 4. Puerto Chile, Centro Agua, Puerto Oficina, Almacén Oficina
wb_puerto = load_workbook(filename_puerto, data_only=True, read_only=True)
ws_puerto = wb_puerto.active

# for row in ws_puerto.iter_rows(7, ws_puerto.max_row, values_only=True):


# ----- 5. Stock planta

wb.save(filename)
wb.close()
print("--- %s seconds ---" % (time.time() - start_time))