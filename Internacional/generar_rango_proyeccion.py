from openpyxl import Workbook, load_workbook
from constants import *

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

  # print(row)
# print(dict_parametros_venta)

wb_parametros.close()
# ----- 2. Creamos el excel de resultados
wb = Workbook()

for office in dict_parametros_venta[tipo_venta.lower()]:
  ws = wb.create_sheet(office)
  ws.append(['Material', 'Stock planta', 'Puerto Chile', 'Centro Agua', 'Puerto Oficina', 'Almacen oficina','Pesimista Proy.', 'Optimista. Proy.'])

wb.save(filename)
wb.close()