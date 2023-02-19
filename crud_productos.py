from openpyxl import load_workbook
from datetime import datetime

rut=r'crudDatos.xlsx'

def borrar(ruta,identificador):
  archivoExcel=load_workbook(ruta)

  hojaDatos=archivoExcel['productos']
  hojaDatos=hojaDatos['A2':'E'+str(hojaDatos.max_row)]
  hoja=archivoExcel.active

  nombre=2
  categoria=3
  precio=4
  cantidad=5
  encontro=False
  for i in hojaDatos:
    if i[0].value==identificador:
      fila=i[0].row
      encontro=True
      hoja.cell(row=fila,column=1).value=''
      hoja.cell(row=fila,column=nombre).value=''
      hoja.cell(row=fila,column=categoria).value=''
      hoja.cell(row=fila,column=precio).value=''
      hoja.cell(row=fila,column=cantidad).value=''
  archivoExcel.save(ruta)
  if encontro == False:
    print('Error: no existe una tarea con ese id')
    print()
  return
