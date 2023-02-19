from openpyxl import load_workbook
from datetime import datetime

rut=r'crudDatos.xlsx'

def actualizar(ruta:str,identificador:int,datosActualizados:dict):
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
      for d in datosActualizados:
        if d=='nombre' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=nombre).value=datosActualizados[d]
        elif d=='categoria' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=categoria).value=datosActualizados[d]
        elif d=='precio' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=precio).value=datosActualizados[d]
        elif d=='cantidad' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=cantidad).value=datosActualizados[d]
  archivoExcel.save(ruta)
  if encontro == False:
    print('Error: no existe una tarea con ese id')
    print()
  return