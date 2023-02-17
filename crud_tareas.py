from openpyxl import load_workbook
from datetime import datetime

rut=r'crudDatos.xlsx'

def leer(ruta:str, extraer:str):
  archivoExcel=load_workbook(ruta)

  hojaDatos=archivoExcel['tareas']
  hojaDatos=hojaDatos['A2':'F'+str(hojaDatos.max_row)]

  info={}

  for i in hojaDatos:
    if isinstance(i[0].value,int):
      info.setdefault(i[0].value,{'nombre':i[1].value, 'descripcion':i[2].value,'estado':i[3].value,'fecha inicio':i[4].value,'fecha finalizacion':i[5].value})

  if not(extraer=='todo'):
    info=filtrar(info,extraer)
  for i in info:
    
    print('********** Tarea ***********')
    print('id:'+str(i)+'\n'+'nombre: '+str(info[i]['nombre'])+'\n'+'descripcion: '+str(info[i]['descripcion'])+'\n'+'estado: '+str(info[i]['estado'])+'\n'+'fecha de creacion: '+str(info[i]['fecha inicio'])+'\n'+'fecha finalizacion: '+str(info[i]['fecha finalizacion']))
    print()
  return info

def filtrar(info:dict, filtro:str):
  aux={}
  for i in info:
    if info[i]['estado']==filtro:
      aux.setdefault(i,info[i])
  return aux

def actualizar(ruta:str,identificador:int,datosActualizados:dict):
  archivoExcel=load_workbook(ruta)

  hojaDatos=archivoExcel['tareas']
  hojaDatos=hojaDatos['A2':'F'+str(hojaDatos.max_row)]
  hoja=archivoExcel.active

  nombre=2
  descripcion=3
  estado=4
  fechaInicio=5
  fechaFinalizado=6
  encontro=False
  for i in hojaDatos:
    if i[0].value==identificador:
      fila=i[0].row
      encontro=True
      for d in datosActualizados:
        if d=='nombre' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=nombre).value=datosActualizados[d]
        elif d=='descripcion' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=descripcion).value=datosActualizados[d]
        elif d=='estado' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=estado).value=datosActualizados[d]
        elif d=='fecha inicio' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=fechaInicio).value=datosActualizados[d]
        elif d=='fecha finalizado' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=fechaFinalizado).value=datosActualizados[d]
  archivoExcel.save(ruta)
  if encontro == False:
    print('Error: no existe una tarea con ese id')
    print()
  return