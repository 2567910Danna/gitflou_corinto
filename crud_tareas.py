from openpyxl import load_workbook
from datetime import datetime

rut=r'crudDatos.xlsx'

while True:
  print('******************************************')
  print('Indique la accion que desea realizar: \nConsultar: 1\nActualizar: 2\nCrear nueva tarea: 3\nBorrar: 4')
  accion =int(input('Escriba la opcion: '))
  if accion<1 or accion>4:
    print('Comando invalido, por favor eliga una opcion valida')
  elif accion==1:
    opcConsulta=''
    print('Indique la tarea que desea consultar:\nTodas las tareas: 1\nEn espera: 2\nEN ejecucion: 3\nPor aprobar: 4\nFinalizada: 5')
    opcConsulta=input('Escriba la tarea que see consultar: ')
    if opcConsulta=='1':
      print('\n\n** Consultado todas las tareas **')
      leer(rut,'todo')
    elif opcConsulta=='2':
      print('\n\n** Consultado todas las tareas **')
      leer(rut,'En espera')
    elif opcConsulta=='3':
      print('\n\n** Consultado todas las tareas **')
      leer(rut,'En ejecucion')
    elif opcConsulta=='4':
      print('\n\n** Consultado todas las tareas **')
      leer(rut,'Por aprobar')
    elif opcConsulta=='5':
      print('\n\n** Consultado todas las tareas **')
      leer(rut,'Finalizada')
  elif accion==2:
    datosActualizados={'nombre':'','descripcion':'','estado':'','fecha inicio':'','fecha finalizacion':''}
    print('** Actualizar Tarea **\n')
    idActualizar=int(input('Indique el ID de la tarea que desea actualizar: '))
    print('\n** Nuevo nombre **\n** Nota: si no desea actualizar el nombre solo oprima ENTER **')
    datosActualizados['nombre']=input('Indique el nuevo nombre de la tarea: ')
    print('\n** Nueva descripcion **\n** Nota: si no desea actualizar la descripcion solo oprima ENTER **')
    datosActualizados['descripcion']=input('Indique la nueva descripcion de la tarea: ')
    print('\n** Nueva estado **\nEn espera: 2\nEn ejecucion: 3\nPor aprobar: 4\nFinalizado: 5\n** Nota: si no desea actualizar el estado solo oprima ENTER **')
    estadoNuevo=input('Indique el nuevo estado de la tarea: ')
    if estadoNuevo=='2':
      datosActualizados['estado']='En espera'
    elif estadoNuevo=='3':
      datosActualizados['estado']='En ejecucion'
    elif estadoNuevo=='4':
      datosActualizados['estado']='Por aprobar'
    elif estadoNuevo=='5':
      now=datetime.now()
      datosActualizados['estado']='Finalizada'
      datosActualizados['fecha finalizacion']=str(now.day) +'/'+str(now.month)+'/'+str(now.year)
    now=datetime.now()
    datosActualizados['fecha inicio']=str(now.day) +'/'+str(now.month)+'/'+str(now.year)
    actualizar(rut,idActualizar, datosActualizados)
    print()
  elif accion==3:
    datosActualizados={'nombre':'','descripcion':'','estado':'','fecha inicio':'','fecha finalizacion':''}
    print('** Crear nueva Tarea **\n')
    print('** Titulo **\n')
    datosActualizados['nombre']=input('Indique el nombre de la tarea: ')
    print('\n** descripcion **')
    datosActualizados['descripcion']=input('Indique la descripcion de la tarea: ')
    print()
    datosActualizados['estado']='En espera'
    now = datetime.now()
    datosActualizados['fecha inicio']=str(now.day) +'/' + str(now.month)+'/'+str(now.year)
    agregar(rut,datosActualizados)
  elif accion==4:
    print('\n** Eliminar tarea **')
    iden=int(input('Indique el ID de la tarea que desea eliminar: '))
    borrar(rut,iden)