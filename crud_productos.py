from openpyxl import load_workbook
from datetime import datetime

rut=r'crudDatos.xlsx'



while True:
  print('******************************************')
  print('Indique la accion que desea realizar: \nConsultar: 1\nActualizar: 2\nCrear nuevo producto: 3\nBorrar: 4')
  accion =int(input('Escriba la opcion: '))
  if accion<1 or accion>4:
    print('Comando invalido, por favor eliga una opcion valida')
  elif accion==1:
    opcConsulta=''
    print('Indique la categoria del producto que desea consultar:\nTodos los productos: 1\nComputación: 2\nAlimentario: 3\nHigiene: 4\nEscolar: 5')
    opcConsulta=input('Escriba la categoria que desee consultar: ')
    if opcConsulta=='1':
      print('\n\n** Consultado todos los productos **')
      leer(rut,'todo')
    elif opcConsulta=='2':
      print('\n\n** Consultado todos los productos **')
      leer(rut,'computacion')
    elif opcConsulta=='3':
      print('\n\n** Consultado todos los productos **')
      leer(rut,'alimentario')
    elif opcConsulta=='4':
      print('\n\n** Consultado todos los productos **')
      leer(rut,'higiene')
    elif opcConsulta=='5':
      print('\n\n** Consultado todos los productos **')
      leer(rut,'escolar')
  elif accion==2:
    datosActualizados={'nombre':'','categoria':'','precio':'','cantidad':''}
    print('** Actualizar Tarea **\n')
    idActualizar=int(input('Indique el ID de el producto que desea actualizar: '))

    print('\n** Nuevo nombre **\n** Nota: si no desea actualizar el nombre solo oprima ENTER **')
    datosActualizados['nombre']=input('Indique el nuevo nombre de el producto: ')

    print('\n** Nueva categoria **\nComputación: 1\nAlimentario: 2\nHigiene: 3\nEscolar: 4\n** Nota: si no desea actualizar la categoria solo oprima ENTER **')
    
    estadoNuevo=input('Indique el nuevo estado de el producto: ')
    if estadoNuevo=='1':
      datosActualizados['categoria']='computacion'
    elif estadoNuevo=='2':
      datosActualizados['categoria']='alimentario'
    elif estadoNuevo=='3':
      datosActualizados['categoria']='higiene'
    elif estadoNuevo=='4':
      datosActualizados['categoria']='escolar'
    print('\n** Nuevo precio **\n** Nota: si no desea actualizar el precio solo oprima ENTER **')
    datosActualizados['precio']=input('Indique el nuevo precio de el producto: ')

    print('\n** Nueva cantidad **\n** Nota: si no desea actualizar la cantidad solo oprima ENTER **')
    datosActualizados['cantidad']=input('Indique la nueva cantidad de el producto: ')

    actualizar(rut,idActualizar, datosActualizados)
    print()
  elif accion==3:
    datosActualizados={'nombre':'','categoria':'','precio':'','cantidad':''}
    print('** Crear nuevo producto **\n')
    print('** Nombre **\n')
    datosActualizados['nombre']=input('Indique el nombre de el producto: ')

    print('\n** Categoria **\nComputación: 1\nAlimentario: 2\nHigiene: 3\nEscolar: 4')
    
    estadoNuevo=input('Indique la categoria de el producto: ')
    if estadoNuevo=='1':
      datosActualizados['categoria']='computacion'
    elif estadoNuevo=='2':
      datosActualizados['categoria']='alimentario'
    elif estadoNuevo=='3':
      datosActualizados['categoria']='higiene'
    elif estadoNuevo=='4':
      datosActualizados['categoria']='escolar'
    print('\n** Precio **')
    datosActualizados['precio']=input('Indique el precio de el producto: ')
    print('\n** cantidad **')
    datosActualizados['cantidad']=input('Indique el nombre de el producto: ')
    agregar(rut,datosActualizados)
  elif accion==4:
    print('\n** Eliminar producto **')
    iden=int(input('Indique el ID de el producto que desea eliminar: '))
    borrar(rut,iden)