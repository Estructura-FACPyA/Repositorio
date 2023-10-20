#CREACIÓN DE UNA BASE DE DATOS CON DOS TABLAS RELACIONADAS UNO A MUCHOS
import random as rd
import sys
import datetime
import sqlite3
from sqlite3 import Error
from datetime import (date, 
                      datetime,
                      timedelta)
import openpyxl
#Crea una tabla en SQLite3



def Crear_tabla ():
  try:
    with sqlite3.connect("34.db") as conn:
          mi_cursor = conn.cursor()
          mi_cursor.execute("CREATE TABLE IF NOT EXISTS Usuarios (clave INTEGER PRIMARY KEY, nombre TEXT NOT NULL);")
          mi_cursor.execute("CREATE TABLE IF NOT EXISTS Salas (clave INTEGER PRIMARY KEY, nombre TEXT NOT NULL, capacidad INTEGER NOT NULL);")
          mi_cursor.execute("CREATE TABLE IF NOT EXISTS Reservaciones (folio INTEGER PRIMARY KEY, nombre TEXT NOT NULL, horario Text NOT NULL, fecha timestamp) ")
          print("Tablas creadas exitosamente")
  except Error as e:
      print (e)
  except:
      print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
  finally:
      conn.close()

Crear_tabla ()

# Opcion 1
def Registrar_Reservacion ():
    while True:
        valor_clave = int(input("Cual es tu clave de cliente: "))
        try:
            with sqlite3.connect("34.db") as conn:
                mi_cursor = conn.cursor()
                valores = {"clave":valor_clave}
                mi_cursor.execute("SELECT * FROM Usuarios WHERE clave = :clave", valores)
                registro = mi_cursor.fetchall()

                if registro:
                    for clave, nombre in registro:
                        print(f"{clave}\t{nombre}")
                else:
                    print(f"No se encontró un proyecto asociado con la clave {valor_clave}")
        except Error as e:
            print (e)
        except:
            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
        else:
            Nombre = input("Ingresa el nombre de la reservación(Escribe SALIR para regresar al menú): ")
            if Nombre == '':
                continue
            elif Nombre == 'SALIR':
                return
            else:
                Horario = input("¿Cual es el horario que quieres? [M, V, N]: ")
                print(Horario)
                Fecha_Ingresada= input("Ingresa la fecha de reservación: ")
                Fecha_dt = datetime.strptime(Fecha_Ingresada, '%d/%m/%Y')
                fecha_actual = datetime.today()
                fecha_permitida = datetime.now() + timedelta( days = 2)
                if Fecha_dt < fecha_permitida:
                    print("Debes hacer la reservacion con 2 dias de anticipación.")
                else:
                    nivel = rd.randint(1,99)
                    try:
                        with sqlite3.connect("34.db") as conn:
                            mi_cursor = conn.cursor()
                            Miami={"folio":nivel,"nombre":Nombre,"horario":Horario,"fecha":Fecha_dt}
                            mi_cursor.execute("INSERT INTO Reservaciones VALUES(:folio,:nombre,:horario,:fecha)",Miami)
                            #print("El folio es ")
                            #print(nivel)
                            #print("El nombre es ")
                            #print(Nombre)
                    except Error as e:
                        print (e)
                    except:
                        print(f"Surgio una falla siendo esta la causa: {sys.exc_info()[0]}")
                    finally:
                        if (conn):
                            conn.close()
                            fecha_consultar = input("Confirma la fecha (dd/mm/aaaa): ")
                            fecha_consultar = datetime.strptime(fecha_consultar, "%d/%m/%Y").date()
                            print("¡Reservacion Realizada con exito!")
                            try:
                                with sqlite3.connect("Miami.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                                    mi_cursor = conn.cursor()
                                    criterios = {"fecha":fecha_consultar}
                                    mi_cursor.execute("SELECT folio, nombre, horario, fecha FROM Reservaciones WHERE DATE(fecha) = (:fecha);", criterios)
                                    #mi_cursor.execute("SELECT clave, nombre, fecha_registro FROM Amigo WHERE DATE(fecha_registro) >= :fecha;", criterios)
                                    registros = mi_cursor.fetchall()
   
                                    for clave, nombre, fecha_registro, fecha in registros:
                                        print(f"Clave = {clave}, tipo de dato {type(clave)}")
                                        print(f"Nombre = {nombre}, tipo de dato {type(nombre)}")
                                        print(f"Horario = {fecha_registro}, tipo de dato {type(fecha_registro)}")
                                        print(f"Fecha = {fecha}, tipo de dato {type(fecha)}")
                                    for clave, nombre, fecha_registro, fecha in registros:
                                        print("Clave\t" + "Nombre\t"+ " Turno\t" + "            Fecha\t")
                                        print(f"{clave}\t {nombre}\t {fecha_registro}\t {fecha}\t")
                                        return

                            except sqlite3.Error as e:
                                print (e)
                            except Exception:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            finally:
                                if (conn):
                                    conn.close()
                                    print("Se ha cerrado la conexión")

#Opcion 2
def modificar_descripciones ():
    while True:
        llave = input("Cual es el nombre de tu sala actual: ")
        try:
            with sqlite3.connect("34.db") as conn:
                mi_cursor = conn.cursor()
                valores1 = {"nombre":llave}
                mi_cursor.execute("SELECT * FROM Reservaciones WHERE nombre = :nombre", valores1)
                registro = mi_cursor.fetchall()

                if registro:
                    for folio, nombre, horario, fecha in registro:
                        print("Clave\t" + "Nombre\t"+ " Turno\t" + "            Fecha\t")
                        print(f"{folio}\t{nombre}\t{horario}\t{fecha}")
                else:
                    print(f"No se encontró un proyecto asociado con la clave {llave}")
        except Error as e:
            print (e)
        except:
            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
        else:
            nuevo_nombre = input("A que nombre lo quieres cambiar?: ")
            id_number = folio
            Turno = horario
            fecha_dt = fecha
            try:
                with sqlite3.connect("34.db") as conn:
                    mi_cursor = conn.cursor()
                    Sydney = {"folio":id_number, "nombre":nuevo_nombre,"turno":Turno,"fecha":fecha_dt}
                    mi_cursor.execute("UPDATE Reservaciones SET nombre = (:nombre) WHERE (folio) = (:folio);", Sydney)
                    print("Modificacion realizada con exito.")
                    return
            except Error as e:
                print (e)
            except:
                print(f"Surgio una falla siendo esta la causa: {sys.exc_info()[0]}")

#opcion 3
def consulta_fecha():
    fecha_consultar = input("Dime una fecha (dd/mm/aaaa): ")
    fecha_dt = datetime.strptime(fecha_consultar, '%d/%m/%Y')
    
    try:
        with sqlite3.connect("34.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
            mi_cursor = conn.cursor()
            criterios = {"fecha":fecha_consultar}
            mi_cursor.execute("SELECT folio, nombre, horario, fecha FROM Reservaciones WHERE DATE(fecha) = :fecha;", criterios)
            registros = mi_cursor.fetchall()
            
            for folio, nombre, horario, fecha in registros:
                print(f'Estas son las fechas ocupadas {fecha.date}')
    except Error as e:
        print (e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        if (conn):
            conn.close()
            print("Se ha cerrado la conexión")
            


#Opcion 4
def reporte_a_Excel ():
    while True:
        fecha_reporte = input("¿De que fecha quieres sacar el reporte? ")
        fecha_reporte = datetime.strptime(fecha_reporte, "%d/%m/%Y")
        print("*"* 75)
        print("**            REPORTE DE RESERVACIONES PARA EL DIA", fecha_reporte, "           **")
        try:
            with sqlite3.connect("34.db", detect_types = sqlite3.PARSE_DECLTYPES  | sqlite3.PARSE_COLNAMES) as conn:
                mi_cursor = conn.cursor()
                criterios = {"fecha":fecha_reporte}
                mi_cursor.execute("SELECT folio, nombre, horario, fecha FROM Reservaciones WHERE fecha = :fecha", criterios)
                registrados = mi_cursor.fetchall()
                
                if registrados:
                    for folio, nombre, horario, fecha in registrados:
                        print("Clave\t" + "Nombre\t"+ " Turno\t" + "            Fecha\t")
                        print(f"{folio}\t{nombre}\t{horario}\t{fecha}")
                        lista = []
                        lista.append (folio, nombre, horario, fecha)
                        print(lista)
                        wb = openpyxl.Workbook()
                        hoja = wb.active
                        hoja = wb.active
                        
                        for listas in lista:
                            hoja.append(listas)
                        print(f'Hoja activa: {hoja.title}')
                        hoja["A1"] = 10
                        a1 = hoja["A1"]
                        print(a1.value)
                        wb.save('productos.xlsx')
                else:
                    print(f"No se encontró un proyecto asociado con la clave {fecha_reporte}")
                    return
        except Error as e:
            print (e)
        except:
            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
            
#Opcion 5
def Registrar_Sala ():
    while True:
        SALA = input("Como se va a llamar la sala?(Escribe SALIR para regresar al menú): ")
        if SALA=='':
            break
        elif SALA=='SALIR':
            return
        else:
            capacity = int(input("Cual va a ser la capacidad?: "))
            N = rd.randint(1,99)
            try:
                with sqlite3.connect("34.db") as conn:
                    mi_cursor = conn.cursor()
                    Valores5={"clave":N,"nombre":SALA,"capacidad":capacity }
                    mi_cursor.execute("INSERT INTO Salas (clave, nombre, capacidad) VALUES(:clave,:nombre,:capacidad)",Valores5)
                    print("Sala Registrada!!")
                    print("Tu clave de la Sala es la siguiente")
                    print(N)
            except Error as e:
                print (e)
            except:
                print(f"Surgio una falla siendo esta la causa: {sys.exc_info()[0]}")
            finally:
                if (conn):
                    conn.close()

# Opcion 6
def Registrar_Cliente ():
    while True:
        Usuario=input("Ingresa al usuario.(Escribe SALIR si quieres regresar al menú principal): ")
        n = rd.randint(1,99)
        if Usuario=='':
            break
        elif Usuario=='SALIR':
            return
        else:
            try:
                with sqlite3.connect("34.db") as conn:
                    mi_cursor = conn.cursor()
                    valores={"clave":n,"nombre":Usuario} 
                    mi_cursor.execute("INSERT INTO Usuarios (clave, nombre) VALUES(:clave,:nombre)",valores)
                print("Usuario registrado!")
                print("Tu clave de cliente es la siguiente: ")
                print(n)
            except Error as e:
                print (e)
            except:
                print(f"Surgio una falla siendo esta la causa: {sys.exc_info()[0]}")
            finally:
                if (conn):
                    conn.close()

# Opcion 7
def salir_del_programa ():
    print ("*"*30)
    print ("Hasta pronto tenga un buen dia :D")
    print ("*"*30)
    print ("Vuelve a visitarnos pronto")
    print ("*"*30)
    
# Opcion 8
def eliminar_reservacion ():
    while True:
        key_code = int(input("Dime tu clave de cliente: "))
        try:
            with sqlite3.connect("34.db") as conn:
                mi_cursor = conn.cursor()
                valores = {"clave":key_code}
                mi_cursor.execute("SELECT * FROM Usuarios WHERE clave = :clave", valores)
                registro = mi_cursor.fetchall()

                if registro:
                    for clave, nombre in registro:
                        print(f"{clave}\t{nombre}")
                else:
                    print(f"No se encontró un proyecto asociado con la clave {key_code}")
        except Error as e:
            print (e)
        except:
            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
        else:
            id_sala = input("¿Cual es el id de la sala que quieres eliminar?(Escribe SALIR para regresar al programa): ")
            if id_sala == '':
                continue
            elif id_sala == 'SALIR':
                return
            else:
                Name = input("Cual era tu nombre del evento?: ")
                Turno = input("Escribe tu horario: ")
                fecha_asignada = input("Ingresa la fecha por favor: ")
                fecha_dt = datetime.strptime(fecha_asignada, '%d/%m/%Y')
                fecha_permitida = datetime.now() + timedelta( days = 3)
                if fecha_dt < fecha_permitida:
                    print("Lo siento pero no la puedes cancelar, debes cancelarla con 3 dias de anticipación")
                else:
                    try:
                        with sqlite3.connect("Miami.db") as conn:
                            mi_cursor = conn.cursor()
                            delete = {"folio":id_sala,"nombre":Name,"horario":Turno,"fecha":fecha_dt}
                            mi_cursor.execute("DELETE FROM Reservaciones WHERE folio = :folio;", delete)
                            #delete={"folio":id_sala,"nombre":Name,"horario":Turno,"fecha":fecha_dt}
                            #mi_cursor.execute("DELETE FROM Reservaciones WHERE (fecha) = (:fecha);", delete)
                            print("Reservacion eliminada. ¡Lamentamos que hayas decidido cancelar tu evento!")
                            return
                    except sqlite3.Error as e:
                        print (e)
                    except Exception:
                        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                    finally:
                        if (conn):
                            conn.close()
                            print("Se ha cerrado la conexión")
                        
def exportar_base_de_datos_a_excel ():
    while True:
        wb = openpyxl.Workbook()
        hoja = wb.active
        print(f'Hoja activa: {hoja.title}')
        hoja.append(('Clave', 'Nombre'))
        hoja["B1"] = 86
        hoja["B2"] = "Alan Javier"
        print("Excel creado correctamente")
        break





def Mi_menú ():
    print("Conexion Establecida")
    #EstablecerConexion ()
    Crear_tabla ()
    print("*" * 20)
    print("1. Registra la reservacion\n"
    "2. Modificar las descripciones de la reservacion\n" +
    "3. Consulta la fecha disponible\n" +
    "4. Reporte de la reservaciones de una fecha\n" +
    "5. Registrar Sala\n" +
    "6. Registrar Cliente\n" +
    "7. Salir del programa\n" +
    "8. Eliminar reservacion \n" +
    "9. Exportar base de datos a Excel \n")
    while True:
        try:
            Opcion = int(input("Seleccione el numero de la accion que quiere realizar \n:"))
        except Error as e:
            print(e)
        except:
            print(f"Ocurrió un problema {sys.exc_info()[0]}")
        else:
            if Opcion == 1:
                Registrar_Reservacion ()
            elif Opcion == 2:
                modificar_descripciones ()
            elif Opcion == 3:
                consulta_fecha()
            elif Opcion == 4:
                reporte_a_Excel ()
            elif Opcion == 5:
                Registrar_Sala ()
            elif Opcion == 6:
                Registrar_Cliente ()
            elif Opcion == 7:
                salir_del_programa ()
                break
            elif Opcion == 8:
                eliminar_reservacion ()
            elif Opcion == 9:
                exportar_base_de_datos_a_excel ()
            else:
                print("Eso no esta disponible checa el menú")
        


Mi_menú ()

# Reservacion
# Disponibilidad
# Modificar
# Elimanar
# Reporte
# Reporte en excel
# Sala
#Cliente
# Salir

#fin