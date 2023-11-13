import sqlite3
import openpyxl

# Creación de una base de datos SQLite y tablas
def Crear_Base_de_Datos():
    try:
        with sqlite3.connect("walmart.db") as conn:
            mi_cursor = conn.cursor()
            # Crear la tabla Producto
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS Producto (id INTEGER PRIMARY KEY, nombre TEXT NOT NULL, categoria TEXT NOT NULL, precio REAL NOT NULL, cantidad INTEGER NOT NULL);")
            # Crear la tabla Venta
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS Venta (id INTEGER PRIMARY KEY, producto_id INTEGER NOT NULL, cantidad INTEGER NOT NULL, fecha_venta TEXT NOT NULL, sucursal TEXT NOT NULL);")
            # Crear la tabla Sucursal
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS Sucursal (id INTEGER PRIMARY KEY, nombre TEXT NOT NULL, ciudad TEXT NOT NULL, pais TEXT NOT NULL);")
            print("Base de datos y tablas creadas exitosamente")
    except sqlite3.Error as e:
        print(e)

# Función para agregar un producto
def Agregar_producto():
    try:
        nombre = input("Ingrese el nombre del producto: ")
        categoria = input("Ingrese la categoría del producto: ")
        precio = float(input("Ingrese el precio del producto: "))
        cantidad = int(input("Ingrese la cantidad de productos que se registrarán: "))

        with sqlite3.connect("walmart.db") as conn:
            mi_cursor = conn.cursor()
            # Insertar los datos del producto en la tabla Producto
            mi_cursor.execute("INSERT INTO Producto (nombre, categoria, precio, cantidad) VALUES (?, ?, ?, ?)",
                              (nombre, categoria, precio, cantidad))
            conn.commit()
            print(f"{cantidad} producto(s) de {nombre} agregado(s) exitosamente")
    except sqlite3.Error as e:
        print(e)

# Función para agregar una sucursal
def Agregar_sucursal():
    try:
        nombre = input("Ingrese el nombre de la sucursal: ")
        ciudad = input("Ingrese la ciudad de la sucursal: ")
        pais = input("Ingrese el país de la sucursal: ")

        with sqlite3.connect("walmart.db") as conn:
            mi_cursor = conn.cursor()
            mi_cursor.execute("INSERT INTO Sucursal (nombre, ciudad, pais) VALUES (?, ?, ?)",
                              (nombre, ciudad, pais))
            conn.commit()
            print("Sucursal agregada exitosamente")
    except sqlite3.Error as e:
        print(e)

# Función para registrar una venta
def Registrar_venta():
    try:
        nombre_producto = input("Ingrese el nombre del producto vendido: ")
        cantidad_vendida = int(input("Ingrese la cantidad vendida: "))
        fecha_venta = input("Ingrese la fecha de venta (DD-MM-AAAA): ")
        sucursal = input("Ingrese el nombre de la sucursal: ")

        with sqlite3.connect("walmart.db") as conn:
            mi_cursor = conn.cursor()
            mi_cursor.execute("SELECT id, cantidad FROM Producto WHERE nombre=?", (nombre_producto,))
            resultado = mi_cursor.fetchone()

            if resultado:
                producto_id, cantidad_disponible = resultado

                if cantidad_vendida <= cantidad_disponible:
                    mi_cursor.execute("INSERT INTO Venta (producto_id, cantidad, fecha_venta, sucursal) VALUES (?, ?, ?, ?)",
                                      (producto_id, cantidad_vendida, fecha_venta, sucursal))

                    nueva_cantidad = cantidad_disponible - cantidad_vendida
                    mi_cursor.execute("UPDATE Producto SET cantidad=? WHERE id=?", (nueva_cantidad, producto_id))

                    conn.commit()
                    print("Venta registrada exitosamente")
                else:
                    print(f"No hay suficientes unidades de {nombre_producto} en stock.")
            else:
                print(f"No se encontró un producto con el nombre {nombre_producto}. Asegúrese de que el producto exista en la base de datos.")
    except sqlite3.Error as e:
        print(e)

# Función para exportar los datos a un archivo de Excel
def Exportar_a_Excel():
    try:
        with sqlite3.connect("walmart.db") as conn:
            mi_cursor = conn.cursor()
            mi_cursor.execute("SELECT * FROM Producto")
            datos_productos = mi_cursor.fetchall()

            libro_excel = openpyxl.Workbook()
            hoja = libro_excel.active
            hoja.title = "Productos"

            encabezados = ["ID", "Nombre", "Categoría", "Precio", "Cantidad"]
            hoja.append(encabezados)

            for producto in datos_productos:
                hoja.append(producto)

            libro_excel.save("productos.xlsx")
            print("Datos exportados a 'productos.xlsx' exitosamente.")
    except sqlite3.Error as e:
        print(e)

# Función principal que muestra el menú
def Menu_principal():
    Crear_Base_de_Datos()
    while True:
        print("\nMenú - Walmart:")
        print("1. Agregar producto")
        print("2. Agregar sucursal")
        print("3. Registrar venta")
        print("4. Exportar datos a Excel")
        print("5. Salir")
        opcion = input("Elija una opción: ")

        if opcion == "1":
            Agregar_producto()
        elif opcion == "2":
            Agregar_sucursal()
        elif opcion == "3":
            Registrar_venta()
        elif opcion == "4":
            Exportar_a_Excel()
        elif opcion == "5":
            print("Saliendo del programa.")
            break
        else:
            print("Opción no válida. Por favor, seleccione una opción válida.")

if __name__ == "__main__":
    Menu_principal()