# PARA INSTALAR MODULO PARA USAR EXCEL
# pip install pandas
# pip install pyarrow
# pip install xlsxwriter

# PARA INSTALAR MODULO PARA USAR MARIADB
# pip install sqlalchemy
# pip install mysqlclient
# pip install mysql
# pip install mysql-connector-python

# PARA INSTALAR MODULO QUE GENERA . EXE
# pip install auto-py-to-exe

# Importamos la libreria para trabajar con SQL
import pandas as pd

# Importamos el modulo para trabajar con fechas y horas
import datetime

# Generamos un bucle infinito
while True:

    # Consultamos el valor de la hora y fecha
    TimePC = datetime.datetime.now()

    # Configuramos el nombre del archivo
    name = str(TimePC.day) + str(TimePC.month) + str(TimePC.year)

    # Almacenamos la hora
    Hour = TimePC.hour

    # Almacenamos los minutos
    Minute =TimePC.minute
    
    # Si la hora coincide...
    # LINEA ORIGINAL
    # if (Hour == 23 and Minute == 59):
    if (Hour == 23 and Minute == 59) or True:
        
        # Importamos la libreria para realizar una conexion SQL
        from sqlalchemy import create_engine

        # Configuracion de la conexion a la base de datos
        engine_db = create_engine('mysql+mysqlconnector://root:admin@127.0.0.1:3306/basededatos')

        # Definimos el nombre de la tabla  1 que vamos a exportar
        table1 = 'tabla1'
        # Definimos el nombre de la tabla  2 que vamos a exportar
        table2 = 'tabla2'
        # Definimos el nombre de la tabla  3 que vamos a exportar
        table3 = 'tabla3'

        # Realizamos la consulta de la tabla 1
        query1 = f'SELECT * FROM {table1}'
        # Realizamos la consulta de la tabla 2
        query2 = f'SELECT * FROM {table2}'
        # Realizamos la consulta de la tabla 3
        query3 = f'SELECT * FROM {table3}'

        # Se ejecuta la consulta y se almacenan los datos en una variable
        df1 = pd.read_sql(query1, engine_db)
        # Se ejecuta la consulta y se almacenan los datos en una variable
        df2 = pd.read_sql(query2, engine_db)
        # Se ejecuta la consulta y se almacenan los datos en una variable
        df3 = pd.read_sql(query3, engine_db)

        # Cerramos la conexion con la base de datos
        # connection_db.close()

        # Guardamos los datos en un archivo CSV
        # df.to_csv('Tabla{}.csv'.format(name), index=False)

        # Guardamos los datos en un archivo CSV
        with pd.ExcelWriter('Data{}.xlsx'.format(name), engine='xlsxwriter') as writer:
            # Escribimos la consulta 1 en el excel
            df1.to_excel(writer, sheet_name='Hoja1', index=False)
            # Escribimos la consulta 1 en el excel
            df2.to_excel(writer, sheet_name='Hoja2', index=False)
            # Escribimos la consulta 1 en el excel
            df3.to_excel(writer, sheet_name='Hoja3', index=False)

        # ELIMINA LA LINEA DE ABAJO PARA EVITAR QUE SE CIERRE EL PROGRAMA
        break