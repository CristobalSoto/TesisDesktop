import pyodbc

# datos de conexión para SQL Server
server = 'localhost'
database = 'preciosconsumidor'
driver = '{MySQL ODBC 3.51 Driver}'
cnx = pyodbc.connect(
    'DRIVER=' + driver + ';PORT=3306;SERVER=' + server + ';DATABASE=' + database + ";USER=root;PWD=epepox")

