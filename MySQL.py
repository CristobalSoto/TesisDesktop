import pyodbc

# datos de conexión para SQL Server
server = 'localhost'
database = 'preciosconsumidor'
driver = '{Mariadb ODBC 3.1 Driver}'
cnx = pyodbc.connect(
    'DRIVER=' + driver + ';PORT=3306;SERVER=' + server + ';DATABASE=' + database + ";USER=root;PWD=epepox")

