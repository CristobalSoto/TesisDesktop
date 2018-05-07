import pyodbc

# datos de conexión para SQL Server
server = 'SECO'
database = 'IRF'
driver = '{SQL Server}'
cnx = pyodbc.connect(
    'DRIVER=' + driver + ';PORT=1433;SERVER=' + server + ';PORT=1443;DATABASE=' + database + ';Trusted_Connection=yes;')

