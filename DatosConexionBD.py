from mysql.connector import (connection)

cnx = connection.MySQLConnection(user='developer', password='epepox',
                                 host='127.0.0.1',
                                 database='preciosconsumidor')
