from selenium import webdriver
from bs4 import BeautifulSoup
from SQLServer import cnx
from selenium.webdriver.chrome.options import Options
import time
import re

# define la sentencia insert con los parámetros necesarios
def makequery(table, semana, fecha_inicio, fecha_termino, tipo_producto, datos):
    query = ("INSERT INTO " + table + " "
             "(semana , fecha_inicio, fecha_termino, tipo_producto, producto, precio_promedio) "
              "VALUES ('"+semana+"','"+fecha_inicio+"','"+fecha_termino+"','"+tipo_producto+"','"+datos+"')")
    return query

tabla = "[dbo].[ODEPA_Precios]"
cursor = cnx.cursor()
# selecciona la semana para guardarla como parámetro de ingreso de datos
cursor.execute("SELECT TOP 1 semana, fecha_inicio, fecha_termino FROM "+tabla+" ORDER BY id DESC")
for (sem, fecha_ini, fecha_fin) in cursor:
    semana = sem
    fecha_inicio = fecha_ini
    fecha_termino = fecha_fin

print("Se ingresaran las bebidas con la siguiente fecha:")
print("Semana: %d" % semana)
print("Fecha Inicio: %s" % fecha_inicio)
print("Fecha Término: %s" % fecha_termino)
input("Presioner ENTER para continuar e ingresar los datos")
chrome_options = Options()
chrome_options.add_argument('--disable-gpu')
driver = webdriver.Chrome(chrome_options=chrome_options)

link = "http://www.jumbo.cl/FO/CategoryDisplay?cab=4012&int=3773&ter=3350"

# obtiene bebidas light y no light
# ingresa a la base de datos
for i in range(2):
    driver.get(link)
    time.sleep(0.1)
    elem = driver.find_element_by_id("tabla_productos")
    html = driver.page_source
    contenido = elem.get_attribute('innerHTML')
    soup = BeautifulSoup(contenido, 'lxml')

    datatags = soup('div')
    i = 0
    litros = 0
    count = 0
    key = 0
    dicc_bebidas = {}
    # formatea los datos obteniendo nombre de la bebida, precio por litro
    for tag in datatags:
        if tag.string is not None:
            try:
                tempo1 = i+1
                if tempo1 % 3 == 0:
                    # solo busca las variables que tengan el precio por litro
                    if 'Litro' in tag.string:
                        dicc_bebidas[key].append(tag.string)
                    else:
                        i += 1
                        if i % 3 == 0:
                            key += 1
                        dicc_bebidas[key].append(tag.string)
                        continue
                else:
                    dicc_bebidas[key].append(tag.string)
            except KeyError:
                dicc_bebidas[key] = []
                dicc_bebidas[key].append(tag.string)
            i += 1
            if i % 3 == 0:
                key += 1
    row_data = ""
    for key in dicc_bebidas:
        try:
            row_data = dicc_bebidas[key][0]+"','"+re.sub("[^0-9]", "", dicc_bebidas[key][2])
        except IndexError:
            continue
        query = makequery(tabla, str(semana), fecha_inicio, fecha_termino, "Bebidas", row_data)
        cursor.execute(query)
        
    link = "http://www.jumbo.cl/FO/CategoryDisplay?cab=4012&int=3773&ter=4814"

cnx.commit()
driver.quit()
print("Cerrando la conexión a la base de datos.")
cursor.close()
cnx.close()
print("Se han ingresado correctamente los datos de Bebidas light y no light")
input("ENTER para finalizar el programa.")

