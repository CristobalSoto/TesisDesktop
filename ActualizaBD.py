# Importación de librerías
#______________________________________________________________________________________________________________________________

#!/usr/bin/python
# -*- coding: 850 -*-
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import *
from bs4 import BeautifulSoup
import time
# importa la conexión para que esta no sea visible en el código del programa
# from SQLServer import cnx
from MySQL import cnx
import urllib.request
import sys

# Inicialización de variables 
#______________________________________________________________________________________________________________________________

anoSem = 2
codRegion = 0
codSector = 0
codTipoProducto = 0
firstloop = True
sector = ""
#tabla = "[dbo].[ODEPA_Precios]"
tabla = "odepa_precios"
link = "http://apps.odepa.cl/menu/PreciosConsumidorNS.action?consumidor=frutas&reporte=precios_consumidor"

# Define métodos para ingresar los datos
def makequery(table, tags, datos, region, producto, sectorp):
    if sectorp == "":
        cueri = ("INSERT INTO " + table +
                 " (region,tipo_producto," + tags + ") "
                 "VALUES ('" + region + "','" + producto + "'," + datos + ")")
    else:
        cueri = ("INSERT INTO " + table +
                 " (region, sector,tipo_producto," + tags + ") "
                 "VALUES ('" + region + "','" + sectorp + "','" + producto + "'," + datos + ")")
    return cueri

# Inicia el cursor que interactua con la base de datos
time_start = time.time()


cursor = cnx.cursor()

print("Consultando la última fecha en la base de datos.")

# Consulta el último valor en la tabla
# sql server
# cursor.execute("SELECT TOP 1 fecha_termino, semana FROM "+tabla+" ORDER BY id DESC")

cursor.execute("SELECT fecha_termino, semana FROM " + tabla + " ORDER BY id desc LIMIT 1")


rs = cursor.fetchall()
for (fecha_termino, semana) in rs:
    last_weekBD = semana
    last_yearBD = int(fecha_termino[6:])

# Si el resultado de la consulta anterior es "vacío", setea semana = 1 y año = 2008 por default.
#______________________________________________________________________________________________________________________
if not rs:
    last_weekBD = 1
    last_yearBD = 2008
    print("La tabla se encuentra sin datos, se considerara el primer año y la primera semana como fecha de inicio.")


print("Consultando la última fecha en la página de la ODEPA.")

# Eventualidad 1: Comprueba si la página esta funcionando
#______________________________________________________________________________________________________________________________
try:
    html = urllib.request.urlopen(link)
except:
    print("La página no se pudo cargar correctamente.")
    print("Finalizando el programa.")
    sys.exit()

# Soup es la librería que permite obtener el contenido de la página de manera accesible
#______________________________________________________________________________________________________________________________
soup = BeautifulSoup(html, "html.parser")

# Obtiene el año desde la página de la ODEPA
#______________________________________________________________________________________________________________________________
last_weekODEPA = [str(x.text) for x in soup.find(id="semanaIni").find_all('option', selected=True)]
last_yearODEPA = soup.find(id="anoSem").text
# Lo pasa a entero
#______________________________________________________________________________________________________________________________
last_weekODEPA = int(last_weekODEPA[0])
last_yearODEPA = int(last_yearODEPA[1:5])

# Saca los datos  (regiones, tipos producto)
#______________________________________________________________________________________________________________________________
regiones = [str(x.text) for x in soup.find(id="codRegion").find_all('option')]
tipos_producto = [str(x.text) for x in soup.find(id="codTipoProducto").find_all('option')]

# carga el motor de navegador y link
# options es para que no muestre el mensaje de error al principio
#______________________________________________________________________________________________________________________________
chrome_options = Options()
chrome_options.add_argument('--disable-gpu')
driver = webdriver.Chrome(chrome_options=chrome_options)
driver.get(link)

print("Comparando ambas fechas y determinando si se encuentra actualizada la base de datos. \n")
# compara la última fecha de la base de datos con la fecha la ODEPA 
# Si el año ODEPA > max año BD => Se selecciona los campos de año y semana en la página web
# Si no, => termina ejecución si la fecha esta actualizada
# cierra el navegador al terminar la comparación
#______________________________________________________________________________________________________________________________
if last_yearODEPA > last_yearBD:
    select = Select(driver.find_element_by_name('params.anoSem'))
    select.select_by_value(str(last_yearBD))
    select_semanafin = driver.find_element_by_name("params.semanaFin")
    options = [x for x in select_semanafin.find_elements_by_tag_name("option")]
    semana_fin = len(options)
    semana_inicio = last_weekBD
    anoSem = last_yearBD
    if last_weekBD == semana_fin:
        last_yearBD += 1
        anoSem += 1
        semana_inicio = 1
else:
    anoSem = last_yearODEPA
    if last_weekODEPA > last_weekBD:
        semana_inicio = last_weekBD + 1
        semana_fin = last_weekODEPA
    else:
        driver.close()
        print("Los datos coinciden con el último registro de la página.")
        print("Los datos se encuentran actualizados.")
        print("El programa ha finalizado.")
        sys.exit()
driver.close()

print("Se han encontrado nuevos datos para ingresar. \n")
print("Fecha de inicio de datos a actualizar.")
print("Año: %d" % anoSem)
print("Semana inicio: %d \n" % semana_inicio)
input("Presione ENTER para confirmar y continuar la ejecución del programa.")
# Carga el navegador 
#______________________________________________________________________________________________________________________________
driver = webdriver.Chrome()
print("Ingresando datos de:")
# En este bloque se obtienen los datos desde la fecha declarada dependiendo del resultado de la comparación
#______________________________________________________________________________________________________________________________
for i in range((last_yearODEPA-last_yearBD)+1):
    for codRegion in range(len(regiones)):
        print()
        print("Region: " + regiones[codRegion])
        for codTipoProducto in range(len(tipos_producto)):
            print("Producto: "+tipos_producto[codTipoProducto])
            try:
                driver.get(link)
                # Define los campos del formulario para la consulta de los datos en la página
                #______________________________________________________________________________________________________________
                select = Select(driver.find_element_by_name('params.anoSem'))
                select.select_by_value(str(anoSem))
                select = Select(driver.find_element_by_name('params.semanaIni'))
                select.select_by_index(semana_inicio-1)
                select = Select(driver.find_element_by_name('params.semanaFin'))
                # Bandera para ver la cantidad de semanas del año(52 o 53)
                #______________________________________________________________________________________________________________
                if firstloop:
                    firstloop = False
                    select.select_by_index(semana_fin-1)
                else:
                    select_semanafin = driver.find_element_by_name("params.semanaFin")
                    options = [x for x in select_semanafin.find_elements_by_tag_name("option")]
                    select.select_by_index(len(options) - 1)

                select = Select(driver.find_element_by_name('params.codRegion'))
                select.select_by_index(codRegion)
                select = Select(driver.find_element_by_name('params.codSector'))
                select.select_by_index(codSector)
                select_sectores = driver.find_element_by_name("params.codSector")
                # Se determina si existe 1 solo sector
                #______________________________________________________________________________________________________________
                sectores = [x for x in select_sectores.find_elements_by_tag_name("option")]

                if len(sectores) == 1:
                    for obj in sectores:
                        sector = obj.text
                else:
                    sector = ""

                select = Select(driver.find_element_by_name('params.codTipoProducto'))
                select.select_by_index(codTipoProducto)

                select = Select(driver.find_element_by_name('params.codProducto'))
                select.select_by_index(0)

                select = Select(driver.find_element_by_name('params.codTipoPuntoMonitoreo'))

                # Obtiene la cantidad de puntos de monitoreo y los selecciona todos
                #_______________________________________________________________________________________________________________
                select_monitoreo = driver.find_element_by_name("params.codTipoPuntoMonitoreo")
                options = [x for x in select_monitoreo.find_elements_by_tag_name("option")]

                for element in range(len(options)):
                    select.select_by_index(element)
            # Excepciones a los errores cuando no se encuentra los datos y cuando la página carga de manera no esperada
            #___________________________________________________________________________________________________________________
            except (StaleElementReferenceException, NoSuchElementException):
                driver.close()
                print("La página ha cargado de manera incorrecta.")
                print("Se recomienda cerrar la ventana y ejecutar el programa nuevamente.")
                print("Se terminará la ejecución del programa.")
                input("Presione ENTER para finalizar.")
                sys.exit()
            # enviar formulario
            #___________________________________________________________________________________________________________________
            driver.find_element_by_css_selector('.boton2').click()

# Se empieza con la extracción de los datos de la tabla resultado del envío del formulario
#___________________________________________________________________________________________________________________
            try:
                elem = driver.find_element_by_tag_name('table')
            # Excepci{on para el caso en que no encuentra la tabla sigue con la siguiente opción (cuando la página no arroja ningún resultado)
            except NoSuchElementException:
                continue

            # obtiene el código html
            #___________________________________________________________________________________________________________________
            contenido = elem.get_attribute('innerHTML')
            # remueve etiqueta molesta <br>
            contenido = contenido.replace("<br>", " ")
            # utiliza beautifulSoup para poder acceder al contenido de las etiquetas HTML
            soup = BeautifulSoup(contenido, "html.parser")

            # busca los tags y los formatea para obtener los datos
            #___________________________________________________________________________________________________________________
            headertags = soup('th')
            count = 0
            columns = ""
            # columnas obtenidas como th (table headers)
            for tag in headertags:
                # remplaza los espacios en blanco (" ") por guión bajo ("_") para ingresarlo como columna en la base de datos
                columns = columns + tag.string.replace(' ', '_').lower() + ","
                count += 1
            columns = columns[:-1]
            i = 0
            # obtiene los td (table data) que llevan el contenido de las filas de la tabla
            datatags = soup('td')
            row = "'"
            # forma la fila formateando y omitiendo los datos innecesarios que arroja el documento
            # la metodología utilizada es concatenar los datos de la fila separados por comas para ingresarlos utilizando
            # la función creada al principio del código y pasarle los datos como párametros de la función
            #___________________________________________________________________________________________________________________
            for tag in datatags:
                if tag.string is None:
                    i += 1
                    row += "','"
                    continue
                # elimina la ultima fila 
                if "Elaborado" in tag.string:
                    continue
                # elimina elemento innecesario
                if "\'" in tag.string:
                    fila = tag.string.replace('.', '')
                    row = row + str(fila.replace("'", "''")) + "','"
                else:
                    row = row + str(tag.string.replace('.', '')) + "','"
                i += 1
                # inserta la fila si I llega a la cantidad de columnas de la tabla
                if i == count:
                    row = row[:-2]
                    query = makequery(tabla, columns, row, regiones[codRegion], tipos_producto[codTipoProducto], sector)
                    # sentencia insert sin commit
                    cursor.execute(query)
                    row = "'"
                    i = 0
    anoSem += 1
    semana_inicio = 1
# Se finaliza la extracción de datos; se cierra la conexión y el navegador; se hace commit
#______________________________________________________________________________________________________________________________
driver.quit()
print("Ingresando los datos a la tabla.")
# confirma los cambios a la base de datos
cnx.commit()
# excepción para los productos que solo tengan 1 clasificación
cursor.execute("UPDATE "+tabla+" "
               "SET producto = "+tabla+".tipo_producto "
               "WHERE producto Is Null")
cnx.commit()
print("Cerrando la conexión a la base de datos.")
cursor.close()
cnx.close()
duracion = time.time() - time_start
print("El programa tardó %d segundos." % int(duracion))
print("Se han ingresado los datos satisfactoriamente.")
input("Presione ENTER para finalizar el programa.")

