import pyodbc
import openpyxl as op
import datetime
from consulta import consulta
import volcador
import pandas as pd
import dato
import os
from openpyxl import Workbook

from openpyxl.styles import Color, PatternFill, Font, Border
#from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.formatting import Rule


def datos_piezas(tipo, numero_cliente, estado_pedido):
    fecha=datetime.datetime.today()
    #fecha -= datetime.timedelta(days=365)
    fecha_str = fecha.strftime('%m/%d/%Y')

    parametros_consulta1={'fecha_inicial':fecha_str, 'fecha_final': fecha_str, 'n_cliente': numero_cliente, 'estado_pedido': estado_pedido}
    cr=consulta()
    frase=cr.sql_query(tipo, parametros_consulta1)

    datos=cr.consulta_pandas(frase)
   
    return datos 

fecha=datetime.datetime.today()
fecha_str = fecha.strftime('%m/%d/%Y')

parametros_consulta1={}
# ruta donde se crea el archivo
nombre_archivo="COMPARA_OFERTA.xlsx"   


# borrado de archivo anterior
ruta2="C:\\activa\\"+ nombre_archivo
if os.path.exists(ruta2):
    os.remove(ruta2)

# print("INTRODUCIR NÚMERO OFERTA")

while true:
    try:
        oferta = int(input("INTRODUCIR NÚMERO OFERTA"))
    except ValueError:
        print("OJO, DEBE SER UN NÚMERO")
        continue
    
