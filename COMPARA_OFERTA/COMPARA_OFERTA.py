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

# comprobamos el número oferta

while true:
    try:
        oferta = int(input("INTRODUCIR NÚMERO OFERTA"))
    except ValueError:
        print("OJO, DEBE SER UN NÚMERO")
        continue
    
    if oferta > datos_piezas('max_oferta'):
        print("NÚMERO OFERTA INCORRECTO")
        continue
    else:
        break

cliente = ''
comparitvo =  datos_piezas('forma_pedido', cliente , oferta)

if len(compartivo):
    comparativo['SUMA'] = comparativo.apply(lambda fila: (fila['PROPMAT'] + fila['VCORTE'] + fila['VTRATMTO']), axis = 1)
    comparitvo['SUBTOTAL_NO_VAC'] = comparitvo.apply(lambda fila: (fila['QPZ'] * fila['SUMA']), axis = 1)
    comparitvo['SUBTOTAL_VAC'] = comparitvo.apply(lambda fila: (fila['QPZ'] * fila['VPU']), axis = 1)
    comparitvo['DIFERENCIA'] = comparitvo.apply(lambda fila: (fila['SUBTOTAL_NO_VAC'] - fila['SUBTOTAL_VAC']), axis = 1)
    volcador.volcado_con_pandas(ordenado,'OFERTA',ruta2, "b" )

    wb = op.load_workbook(ruta2)
    ws = wb.active


    longitud = len(ws['A'])
    print(longitud)

    """lista_columna = ['H', 'I', 'J']
    for letra in lista_columna:
        campo = letra + '2:' + letra + str(longitud)

        redFill = PatternFill(start_color='EE1111',
                       end_color='EE1111',
                       fill_type='solid')
        greenfill = PatternFill(start_color='00F000',
                       end_color='00F000',
                       fill_type='solid' )
        ws.conditional_formatting.add(campo,
            FormulaRule(formula=[(letra +'2>0')], fill=redFill))
        ws.conditional_formatting.add(campo,
            FormulaRule(formula=[(letra + '2<=0')], fill=greenfill))"""

    wb.save(ruta2)
    #volcador.correo(ruta2,'RESP. AUTOMATICA', 'RESUMEN.xlsx', destino)