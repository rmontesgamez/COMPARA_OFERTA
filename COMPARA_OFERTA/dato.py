import pandas as pd
import openpyxl as op
import shutil
#from PyPDF2 import PdfFileWriter, PdfFileReader
import io
#from reportlab.pdfgen import canvas
#from reportlab.lib.pagesizes import A4
import os



def pie_de_pagina(archivo, extension, num_pedido, cant_p, cliente_p, pza_pedido_p):
    
    nombre_pdf_original = 'P:\\Personal Láser Guadalquivir\\RAFA\PLEGADO\\TEMP\\' + archivo + 'int' + extension # nombre original
    #nombre_nuevo="390911r.pdf" # nombre intermedio para hacer pruebas en la propia carpeta

    #os.rename(nombre_pdf_original, nombre_nuevo)

    nombre_pdf_salida =  'P:\\Personal Láser Guadalquivir\\RAFA\PLEGADO\\PLANOS\\' + str(pza_pedido_p) + extension # nombre salida que coincide con el original
    mensaje = 'PEDIDO '+str(num_pedido) + '  CANT  ' + str(cant_p) + '  ' + str(cliente_p)
    packet = io.BytesIO()

    mi_canvas = canvas.Canvas(packet, pagesize=A4)

    """
        También podemos cambiar las coordenadas del mensaje.
        La posición (0, 0) es la esquina inferior izquierda
        por eso es que nuestro mensaje sale tan cerca a 
        dicho lugar
    """
    mi_canvas.drawString(250, 15, mensaje)
    mi_canvas.save()

    packet.seek(0)
    pdf_con_pie = PdfFileReader(packet)

    pdf_existente = PdfFileReader(open(nombre_pdf_original, "rb"))
    page=pdf_existente.getPage(0)

    output = PdfFileWriter()

    page.mergePage(pdf_con_pie.getPage(0))

    output.addPage(page)
    outputStream = open(nombre_pdf_salida, "wb")
    output.write(outputStream)
    outputStream.close()

    

    #os.remove(nombre_pdf_original)




def existe_archivo(nombre, extension):

    if nombre[:3]=='0000':
        ruta= '\\\\192.168.8.2\\piezas\\00' + nombre[4:5] +'000\\' + nombre[3:] + extension
    
    elif nombre[:2] == '000':
        ruta = '\\\\192.168.8.2\\piezas\\0' + nombre[3:5] +'000\\' + nombre[3:] + extension
   
    else:
        ruta = '\\\\192.168.8.2\\piezas\\' + nombre[2:5] +'000\\' + nombre[2:] + extension

    print (ruta)

    ruta_destino= 'P:\\Personal Láser Guadalquivir\\RAFA\PLEGADO\\TEMP\\' + nombre + 'int' + extension
    if not os.path.exists(ruta_destino):

        if os.path.exists(ruta):
        
        
            try:
                shutil.copy(ruta, ruta_destino)

                return True
            except:
                print('ARCHIVO NO MOVIDO:', nombre)

        else: 
            return False

def borrador_archivos():
    directorios=['P:\\Personal Láser Guadalquivir\\RAFA\PLEGADO\\TEMP', 'P:\\Personal Láser Guadalquivir\\RAFA\PLEGADO\\PLANOS']

    for directorio_a_borrar in directorios:
        with os.scandir(directorio_a_borrar) as entries:
            for entry in entries:
                if entry.is_file() or entry.is_symlink():
                    os.remove(entry.path)
                #elif entry.is_dir():
                    #shutil.rmtree(entry.path)