import pyodbc
import pandas as pd

class consulta(object):

    """description of class"""


    def __init__(self):
        self.ruta='DSN=FirebirdDSN'
        self.cnxn = pyodbc.connect(self.ruta)
       

    def sql_query(self, tipo, parametros):
        self.tipo = tipo
        self.parametros = {}
        self.fecha_inicial=parametros.get('fecha_inicial')
        self.fecha_final=parametros.get('fecha_final')
        self.nume_cliente = parametros.get('n_cliente')
        self.e_pedido = parametros.get('estado_pedido')

        if tipo=='ofertas':
            
            sql_query=(""" 
            SELECT CLIENTES.NOM, OFERTAS.CANT, OFERTAS.ESTADO, OFERTAS.OFERTA, CLIENTES.TEL, OFERTAS.FECHA
FROM CLIENTES INNER JOIN OFERTAS ON CLIENTES.CODIGO = OFERTAS.CODCLIENTE
WHERE (((OFERTAS.CANT)>1) AND ((OFERTAS.FECHA)='""" + self.fecha_final + "')) OR (((OFERTAS.FECHA)>'" +self.fecha_final +" 0:0:1' And (OFERTAS.FECHA)<'" + self.fecha_final+ " 23:59:59')) ORDER BY OFERTAS.OFERTA;"
            )
            return sql_query

        elif tipo=='oferta_semana':
            sql_query=("""
            SELECT CLIENTES.NOM, OFERTAS.CANT, OFERTAS.ESTADO, OFERTAS.OFERTA, CLIENTES.TEL, OFERTAS.FECHA
FROM CLIENTES INNER JOIN OFERTAS ON CLIENTES.CODIGO = OFERTAS.CODCLIENTE
WHERE (((OFERTAS.CANT)>1) AND ((OFERTAS.FECHA)>='"""+ self.fecha_inicial + "' And (OFERTAS.FECHA)<='" + self.fecha_final+ "')) OR ((((OFERTAS.FECHA)>'" + self.fecha_inicial + " 0:0:1' And (OFERTAS.FECHA)<'" + self.fecha_final + " 23:59:59'))) ORDER BY OFERTAS.OFERTA;"
                )
            return sql_query

        elif tipo=='ofertas_detalle':
            sql_query=("""
            SELECT CLIENTES.NOM, OFERTAS.OFERTA, OFERTAS.CANT, PZOFERTA.VPU, PZOFERTA.QPZ, PZOFERTA.REF, PZOFERTA.TIPOM, PZOFERTA.TRATMTO
FROM CLIENTES INNER JOIN (OFERTAS INNER JOIN PZOFERTA ON OFERTAS.OFERTA = PZOFERTA.OFERTA) ON CLIENTES.CODIGO = OFERTAS.CODCLIENTE
WHERE (((OFERTAS.FECHA)='""" + self.fecha_final + "')) OR (((OFERTAS.FECHA)>'" +self.fecha_final +" 0:0:1' And (OFERTAS.FECHA)<'" + self.fecha_final+
" 23:59:59')) ORDER BY OFERTAS.OFERTA;"
            )
            return sql_query

        elif tipo=='pedidos_simplificados':
            sql_query=("""
           SELECT CLIENTES.NOM, PEDIDOS.PEDIDO, PEDIDOS.REF
FROM CLIENTES INNER JOIN PEDIDOS ON CLIENTES.CODIGO = PEDIDOS.CODCLIENTE
WHERE (((PEDIDOS.F_PED)='""" + self.fecha_inicial + "')) ORDER BY PEDIDOS.PEDIDO;"
                )
            return sql_query

        elif tipo=='pedidos':
            sql_query=("""
        SELECT CLIENTES.NOM, PEDIDOS.PEDIDO, PIEZAS.PU, PIEZAS.REF_C, PIEZAS.TIPOM, PIEZAS.TRATMTO, PIEZAS.LARGO, PIEZAS.ANCHO, PEDIDOS.REF
       FROM PIEZAS INNER JOIN (CLIENTES INNER JOIN (PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) ON CLIENTES.CODIGO = PEDIDOS.CODCLIENTE) ON PIEZAS.REF_N = PZPEDIDO.REF_N
WHERE (((PEDIDOS.F_PED)='""" + self.fecha_inicial + "')) ORDER BY PEDIDOS.PEDIDO;"
                )
            return sql_query


        elif tipo=='clientes':
            sql_query=("""
        SELECT CLIENTES.NOM, CLIENTES.CODIGO, TABLADIR.DIR, CONTACTOS_CLIENTE.NOMBRE, CLIENTES.TEL
FROM (CLIENTES INNER JOIN CONTACTOS_CLIENTE ON CLIENTES.CODIGO = CONTACTOS_CLIENTE.CODCLIENTE) INNER JOIN TABLADIR ON CLIENTES.CODIGO = TABLADIR.CODCLIENTE
WHERE (((CLIENTES.CODIGO)>'02000'))
ORDER BY CLIENTES.CODIGO;"""
                )
            return sql_query
        elif tipo=='ofertas_agrupadas_periodo':
            sql_query=("""
         SELECT CLIENTES.CODIGO, CLIENTES.NOM, Sum(OFERTAS.CANT) AS SumaDeCANT
FROM CLIENTES INNER JOIN OFERTAS ON CLIENTES.CODIGO = OFERTAS.CODCLIENTE
WHERE (((OFERTAS.FECHA)>='""" + self.fecha_inicial +"' And (OFERTAS.FECHA)<='" + self.fecha_final + "')) OR (((OFERTAS.FECHA)>'" + self.fecha_inicial + " 0:0:1' And (OFERTAS.FECHA)<'" + self.fecha_final + """ 23:59:59')) 
GROUP BY CLIENTES.CODIGO, CLIENTES.NOM
ORDER BY Sum(OFERTAS.CANT) DESC;"""

                )
            return sql_query

        elif tipo=='pedidos_periodo_simplificados':
            sql_query=("""
         SELECT PEDIDOS.CODCLIENTE, CLIENTES.NOM, Sum(PEDIDOS.CANT) AS SumaDeCANT
FROM CLIENTES INNER JOIN PEDIDOS ON CLIENTES.CODIGO = PEDIDOS.CODCLIENTE
WHERE (((PEDIDOS.F_PED)>='""" + self.fecha_inicial +"' And (PEDIDOS.F_PED)<='" + self.fecha_final + "')) OR (((PEDIDOS.F_PED)>'" + self.fecha_inicial + " 0:0:1' And (PEDIDOS.F_PED)<'" + self.fecha_final + """ 23:59:59')) 
GROUP BY PEDIDOS.CODCLIENTE, CLIENTES.NOM
ORDER BY Sum(PEDIDOS.CANT) DESC;"""

                )
            return sql_query


        elif tipo=='piezas_plegadas':
            sql_query=(

"""
SELECT PZPEDIDO.CODPZPEDIDO, PIEZAS.REF_N, CLIENTES.NOM, PEDIDOS.PEDIDO, PZASPROCESOSASIGN.O_PED, PIEZAS.TIPOM, PLEGADO_DATOS.MATRIZ, PIEZAS.LARGO, PIEZAS.ANCHO, PZASPROCESOSASIGN.CANT
FROM CLIENTES INNER JOIN ((PIEZAS INNER JOIN ((PROCESOSASIGN INNER JOIN (PEDIDOS INNER JOIN PZASPROCESOSASIGN ON PEDIDOS.PEDIDO = PZASPROCESOSASIGN.PEDIDO) ON PROCESOSASIGN.CODIGO = PZASPROCESOSASIGN.CODIGO) INNER JOIN PZPEDIDO ON (PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) AND (PZASPROCESOSASIGN.O_PED = PZPEDIDO.O_PED)) ON PIEZAS.REF_N = PZPEDIDO.REF_N) INNER JOIN PLEGADO_DATOS ON (PIEZAS.REF_N = PLEGADO_DATOS.REF_N) AND (PZPEDIDO.REF_N = PLEGADO_DATOS.REF_N)) ON (CLIENTES.CODIGO = PIEZAS.CODCLIENTE) AND (CLIENTES.CODIGO = PEDIDOS.CODCLIENTE)
WHERE (((PEDIDOS.PEDIDO)>'""" + '131000' + "') AND ((PZASPROCESOSASIGN.CANT)>(PZASPROCESOSASIGN.CANTREALZDAS)) AND ((PROCESOSASIGN.ESTADO)='" + '0' +"""') AND ((PROCESOSASIGN.GRUPO_CALDERERIA)='""" + '7' + """') )
ORDER BY PEDIDOS.PEDIDO, PLEGADO_DATOS.MATRIZ;"""


                )
            return sql_query
                      
        elif tipo=='piezas_programadas':
            sql_query=("""
            SELECT PZPEDIDO.CODPZPEDIDO, PZPEDIDO.REF_N, PZPEDIDO.PEDIDO
FROM TRABAJOS INNER JOIN ((PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) INNER JOIN PZASTRABAJO ON (PEDIDOS.PEDIDO = PZASTRABAJO.PEDIDO) AND (PZPEDIDO.O_PED = PZASTRABAJO.O_PED)) ON TRABAJOS.CODIGO = PZASTRABAJO.CODIGO
WHERE (((PZPEDIDO.TRATMTO) Like '%09%'))
GROUP BY PZPEDIDO.CODPZPEDIDO, PZPEDIDO.REF_N, PZPEDIDO.PEDIDO, PEDIDOS.COMPLETADO, TRABAJOS.ESTADO
HAVING (((PEDIDOS.COMPLETADO)='N') AND ((Sum(TRABAJOS.MAQUINA))<>'""" + '0' + "') AND ((TRABAJOS.ESTADO)='" + '0' + "'));"

                )
            return sql_query

        elif tipo=='piezas_cortadas':
           
            sql_query=("""

            SELECT PZPEDIDO.CODPZPEDIDO, PZPEDIDO.REF_N, PEDIDOS.PEDIDO
FROM TRABAJOS INNER JOIN ((PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) INNER JOIN PZASTRABAJO ON (PZPEDIDO.O_PED = PZASTRABAJO.O_PED) AND (PEDIDOS.PEDIDO = PZASTRABAJO.PEDIDO)) ON TRABAJOS.CODIGO = PZASTRABAJO.CODIGO
GROUP BY PZPEDIDO.CODPZPEDIDO, PZPEDIDO.REF_N, PEDIDOS.PEDIDO, PEDIDOS.COMPLETADO, PZPEDIDO.TRATMTO, TRABAJOS.ESTADO
HAVING (((PEDIDOS.COMPLETADO)='N') AND ((PZPEDIDO.TRATMTO) Like '%09%') AND ((TRABAJOS.ESTADO)<>'""" + '0' + "'));"

                )
            return sql_query
        
        elif tipo=='plegado_sin_matriz':           

            sql_query=("""
        SELECT PZPEDIDO.CODPZPEDIDO, PIEZAS.REF_N, CLIENTES.NOM, PEDIDOS.PEDIDO, PZASPROCESOSASIGN.O_PED, PIEZAS.TIPOM, PIEZAS.LARGO, PIEZAS.ANCHO, PZASPROCESOSASIGN.CANT
FROM PROCESOSASIGN INNER JOIN (CLIENTES INNER JOIN (PIEZAS INNER JOIN ((PEDIDOS INNER JOIN PZASPROCESOSASIGN ON PEDIDOS.PEDIDO = PZASPROCESOSASIGN.PEDIDO) INNER JOIN PZPEDIDO ON (PZASPROCESOSASIGN.O_PED = PZPEDIDO.O_PED) AND (PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO)) ON PIEZAS.REF_N = PZPEDIDO.REF_N) ON (CLIENTES.CODIGO = PIEZAS.CODCLIENTE) AND (CLIENTES.CODIGO = PEDIDOS.CODCLIENTE)) ON PROCESOSASIGN.CODIGO = PZASPROCESOSASIGN.CODIGO
WHERE (((PEDIDOS.PEDIDO)>'""" + '131000' + "') AND ((PZASPROCESOSASIGN.CANT)>(PZASPROCESOSASIGN.CANTREALZDAS)) AND ((PROCESOSASIGN.ESTADO)='" + '0' + "') AND ((PROCESOSASIGN.GRUPO_CALDERERIA)='" + '7' + "')) ORDER BY PEDIDOS.PEDIDO;"
           
                )
            return sql_query

        elif tipo=='plegado_concreto':           

            sql_query=("""

            SELECT PZPEDIDO.CODPZPEDIDO, PIEZAS.REF_N, CLIENTES.NOM, PEDIDOS.PEDIDO, PZASPROCESOSASIGN.O_PED, PIEZAS.TIPOM, PIEZAS.LARGO, PIEZAS.ANCHO, PZASPROCESOSASIGN.CANT, PROCESOSASIGN.ESTADO, PROCESOSASIGN.GRUPO_CALDERERIA
FROM CLIENTES INNER JOIN (PIEZAS INNER JOIN ((PROCESOSASIGN INNER JOIN (PEDIDOS INNER JOIN PZASPROCESOSASIGN ON PEDIDOS.PEDIDO = PZASPROCESOSASIGN.PEDIDO) ON PROCESOSASIGN.CODIGO = PZASPROCESOSASIGN.CODIGO) INNER JOIN PZPEDIDO ON (PZASPROCESOSASIGN.O_PED = PZPEDIDO.O_PED) AND (PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO)) ON PIEZAS.REF_N = PZPEDIDO.REF_N) ON (CLIENTES.CODIGO = PEDIDOS.CODCLIENTE) AND (CLIENTES.CODIGO = PIEZAS.CODCLIENTE)
WHERE (((PEDIDOS.PEDIDO)='""" + '133248' + "'));"


                )
            return sql_query

        elif tipo=='pedidos_n_cerrados':           

            sql_query=("""
            
            
            SELECT PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PIEZAS.TIPOM, PZPEDIDO.TRATMTO, PZPEDIDO.QP, PZASPDTES.PDTESP, PZASPDTES.PDTESC
FROM (PIEZAS INNER JOIN (PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) ON PIEZAS.REF_N = PZPEDIDO.REF_N) INNER JOIN PZASPDTES ON (PEDIDOS.PEDIDO = PZASPDTES.PEDIDO) AND (PIEZAS.REF_N = PZASPDTES.REF_N)
GROUP BY PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PIEZAS.TIPOM, PZPEDIDO.TRATMTO, PZPEDIDO.QP, PZASPDTES.PDTESP, PZASPDTES.PDTESC, PIEZAS.REF_N, PEDIDOS.COMPLETADO, PZASPDTES.PEDIDO, PZASPDTES.REF_N
HAVING (((PEDIDOS.CODCLIENTE)='""" + self.nume_cliente + "') AND ((PEDIDOS.COMPLETADO)='N') AND ((PZASPDTES.PEDIDO)=(PEDIDOS.PEDIDO)) AND ((PZASPDTES.REF_N)=(PIEZAS.REF_N)));"

            

                )
            return sql_query

        elif tipo=='pedidos_plegado':           

            sql_query=("""

            SELECT PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PZASPROCESOSASIGN.CANT, PZASPROCESOSASIGN.CANTREALZDAS, PZPEDIDO.TRATMTO
FROM (PIEZAS INNER JOIN (PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) ON PIEZAS.REF_N = PZPEDIDO.REF_N) INNER JOIN PZASPROCESOSASIGN ON PEDIDOS.PEDIDO = PZASPROCESOSASIGN.PEDIDO
WHERE (((PZASPROCESOSASIGN.PEDIDO)=((PEDIDOS.PEDIDO))) AND ((PZASPROCESOSASIGN.O_PED)=((PZPEDIDO.O_PED))))
GROUP BY PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PZASPROCESOSASIGN.CANT, PZASPROCESOSASIGN.CANTREALZDAS, PZPEDIDO.TRATMTO, PEDIDOS.COMPLETADO
HAVING (((PEDIDOS.CODCLIENTE)='""" + self.nume_cliente + """') AND ((PEDIDOS.COMPLETADO)='N'))
ORDER BY PEDIDOS.PEDIDO;"""

                 )

            return sql_query


        elif tipo == 'estado_pedido':           

            sql_query=("""
            
            SELECT PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PIEZAS.TIPOM, PZPEDIDO.TRATMTO, PZPEDIDO.QP, PZASPDTES.PDTESP, PZASPDTES.PDTESC
FROM (PIEZAS INNER JOIN (PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) ON PIEZAS.REF_N = PZPEDIDO.REF_N) INNER JOIN PZASPDTES ON (PIEZAS.REF_N = PZASPDTES.REF_N) AND (PEDIDOS.PEDIDO = PZASPDTES.PEDIDO)
GROUP BY PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PIEZAS.TIPOM, PZPEDIDO.TRATMTO, PZPEDIDO.QP, PZASPDTES.PDTESP, PZASPDTES.PDTESC, PIEZAS.REF_N, PEDIDOS.COMPLETADO, PZASPDTES.PEDIDO, PZASPDTES.REF_N
HAVING (((PEDIDOS.REF) Like '%""" + self.e_pedido + "%') AND ((PEDIDOS.COMPLETADO)='N') AND ((PZASPDTES.PEDIDO)=(PEDIDOS.PEDIDO)) AND ((PZASPDTES.REF_N)=(PIEZAS.REF_N)));"
            
                 )

            return sql_query
        
        elif tipo == 'est_pedidos_pleg':           

            sql_query=("""

            SELECT PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PZASPROCESOSASIGN.CANT, PZASPROCESOSASIGN.CANTREALZDAS, PZPEDIDO.TRATMTO
FROM (PIEZAS INNER JOIN (PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) ON PIEZAS.REF_N = PZPEDIDO.REF_N) INNER JOIN PZASPROCESOSASIGN ON PEDIDOS.PEDIDO = PZASPROCESOSASIGN.PEDIDO
WHERE (((PZASPROCESOSASIGN.PEDIDO)=((PEDIDOS.PEDIDO))) AND ((PZASPROCESOSASIGN.O_PED)=((PZPEDIDO.O_PED))))
GROUP BY PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PZASPROCESOSASIGN.CANT, PZASPROCESOSASIGN.CANTREALZDAS, PZPEDIDO.TRATMTO, PEDIDOS.COMPLETADO
HAVING (((PEDIDOS.REF) Like '%""" + self.e_pedido+ "%') AND ((PEDIDOS.COMPLETADO)='N')) ORDER BY PEDIDOS.PEDIDO;"

                )

            return sql_query
        
        elif tipo == 'estado_pieza':           

            sql_query=("""
        SELECT PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PIEZAS.TIPOM, PZPEDIDO.TRATMTO, PZPEDIDO.QP, PZASPDTES.PDTESP, PZASPDTES.PDTESC 
        FROM (PIEZAS INNER JOIN (PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) ON PIEZAS.REF_N = PZPEDIDO.REF_N) INNER JOIN PZASPDTES ON (PIEZAS.REF_N = PZASPDTES.REF_N) AND (PEDIDOS.PEDIDO = PZASPDTES.PEDIDO)
GROUP BY PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PIEZAS.TIPOM, PZPEDIDO.TRATMTO, PZPEDIDO.QP, PZASPDTES.PDTESP, PZASPDTES.PDTESC, PIEZAS.REF_N, PEDIDOS.COMPLETADO, PZASPDTES.PEDIDO, PZASPDTES.REF_N 
HAVING (((PIEZAS.REF_C) Like '%""" + self.e_pedido + "%') AND ((PEDIDOS.COMPLETADO)='N') AND ((PZASPDTES.PEDIDO)=(PEDIDOS.PEDIDO)) AND ((PZASPDTES.REF_N)=(PIEZAS.REF_N))) ORDER BY PEDIDOS.PEDIDO;"

  )

            return sql_query

        elif tipo == 'est_piezas_pleg':           

            sql_query=("""
            SELECT PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PZASPROCESOSASIGN.CANT, PZASPROCESOSASIGN.CANTREALZDAS, PZPEDIDO.TRATMTO
FROM (PIEZAS INNER JOIN (PEDIDOS INNER JOIN PZPEDIDO ON PEDIDOS.PEDIDO = PZPEDIDO.PEDIDO) ON PIEZAS.REF_N = PZPEDIDO.REF_N) INNER JOIN PZASPROCESOSASIGN ON PEDIDOS.PEDIDO = PZASPROCESOSASIGN.PEDIDO 
WHERE (((PZASPROCESOSASIGN.PEDIDO)=((PEDIDOS.PEDIDO))) AND ((PZASPROCESOSASIGN.O_PED)=((PZPEDIDO.O_PED))))
GROUP BY PEDIDOS.CODCLIENTE, PEDIDOS.PEDIDO, PZPEDIDO.CODPZPEDIDO, PZPEDIDO.O_PED, PEDIDOS.REF, PIEZAS.REF_C, PZASPROCESOSASIGN.CANT, PZASPROCESOSASIGN.CANTREALZDAS, PZPEDIDO.TRATMTO, PEDIDOS.COMPLETADO 
HAVING (((PIEZAS.REF_C) Like '%""" + self.e_pedido + "%') AND ((PEDIDOS.COMPLETADO)='N')) ORDER BY PEDIDOS.PEDIDO;"

                )

            return sql_query

        elif tipo == 'max_pedido':   
            sql_query=("""
            SELECT Max((PEDIDOS.PEDIDO))  FROM PEDIDOS;"""

                 )

            return sql_query

        elif tipo == 'max_pieza':   
            sql_query=("""
           SELECT Max((PIEZAS.REF_N))  FROM PIEZAS;"""
                )
            return sql_query


        elif tipo == 'forma_oferta':   
            sql_query=("""
            SELECT OFERTAS.OFERTA, PZOFERTA.REF, PIEZAS.REF_N, PIEZAS.C_VR, PZOFERTA.VPU, PZOFERTA.QPZ, PZOFERTA.PROPMAT, PZOFERTA.PR_TRANSPORTE, PIEZAS.VCORTE, PZOFERTA.VTRATMTO
FROM CLIENTES INNER JOIN (PIEZAS INNER JOIN (OFERTAS INNER JOIN PZOFERTA ON OFERTAS.OFERTA = PZOFERTA.OFERTA) ON (PZOFERTA.C_VR = PIEZAS.C_VR) AND (PIEZAS.REF_C = PZOFERTA.REF)) ON (CLIENTES.CODIGO = PIEZAS.CODCLIENTE) AND (CLIENTES.CODIGO = OFERTAS.CODCLIENTE)
WHERE (((OFERTAS.OFERTA)='""" + str(self.e_pedido) + "'));"

                 )
            return sql_query

        elif tipo == 'max_oferta':   
            sql_query=("""
            SELECT MAX(OFERTAS.OFERTA) FROM OFERTAS;"""
                )

            return sql_query


    def fetch_query(self,frase):
        c = self.cnxn.cursor()
        try:
            c.execute(frase)
            results = c.fetchall()
        except:
            results=[()]
            print("CONSULTA ERRONEA")
        self.cnxn.close()
        return results

    def exec_query(self, frase):
        try:
            c = self.cnxn.cursor()
            c.execute(frase)
            self.cnxn.commit()
        except:
            print("ACCION ERRONEA")

        self.cnxn.close()

    def close(self):
        self.cnxn.close()
        
    def consulta_pandas(self, frase):
     
        df=pd.read_sql_query (frase, self.cnxn)

        self.cnxn.close()
        return df
      

        