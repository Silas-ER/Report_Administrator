#############################################################################################################
################################## FUNÇÃO PARA CONEXÃO COM O BANCO DE DADOS #################################
#############################################################################################################

import pyodbc
import pandas as pd

def consultar(mes_atual):
    # conexão com o banco de dados
    connect = pyodbc.connect('Driver={SQL Server};' 'Server=SQL;' 'Database=SQL;' 'UID=USER;' 'PWD=PASSWORD;')
    cursor = connect.cursor()
    
    #consulta e resultado
    consulta = """ 
        SELECT 
            CONVERT(VARCHAR(10), BASE.data_hora, 103) AS DATA,
            CONVERT(VARCHAR(5), BASE.data_hora, 108) AS HORA,
            SUM(CASE WHEN REF.nome_lingua2 IN ('Panulirus argus','palinuridae-panulirus argus','Panulirus spp','palinuridae-panulirus spp') THEN MV.QTDE_UND ELSE 0 END) AS [LAGOSTA VERMELHA],
			'' AS COLCALC1,
            SUM(CASE WHEN REF.nome_lingua2 IN ('Panulirus laevecauda','palinuridae-panulirus laevecauda') THEN MV.QTDE_UND ELSE 0 END) AS [LAGOSTA CABO VERDE],
			'' AS COLCALC2,
            SUM(CASE WHEN REF.nome_lingua2 IN ('Panulirus echinatus') THEN MV.QTDE_UND ELSE 0 END) AS [LAGOSTA DE PEDRA],
			'' AS COLCALC3,
            MV.NUM_DOCTO AS NF,
            SUBSTRING(MV.DESC_PRODUTO_EST, 1, 30) AS OBSERVACAO
            --OBS.OBSERVACAO AS BARCO
            
        FROM 
            vwMovtoEntrada MV
            LEFT JOIN vwAtak4Net_EntradasBase AS BASE ON (BASE.Chave_fato = MV.CHAVE_FATO)
            LEFT JOIN tbProdutoRef AS REF ON (MV.COD_PRODUTO = REF.Cod_produto)
            LEFT JOIN tbEntradasObs OBS ON OBS.Chave_fato = MV.CHAVE_FATO

        WHERE 
            MV.COD_DIVISAO1 = '004'
            AND MV.COD_TIPO_MV = 'T221'
            AND MV.DESC_PRODUTO_EST NOT LIKE '%SAPATA%'
			AND MONTH(DATA_HORA) = {}
            
        GROUP BY 
            CONVERT(VARCHAR(10), BASE.data_hora, 103),
            CONVERT(VARCHAR(5), BASE.data_hora, 108),
            MV.DESC_PRODUTO_EST,
            MV.QTDE_UND,
            MV.NUM_DOCTO,
            OBS.OBSERVACAO
        """.format(mes_atual)

    resultado = pd.read_sql(consulta, connect)        
    return resultado