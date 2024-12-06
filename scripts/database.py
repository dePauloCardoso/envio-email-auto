import pyodbc
import pandas as pd
from datetime import datetime
from config.database_config import DB_SERVER, DB_DATABASE, DB_USERNAME, DB_PASSWORD

# 1. Estabelecer Conexão com o Banco de Dados
def connect_to_db():
    conn_str = (
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={DB_SERVER};'
        f'DATABASE={DB_DATABASE};'
        f'UID={DB_USERNAME};'
        f'PWD={DB_PASSWORD}'
    )
    conn = pyodbc.connect(conn_str)
    return conn

# 2. Executar Consultas SQL
def execute_query(conn, query):
    cursor = conn.cursor()
    cursor.execute(query)
    result = cursor.fetchall()
    return result

# 3. Manipular e Processar Dados
def query_to_dataframe(conn):
    today = datetime.now()

       # Se for segunda-feira, pegar d-3; caso contrário, pegar d-1
    if today.weekday() == 0:  # 0 é segunda-feira
        query = """
        SELECT 
            TRIM(ZZ3_PRODUT) AS [PRODUTO],
            TRIM(ZZ7_PRODUT) AS [SKU AVULSO],
            TRIM(ZZ3_DESCRI) AS [DESCRICAO PRODUTO],
            SUM(ZZ3_QUANT) AS [QUANT_PEDIDO],
            (ZZ7_QUANT) AS [QUANT_ZZ7],
            SUM(IIF(ZZ7_QUANT <> 1, ZZ3_QUANT * ZZ7_QUANT, ZZ3_QUANT)) AS [QUANT_FAT],
            TRIM(FORMAT(CAST(F2_EMISSAO AS DATE), 'dd/MM/yyyy')) AS [DATA_FAT]
        FROM ZZ3030 (NOLOCK)
            INNER JOIN ZZ2030 (NOLOCK) ON 
                ZZ3_FILIAL = ZZ2_FILIAL AND    
                ZZ3_NUMERO = ZZ2_NUMERO AND 
                ZZ3_MEDICA = ZZ2_MEDICA AND 
                ZZ3_CLIENT = ZZ2_CLIENT
            LEFT JOIN SF2030 NF (NOLOCK) ON
                ZZ2_CLIENT = NF.F2_CLIENTE AND
                (ZZ2_PV02NF = NF.F2_DOC OR ZZ2_PV01NF = NF.F2_DOC) AND
                ZZ2_FILIAL = NF.F2_FILIAL
            LEFT JOIN ZZ7030 (NOLOCK) ON 
                ZZ3_PRODUT = ZZ7_CODPAI
        WHERE
            ZZ3030.D_E_L_E_T_ <>'*'
            AND (ZZ7030.D_E_L_E_T_ IS NULL OR ZZ7030.D_E_L_E_T_ = '')
            AND ZZ2_FILIAL IN ('110102')
            AND CAST(F2_EMISSAO AS DATE) = CAST(GETDATE()-3 AS DATE)
            AND ZZ2_TIPO <> '000033'
            AND F2_SERIE = '1'
        GROUP BY ZZ7_PRODUT, ZZ3_PRODUT, ZZ3_DESCRI, F2_EMISSAO, ZZ7_QUANT
        ORDER BY F2_EMISSAO, ZZ3_PRODUT, ZZ7_PRODUT
        """
    else:
        query ="""
        SELECT 
            TRIM(ZZ3_PRODUT) AS [PRODUTO],
            TRIM(ZZ7_PRODUT) AS [SKU AVULSO],
            TRIM(ZZ3_DESCRI) AS [DESCRICAO PRODUTO],
            SUM(ZZ3_QUANT) AS [QUANT_PEDIDO],
            (ZZ7_QUANT) AS [QUANT_ZZ7],
            SUM(IIF(ZZ7_QUANT <> 1, ZZ3_QUANT * ZZ7_QUANT, ZZ3_QUANT)) AS [QUANT_FAT],
            TRIM(FORMAT(CAST(F2_EMISSAO AS DATE), 'dd/MM/yyyy')) AS [DATA_FAT]
        FROM ZZ3030 (NOLOCK)
            INNER JOIN ZZ2030 (NOLOCK) ON 
                ZZ3_FILIAL = ZZ2_FILIAL AND    
                ZZ3_NUMERO = ZZ2_NUMERO AND 
                ZZ3_MEDICA = ZZ2_MEDICA AND 
                ZZ3_CLIENT = ZZ2_CLIENT
            LEFT JOIN SF2030 NF (NOLOCK) ON
                ZZ2_CLIENT = NF.F2_CLIENTE AND
                (ZZ2_PV02NF = NF.F2_DOC OR ZZ2_PV01NF = NF.F2_DOC) AND
                ZZ2_FILIAL = NF.F2_FILIAL
            LEFT JOIN ZZ7030 (NOLOCK) ON 
                ZZ3_PRODUT = ZZ7_CODPAI
        WHERE
            ZZ3030.D_E_L_E_T_ <>'*'
            AND (ZZ7030.D_E_L_E_T_ IS NULL OR ZZ7030.D_E_L_E_T_ = '')
            AND ZZ2_FILIAL IN ('110102')
            AND CAST(F2_EMISSAO AS DATE) = CAST(GETDATE()-1 AS DATE)
            AND ZZ2_TIPO <> '000033'
            AND F2_SERIE = '1'
        GROUP BY ZZ7_PRODUT, ZZ3_PRODUT, ZZ3_DESCRI, F2_EMISSAO, ZZ7_QUANT
        ORDER BY F2_EMISSAO, ZZ3_PRODUT, ZZ7_PRODUT
    """
    df = pd.read_sql_query(query, conn)
    return df

# 4. Fechar a Conexão
def close_connection(conn):
    conn.close()
