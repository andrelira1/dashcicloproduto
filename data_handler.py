#import pymysql
#import calendar
import pandas as pd
from consulta import *
from datetime import datetime

def atualiza():
    user = 'teste'
    password = 'teste'
    dsn = 'localhost'

    try:
        connection = cx_Oracle.connect(user, password, dsn)
    except cx_Oracle.DatabaseError as e:
        st.error("Falha na conexão: " + str(e))

    else:
    ###-------CONSULTAS--------###

        dfc_rf = pd.read_sql_query(SQLs.query1, connection)
        df_rc = pd.read_sql_query(SQLs.query2, connection)

    connection.close()

    ultima_atualizacao = datetime.now().strftime("%d/%m/%Y %H:%M")

    return dfc_rf, df_rc, ultima_atualizacao
