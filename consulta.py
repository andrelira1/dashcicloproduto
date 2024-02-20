import streamlit as st
import cx_Oracle
import pandas as pd
import plotly.express as px
import SQLs

# Define as credenciais de conexão com o banco Oracle
#cx_Oracle.init_oracle_client(lib_dir="C:\instantclient_19_10")

user = 'teste'
password = 'teste'
dsn = 'localhost'

try:
  connection = cx_Oracle.connect(user, password, dsn)
##except KeyError as e:
except cx_Oracle.DatabaseError as e:
 st.error("Falha na conexão: " + str(e))

else:
  ###-------CONSULTAS--------###
  # Executa a query SQL e carrega os resultados em um DataFrame pandas
  dfc_rf = pd.read_sql_query(SQLs.query1, connection)

  df_rc = pd.read_sql_query(SQLs.query2, connection)

###------TRATATIVAS--------###
  connection.close()