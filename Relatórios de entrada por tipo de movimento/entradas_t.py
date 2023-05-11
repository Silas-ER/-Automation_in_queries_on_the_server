import pandas as pd
import pyodbc
from datetime import date

conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=;'                      
                        'Database=;'
                        'UID=;'
			            'PWD=;'                     
                        'Trusted_Connection=yes;')
cursor = conn.cursor()

#ajuste de datas para automação
data_hoje = date.today()
mes_anterior = data_hoje.month - 1
ano_atual = data_hoje.year

if(data_hoje.month == 1):
    mes_anterior = 12
    ano_atual -= 1

df = pd.read_sql(r"SELECT DISTINCT DATA_MOVTO, NUM_DOCTO, NOME_CADASTRO, VALOR_TOTAL, VALOR_LIQUIDO FROM vwMovtoEntrada WHERE COD_TIPO_MV IN ('T120','T121','T142') AND MONTH(DATA_MOVTO) = {} AND YEAR(DATA_MOVTO) = {}".format(mes_anterior,ano_atual), conn)
df.to_excel(r"\\servidor\Almoxarifado\Relatórios para conferência de NFs\T120, T121, T142\ENTRADAS_NOTAS_T120_{}_{}.xlsx".format(mes_anterior,ano_atual))

dt = pd.read_sql(r"SELECT DISTINCT DATA_MOVTO, NUM_DOCTO, NOME_CADASTRO, VALOR_TOTAL, VALOR_LIQUIDO FROM vwMovtoEntrada WHERE COD_TIPO_MV IN ('T139') AND MONTH(DATA_MOVTO) = {} AND YEAR(DATA_MOVTO) = {}".format(mes_anterior,ano_atual), conn)
dt.to_excel(r"\\servidor\Almoxarifado\Relatórios para conferência de NFs\T139\ENTRADAS_NOTAS_T130_{}_{}.xlsx".format(mes_anterior,ano_atual))

dq = pd.read_sql(r"SELECT DISTINCT DATA_MOVTO, NUM_DOCTO, NOME_CADASTRO, VALOR_TOTAL, VALOR_LIQUIDO FROM vwMovtoEntrada WHERE COD_TIPO_MV IN ('T146') AND MONTH(DATA_MOVTO) = {} AND YEAR(DATA_MOVTO) = {}".format(mes_anterior,ano_atual), conn)
dq.to_excel(r"\\servidor\Almoxarifado\Relatórios para conferência de NFs\T146\ENTRADAS_NOTAS_T146_{}_{}.xlsx".format(mes_anterior,ano_atual))