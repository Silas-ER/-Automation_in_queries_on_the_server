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

df = pd.read_sql(r"SELECT CHAVE_FATO, COD_FILIAL, NUM_DOCTO, COD_CLI_FOR, COD_TIPO_MV, DATA_MOVTO, STATUS, VALOR_TOTAL, VALOR_LIQUIDO, OBSERVACAO FROM tbEntradas WHERE COD_TIPO_MV = 'F105' AND COD_CLI_FOR = '90260' AND MONTH(DATA_MOVTO) = {} AND YEAR(DATA_MOVTO) = {}".format(mes_anterior,ano_atual), conn)
df.to_excel(r"\\servidor\Almoxarifado\Relatórios para conferência de NFs\HEVILE - F105\HEVILES_F105_{}_{}.xlsx".format(mes_anterior,ano_atual))