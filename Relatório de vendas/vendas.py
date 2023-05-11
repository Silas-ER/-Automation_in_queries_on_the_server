import pyodbc 
import pandas as pd
from datetime import date

#conexão com o servidor
conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=SERVER;'                      
                        'Database=DATA;'
                        'UID=USER;'
			            'PWD=PASSWORD;'                     
                        'Trusted_Connection=yes;')
cursor = conn.cursor()

#ajuste de datas e condição para virada de ano
data_de_hoje = date.today()
ano_atual = data_de_hoje.year
mes_atual = 2

#ajuste caso vire o ano para emitir do mês 12 do ano anterior            
if(mes_atual == 1):
    mes_atual = 13
    ano_atual = ano_atual - 1

#recebendo o mês anterior para a formulação da pesquisa
mes_anterior = mes_atual - 1

#criação do arquivo em excel da pesquisa solicitada
df = pd.read_sql(r"select CHAVE_FATO, NUM_DOCTO, COD_TIPO_MV, DATA_MOVTO, COD_CLI_FOR, NOME_CLI_FOR, COD_PRODUTO, QTDE_UND, COD_UNIDADE, VALOR_UNITARIO, VALOR_TOTAL_ITEM, VALOR_TOTAL, DESC_PRODUTO_EST, NOME_SECAO, NOME_GRUPO, NOME_SUBGRUPO from vwAtak4Net_Saidas where Month(DATA_MOVTO) = '{}' and Year(DATA_MOVTO) = '{}' ".format(mes_anterior,ano_atual),conn)
if(mes_anterior > 9):
    df.to_excel(r"\\servidor\Exportação2\CONTROLES\DADOS ESTATISTICO-MAPA\VENDAS_{}_MES_{}.xlsx".format(ano_atual,mes_anterior))
else:
    df.to_excel(r"\\servidor\Exportação2\CONTROLES\DADOS ESTATISTICO-MAPA\VENDAS_{}_MES_0{}.xlsx".format(ano_atual,mes_anterior))
