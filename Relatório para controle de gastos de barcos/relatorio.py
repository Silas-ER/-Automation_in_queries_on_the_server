from datetime import date, datetime
import pandas as pd
import pyodbc 

conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=SQL;'                      
                        'Database=;'
                        'UID=;'
			            'PWD=;'
                        'Trusted_Connection=;')
cursor = conn.cursor()

def relatorio_saidas():
    str_data_inicial = input("\nInsira a data inicial no formato dd/mm/aaaa: ")
    str_data_final =  input("\nInsira a data final no formato dd/mm/aaaa: ")

    print("\n")

    data_inicial = datetime.strptime(str_data_inicial, '%d/%m/%Y').date()
    data_final = datetime.strptime(str_data_final, '%d/%m/%Y').date()
    
    df = pd.read_sql(r"select Data_movto, Num_docto, Cod_produto, Desc_produto_est, Qtde_und, Valor_unitario, Valor_total, Nome_ccusto from vwSaidasGCompras2 where Cod_tipo_mv in ('1151', '1152', '1153', '1154') and Data_movto >= '{}' and Data_movto < '{}'".format(data_inicial,data_final),conn)
                   
    df.to_excel(r"\\servidor\Compras\RELATORIOS\RELATORIO_MATERIAIS_{}.xlsx".format(data_final))

def relatorio_entradas():
    str_data_inicial = input("\nInsira a data inicial no formato dd/mm/aaaa: ")
    str_data_final =  input("\nInsira a data final no formato dd/mm/aaaa: ")

    data_inicial = datetime.strptime(str_data_inicial, '%d/%m/%Y').date()
    data_final = datetime.strptime(str_data_final, '%d/%m/%Y').date()
 
    df = pd.read_sql(r"select Data_lancto, Num_docto_aux, Cod_cli_for, Nome_cadastro, Desc_Rateio, valor_dc, Nome_ccusto from vwEntradasG where Cod_tipo_mv = 'T139' or Cod_tipo_mv = 'F105' and Data_lancto >= '{}' and Data_lancto < '{}'".format(data_inicial,data_final),conn)
                   
    df.to_excel(r"\\servidor\Compras\RELATORIOS\RELATORIO_SERVIÇOS_{}.xlsx".format(data_final))
