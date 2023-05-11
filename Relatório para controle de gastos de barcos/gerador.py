from layouts import layout_inicial, layout_final
from relatorio import relatorio_entradas, relatorio_saidas 

validacao = 'T'

def comunica_laco():
    escolha = input("\nDeseja gerar outro relatorio?[sim/não] ")
    print("\n")
    escolha = escolha.upper()
    if(escolha == 'SIM' or escolha == 'S'):     
        return 'T'
    else:
        return 'F'

layout_inicial()

while(validacao != 'F'):
    seletor = int(input("1. Entradas por serviço\n2. Saídas por barco\nQual relatorio deseja emitir (1 ou 2)? "))

    if(seletor == 1):
        relatorio_entradas()
        validacao = comunica_laco()

    elif(seletor == 2):
        relatorio_saidas()
        validacao = comunica_laco()

    else:
        print("\nSelecione uma opção válida!\n")

layout_final()