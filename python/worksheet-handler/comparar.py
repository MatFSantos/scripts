# Compara duas tabelas excel, e verifica se o email presente na segunda tabela
# também está presente na primeira tabela.
#
import pandas;

print("=======================================================================")
print("========= Bem vindo ao script de comparação de tabelas ================")
print("=======================================================================")

print("Esse script verifica se um email da tabela 2 está presente na tabela 1.")

print("\n*obs:* Para fazer a comparação é preciso que os arquivos xlsx")
print("estejam na pasta 'planilhas' e possuam colunas 'Nome completo' e 'E-mail'.\n")

try:
    tableName1 = None
    while tableName1 is None:
        tableName1 = input("tabela 1 (sem extensão): ")
        try:
            table1 = pandas.read_excel("planilhas/"+tableName1+".xlsx")
        except KeyboardInterrupt:
            print("\n\nExiting...")
            exit(1)
        except:
            print("Um erro aconteceu, tente novamete.\n")
            tableName1 = None
            table1 = None
    
    tableName2 = None
    while tableName2 is None:
        tableName2 = input("tabela 2 (sem extensão): ")
        try:
            table2 = pandas.read_excel("planilhas/"+tableName2+".xlsx")
        except KeyboardInterrupt:
            print("\n\nExiting...")
            exit(1)
        except:
            print("Um erro aconteceu, tente novamete.\n")
            tableName2 = None
            table2 = None
        
except KeyboardInterrupt:
    print("\n\nExiting...")
    exit(1)

table_result = pandas.DataFrame(columns=['Nome completo', 'E-mail', 'Presença'])

i = 0

for data1 in table2['E-mail']:
    contain = 0
    email1 = str(data1).lower()
    for data2 in table1['E-mail']:
        email2 = str(data2).lower()
        if email1 == email2:
            contain = contain + 1

    if contain == 0:
        table_result = table_result.append({
            'Nome completo': table2['Nome completo'][i],
            'E-mail': table2['E-mail'][i],
            'Presença': table2['Presença'][i]
        },ignore_index=True)
    i = i + 1

try:
    table_result.to_excel('resultado.xlsx')
except:
    print("Um erro ocorreu ao gerar a tabela resultado.")

print("Tabela resultado gerada na raiz com sucesso!\n")
print("Procure por resultado.xlsx. . .")
