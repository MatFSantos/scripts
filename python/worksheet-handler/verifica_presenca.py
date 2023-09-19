# Scrpit usado para verificar a presença de emails em listas de presença com
# arquivos xlsx com base em uma lista de inscritos.
# 

from numpy import e
import pandas;

print("=======================================================================")
print("========= Bem vindo ao script de verificação de presença ==============")
print("=======================================================================")
print("Esse script faz a verificação a partir de uma lista de inscritos.\n")
print("Forneça alguns dados a seguir para continuar com a verificação de presença.\n")

print("\n*obs:* Para fazer a verificação de presença é preciso que os arquivos xlsx")
print("estejam na pasta 'planilhas' e possuam colunas de email e de nome.\n")

try:
    file_path = None
    print("Nome do arquivo de inscritos ( omita a extensão ): ")
    while file_path is None:
        file_path = input()
        try:
            file = pandas.read_excel("planilhas/"+file_path+".xlsx")
            print('Sucesso! Arquivo "' + file_path + '" encontrado\n')
        except:
            print('erro ao abrir arquivo "' + file_path + '". . .')
            print('tente novamente: ')
            file_path = None
            file = None
        
    print('Informe o nome das colunas de e-mail e nome do inscrito. . .')

    filter = None
    while filter is None:
        filter = input('Coluna e-mail: ')
        try:
            file[filter]
        except KeyError:
            print('Coluna de e-mail não encontrada no arquivo informado (' + file_path + '). . .')
            print("tente novamente")
            filter = None
        except:
            print("erro inesperado")
            exit(1)

    name = None
    while name is None:
        name = input('coluna nome: ')
        try:
            file[name]
        except KeyError:
            print('Coluna de nome não encontrada no arquivo informado (' + file_path + '). . .')
            print("tente novamente")
            name = None
        except:
            print("erro inesperado")
            exit(1)
    
    number_file = None
    while number_file is None:
        try:
            number_file = int(input('Quantos arquivos de presença serão usados: '))
        except ValueError:
            print("valor fornecido inválido. Forneça um número para continuar.")
            number_file = None
        except KeyboardInterrupt:
            print("Script abortado. . .")
            exit(1)
        except:
            print("Algo deu errado, tente novamente:")
            number_file = None

    number_presence = None
    while number_presence is None:
        try:
            number_presence = int(input('Quantas presenças o participante deve ter: '))
        except ValueError:
            print("valor fornecido inválido. Forneça um número para continuar.")
            number_presence = None
        except KeyboardInterrupt:
            print("Script abortado. . .")
            exit(1)
        except:
            print("Algo deu errado, tente novamente:")
            number_presence = None
except KeyboardInterrupt:
    print('\nExiting...\n')
    exit(1)

file_result = pandas.DataFrame(columns=['Nome completo', 'E-mail', 'Presença'])

i = 0

for data in file[filter]:
    email = str(file[filter][i]).replace(" ", "").lower()
    file_result = pandas.concat([
        file_result,
        pandas.DataFrame(data={
            "Nome completo": [file[name][i]],
            "E-mail": [email],
            "Presença": [0]
        })
    ], ignore_index=True)
    i += 1


for lista in range(0, number_file):
    try:
        # Captura o nome do arquivo e verifica se ta certo.
        file_path = None
        print('Nome do arquivo de presença ( obs: apenas arquivos .xlsx [ omita a extensão ] ): ')
        while file_path is None:
            file_path = input()
            try:
                file_presence = pandas.read_excel("planilhas/"+file_path+".xlsx")
            except KeyboardInterrupt:
                print("Script abortado. . .")
                exit(1)
            except:
                print('Não foi possível abrir o arquivo.')
                print("tente novamente: ")
                file_path = None
                file_presence = None
        
        # Captura o nome da coluna de e-mail e verifica se tá certo.
        print('Nome da coluna de e-mail no arquivo: ')
        column = None
        while column is None:
            column = input()
            try:
                file_presence[column]
            except KeyError:
                print('Coluna de e-mail não encontrada no arquivo informado (' + file_path + '). . .')
                print("tente novamente: ")
                column = None
            except KeyboardInterrupt:
                print("Script abortado. . .")
                exit(1)
            except:
                print("Um erro inesperado aconteceu. . .")
                print("tente novamente: ")
                column = None
    except KeyboardInterrupt:
        print("Script abortado...")
        exit(1)
    i = 0
    list = []
    for data1 in file_result['E-mail']:
        for data2 in file_presence[column]:
            email = str(data2).replace(' ', '').lower()
            if data1 == email:
                if not data1 in list:
                    value = int(file_result['Presença'][i]) + 1
                    file_result['Presença'][i] = value
                    list.append(data1)

        i = i + 1


file_finally = pandas.DataFrame(columns=['Nome completo', 'E-mail', 'Presença'])
nao_aprovados = pandas.DataFrame(columns=['Nome completo', 'E-mail', 'Presença'])

i = 0

for sub in file_result['Presença']:
    if int(sub) >= number_presence:
        file_finally = pandas.concat([
            file_finally,
            pandas.DataFrame(data={
                'Nome completo': [file_result['Nome completo'][i]],
                'E-mail': [file_result['E-mail'][i]],
                'Presença': [file_result['Presença'][i]],
            }),
        ], ignore_index=True)
    else:
        nao_aprovados = pandas.concat([
            nao_aprovados,
            pandas.DataFrame(data={
                'Nome completo': [file_result['Nome completo'][i]],
                'E-mail': [file_result['E-mail'][i]],
                'Presença': [file_result['Presença'][i]],
            })
        ], ignore_index=True)
    i = i + 1


try:
    nao_aprovados.to_excel('nao_aprovados.xlsx')
    file_finally.to_excel('aprovados.xlsx')
except:
    print('\n\nErro ao gerar os arquivos. . .')
    exit(1)

print('\n\nSUCESSO!\n\narquivos gerados na raiz. . .')
print("\nProcure por 'nao_aprovados.xlsx' e 'aprovados.xlsx'")