# Scrpit usado para verificar a presença de emails em listas de presença com
# arquivos xlsx.
# 

from numpy import e
import pandas;

try:
    print("=======================================================================")
    print("========= Bem vindo ao script de verificação de presença ==============")
    print("=======================================================================")

    print("Fornceça alguns dados a seguir para continuar com a verificação de presença:\n")

    print("\n*obs:* Para fazer a verificação de presença é preciso que os arquivos xlsx")
    print("estejam na pasta 'planilhas' e possuam colunas de email e de nome.\n")

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

    file_result = pandas.DataFrame(columns=['Nome completo', 'E-mail', 'Presença'])

    i = 0

    for lista in range(0, number_file):

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
        columnEmail = None
        while columnEmail is None:
            columnEmail = input()
            try:
                file_presence[columnEmail]
            except KeyError:
                print('Coluna de e-mail não encontrada no arquivo informado (' + file_path + '). . .')
                print("tente novamente: ")
                columnEmail = None
            except KeyboardInterrupt:
                print("Script abortado. . .")
                exit(1)
            except:
                print("Um erro inesperado aconteceu. . .")
                print("tente novamente: ")
                columnEmail = None
        
        # Captura o nome da coluna de nome e verifica se tá certo.
        print('Nome da coluna de e-mail no arquivo: ')
        columnName = None
        while columnName is None:
            columnName = input()
            try:
                file_presence[columnName]
            except KeyError:
                print('Coluna de nome não encontrada no arquivo informado (' + file_path + '). . .')
                print("tente novamente: ")
                columnName = None
            except KeyboardInterrupt:
                print("Script abortado. . .")
                exit(1)
            except:
                print("Um erro inesperado aconteceu...")
                print("tente novamente: ")
                columnName = None

        i = 0
        list = []

        for email in  file_presence[columnEmail]:
            email = str(email).replace(" ", "").lower()
            notIn = True
            for data in file_result["E-mail"]:
                data = str(data).replace(" ", "").lower()
                if data == email:
                    notIn = False
                    break
            if notIn:
                file_result = pandas.concat(
                    [
                        file_result,
                        pandas.DataFrame(
                            data={
                                "Nome completo": [file_presence[columnName][i]],
                                "E-mail": [email],
                                "Presença": [1]
                            }
                        )
                    ],
                    ignore_index=True
                )
            else:
                if not email in list:
                    index = file_result.index[file_result['E-mail'] == email].tolist()[0]
                    value = int(file_result['Presença'][index]) + 1
                    file_result['Presença'][index] = value
                    list.append(email)
            i = i + 1
except KeyboardInterrupt:
    print("\n\nInterrupted")
    exit(1)

file_finally = pandas.DataFrame(columns=['Nome completo', 'E-mail', 'Presença'])
nao_aprovados = pandas.DataFrame(columns=['Nome completo', 'E-mail', 'Presença'])

i = 0

for sub in file_result['Presença']:
    if int(sub) >= number_presence:
        file_finally = pandas.concat([
                file_finally,
                pandas.DataFrame(
                    data={
                        "Nome completo": [file_result['Nome completo'][i]],
                        'E-mail': [file_result['E-mail'][i]],
                        'Presença': [file_result['Presença'][i]],
                    }
                )
            ],
            ignore_index=True
        )
    else:
        nao_aprovados = pandas.concat([
                nao_aprovados,
                pandas.DataFrame(
                    data={
                        "Nome completo": [file_result['Nome completo'][i]],
                        'E-mail': [file_result['E-mail'][i]],
                        'Presença': [file_result['Presença'][i]],
                    }
                )
            ],
            ignore_index=True
        )
    i = i + 1


try:
    nao_aprovados.to_excel('nao_aprovados.xlsx')
    file_finally.to_excel('aprovados.xlsx')
except:
    print('\n\nErro ao gerar os arquivos. . .')
    print("Script abortado. . .")
    exit(1)

print('\n\nSUCESSO!\n\narquivos gerados na raiz. . .')
print("\nProcure por 'nao_aprovados.xlsx' e 'aprovados.xlsx'")