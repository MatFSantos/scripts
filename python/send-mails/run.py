from app import app
from openFile import openExcel

def run():
    print("[use ctrl + c para cancelar a operação]")
    print("Nome do arquivo (sem a extensão .xlsx):")
    fileName = None
    while fileName is None:
        fileName = input()
        try:
            sheet = openExcel(fileName=fileName+".xlsx")
        except KeyboardInterrupt:
            print("\n\nExiting...")
            exit(1)
        except Exception as e:
            print("Erro: ", e)
            print("Um erro aconteceu, tente novamete.\n")
            fileName = None
    print("\n\n\n\n\nNome do template (com extensão):")
    templateName = None
    while templateName is None:
        templateName = input()
        try:
            open("./templates/" + templateName, "r").close()
        except KeyboardInterrupt:
            print("\n\nExiting...")
            exit(1)
        except Exception as e:
            print("Erro: ", e)
            print("Um erro aconteceu, tente novamete.\n")
            templateName = None
    # Captura o nome da coluna de e-mail e verifica se tá certo.
    print('\n\n\n\n\nNome da coluna de e-mail no arquivo: ')
    columnEmail = None
    while columnEmail is None:
        columnEmail = input()
        try:
            sheet[columnEmail]
        except KeyError:
            print('Coluna de e-mail não encontrada no arquivo informado (' + fileName + '). . .')
            print("tente novamente: ")
            columnEmail = None
        except KeyboardInterrupt:
            print("Script abortado. . .")
            exit(1)
        except:
            print("Um erro inesperado aconteceu. . .")
            print("tente novamente: ")
            columnEmail = None
    print("\n\n\n\n\nInforme o texto do email: ")
    text = None
    while text is None:
        text = input()
        print("\n\n\nVocê digitou:\n" + text)
        choise = ""
        while choise.lower() != "s" and choise.lower() != "n":
            choise = input("Deseja fazer alguma correção?\n[s-S] - sim  [n-N] - não\n")
            if choise.lower() == "s":
                print("Correção: ")
                text = None
    fields = []
    choiseFields = None
    print("\n\n\n\n\nForneça nomes dos campos que serão usados")
    while choiseFields is None:
        fields.append(input())
        print(fields)
        choise = ""
        while choise.lower() != "s" and choise.lower() != "n" and choise != None:
            choise = input("Deseja fornecer mais?\n[s-S] - sim  [n-N] - não\n")
            if choise.lower() == "n":
                choiseFields = ""
            elif choise.lower() == "s":
                choiseFields = None
    print(fields)

    app(sheet, templateName, columnEmail, text,"Feira de Santana, 10 de Feveireiro de 2023" , fields)
run()