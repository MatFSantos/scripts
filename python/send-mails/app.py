from dotenv import load_dotenv

load_dotenv()

from conectServer import getServer
from openFile import openExcel
from createEmail import createEmail
from certificateGenerate import certificateGenerate

def app():

    tableName = "coordenação.xlsx"
    columnName = "Nome"
    columnEmail = "E-mail"
    columnCondition = "Condição"
    columnJob = "Título do Trabalho"
    columnPeriod = "Período "
    columnDate = "Data"
    columnWorkload = "CARGA HORÁRIA"

    templateName = ""

    i= 0
    table = openExcel(fileName=tableName)
    server = getServer()

    i = 0
    j = 0
    print("Enviando emails...")
    for aprovedEmail in table[columnEmail]:
        print(f"{i}: email para {aprovedEmail}. . .")
        text = "Certificamos que " + str(table[columnName][i]).upper() + " " + str(table[columnCondition][i]) + " " + str(table[columnJob][i])
        text += ", no período de  " + str(table[columnDate][i]) + "."


        if str(aprovedEmail) == "rbreda@uefs.br":
            templateName = "assinatura_cristiana.jpg"
        elif str(aprovedEmail) == "cristiana@uefs.br":
            templateName = "assinatura_rita.jpg"
        else:
            templateName = "duas_assinaturas.jpg"
        pdfName = certificateGenerate(
            name= table[columnName][i],
            count= i,
            templateName=templateName,
            text= text,
        )

        email = createEmail(
            subject="Certificado do Seminário FLIFS 2022",
            toEmail=aprovedEmail,
            file=pdfName,
        )

        server.sendmail(email["From"],email["To"], email.as_string())
        i += 1
        j += 1
        print(f"email para {aprovedEmail} enviado com sucesso!\n")
        if j >= 49:
            print("desconectando o servidor")
            server.quit()
            server = getServer()
            j = 0
    server.quit()


app()