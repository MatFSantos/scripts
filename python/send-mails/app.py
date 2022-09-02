from dotenv import load_dotenv

load_dotenv()

from conectServer import getServer
from openFile import openExcel
from createEmail import createEmail
from certificateGenerate import certificateGenerate

def app():

    tableName = "apend.xlsx"
    columnName = "Nome completo"
    columnEmail = "E-mail"

    templateName = ""
    action="Ouvinte"
    eventName="Seminário FLIFS virtual"
    theme="Letramento Literário: propostas, práticas e possibilidades"
    time="19 a 21 de julho de 2022"
    workload=20

    i= 0
    table = openExcel(fileName=tableName)
    # table = {"Nome completo": ["Nélia de Medeiros Sampaio"], "E-mail": ["nmsampaio@uefs.br"]}
    server = getServer()

    i = 0
    j = 0
    print("Enviando emails...")
    for aprovedEmail in table[columnEmail]:
        print(f"{i}: email para {aprovedEmail}. . .")
        text = "Certificamos que " + str(table[columnName][i]).upper() + ", participou como " + action.upper() + " do "
        text += eventName.upper() + ", cujo o tema foi " + theme.upper() + ", no período de "
        text += time + ", com carga horária de " + str(workload) + " horas."


        if str(table[columnName][i]) == "Rita de Cássia Brêda Mascarenhas Lima":
            templateName = "assinatura_cristiana.jpg"
        elif str(table[columnName][i]) == "Cristiana Barbosa de Oliveira ramos":
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