import os
from dotenv import load_dotenv

load_dotenv()

from conectServer import getServer
from openFile import openExcel
from createEmail import createEmail
from certificateGenerate import certificateGenerate


def app():

    table = {"E-mail": ["ferrmathh2@gmail.com"], "Nome completo": ["Mateus Fernandes Santos"]}

    action="monitor"
    eventName="Seminário FLIFS virtual"
    theme="Letramento Literário: propostas, práticas e possibilidades"
    time="19 a 21 de julho de 2022"
    workload=20


    # table = openExcel(fileName="aprovados.xlsx")
    server = getServer()

    i = 0
    for aprovedEmail in table["E-mail"]:

        text = "Certificamos que " + str(table["Nome completo"][i]).upper() + ", participou como " + action.upper() + " do "
        text += eventName.upper() + ", cujo o tema foi " + theme.upper() + ", no período de "
        text += time.upper() + ", com carga horária de " + str(workload) + " horas."


        if str(table["Nome completo"][i]).__contains__("Rita"):
            templateName = "assinatura_cristiana.jpg"
        elif str(table["Nome completo"][i]).__contains__("Cristiana"):
            templateName = "assinatura_rita.jpg"
        else:
            templateName = "duas_assinaturas.jpg"
        pdfName = certificateGenerate(
            name= table["Nome completo"][i],
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
    server.quit()


app()