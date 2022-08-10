import os
from dotenv import load_dotenv

load_dotenv()

from conectServer import getServer
from openFile import openExcel
from createEmail import createEmail
from certificateGenerate import certificateGenerate


def app():
    table = openExcel(fileName="aprovados.xlsx")
    server = getServer()

    i = 0
    for aprovedEmail in table["E-mail"]:
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
            text= "",
        )