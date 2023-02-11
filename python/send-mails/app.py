from dotenv import load_dotenv

load_dotenv()

from conectServer import getServer
from createEmail import createEmail
from certificateGenerate import certificateGenerate

def app(
    table: str,
    templateName: str,
    columnEmail: str,
    text: str,
    signature: str,
    fields
):
    print(fields)

    i= 0
    server = getServer()

    i = 0
    j = 0
    print("Enviando emails...")
    for aprovedEmail in table[columnEmail]:
        print(f"{i}: email para {aprovedEmail}. . .")
        args = {}
        for field in fields:
            args[field] = table[field][i]
            
        newText = text.format(**args)
        pdfName = certificateGenerate(
            fileName=f"cert_{i}.pdf" ,
            templateName=templateName,
            signature=signature,
            text= newText,
        )

        email = createEmail(
            subject="Certificado do III Simpósio de Biotecnologia",
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