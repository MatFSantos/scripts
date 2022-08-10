import os

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

"""
    método que cria uma email com 'subject' e 'body' e envia a partir do servidor e conta
    informados nas variáveis de ambiente.

    Pode ser passado um body em forma de string, nesse caso, é necessário informá-lo
    na chamada do método, passando:
        body="" e subtype="text"
    
    Caso seja usado um body em html o caminho deste deve ser informado nas variáveis de ambiente
    e deve ser omitido o body na chamada do método.

    @param subject: Título do email
    @param name: Nome do destinatário
    @param toEmail: email destinatário
    @param body: corpo do email (opcional)

    @return: um objeto MIMEMultipart contendo as configurações passadas.
"""
def createEmail(
    subject: str,
    toEmail: str,
    body: str | None = None,
    subtype: str ="html",
    file: str | None = None
) -> MIMEMultipart: 

    # monta o email
    emailMsg = MIMEMultipart()
    emailMsg['Subject'] = subject
    emailMsg['To'] = toEmail
    emailMsg['From'] = os.environ["EMAIL_ADDRESS"]

    #monto o corpo do email
    if not body:
        bodyFile = open(os.environ["BODY_PATH"], 'rt',encoding = 'utf8')
        body = bodyFile.read()
        bodyFile.close()
    emailMsg.attach(MIMEText(body,subtype))

    if file:
        pdfFile = open("certificates/"+file, 'rb')
        pdf = MIMEBase("application","octet-stream")
        pdf.set_payload(pdfFile.read())
        encoders.encode_base64(pdf)

        pdf.add_header("Content-Disposition", f"attachment; filename={file}")
        pdfFile.close()
        emailMsg.attach(pdf)

    return emailMsg
