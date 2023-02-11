from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

"""
    Método que faz a criação de um certificado em pdf com base em um template e um
    texto

    @param fileName: nome para arquivo
    @param text: Texto que será imprimido no PDF
    @param templateName: nome do template, com a extensão
    @param signature: Assinatura colocada ao fim do certificado com data e local
    @param positionY: Posição do ponteiro na vertical para desenhar o texto

    @return: o nome do pdf
"""
def certificateGenerate(
    fileName: str,
    text: str,
    templateName: str,
    signature: str,
    positionY: int = 416,
) -> str:

    lines = []
    
    i = 0
    line = ""
    for char in text:
        line += char
        if i > 90:
            if  char == " ":
                lines.append(line)
                line = ""
                i = 0
        i+=1
    if line != "":
        lines.append(line)

    cv = canvas.Canvas("certificates/"+fileName,pagesize=(297*mm,210*mm))
    cv.drawImage("templates/"+templateName,0,0,width=842,height=596)

    for line in lines:
        cv.drawCentredString(x=421,y=positionY,text=line, charSpace=1)
        positionY -= 12

    cv.drawCentredString(x=421,y=positionY-24,text=signature)
    cv.save()
    return fileName