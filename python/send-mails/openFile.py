import pandas

def openExcel(fileName: str)-> pandas.DataFrame:
    print("Abrindo o excel "+ fileName)
    table = pandas.read_excel("planilhas/"+fileName)
    print("Excel "+ fileName + " aberto")
    return table