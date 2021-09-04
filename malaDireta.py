import os
import win32com.client as win32 # pip install pywin32

dirCentral = os.getcwd()
nome_excel = "Nome do excel.xlsx" #nome do excel que vc quer pegar
pasta_destino = os.path.join(dirCentral, "PlanejamentoSeparados") #localização da pasta que irá guardar o word
#print(pasta_destino)

""" 
Criando a instância da aplicação word doc
"""

wordApp = win32.Dispatch("Word.application")
wordApp.Visible = True

""" 
Abrindo Template Word + Abrindo Página do Excel
 """
docRaiz = wordApp.Documents.Open(os.path.join(dirCentral, "NOME DO DOCUMENTO.docx")) #Nome do docx
#print(type(docRaiz))
mala_direta = docRaiz.MailMerge
mala_direta.OpenDataSource(
    Name:=os.path.join(dirCentral, nome_excel),
    sqlstatement:="SELECT * FROM [NOME DA Página do EXCEL$]"  # Select no nome da página do excel
) 

registro_contagem = mala_direta.DataSource.RecordCount

""" 
Realizando a mala direta
"""
for i in range(1, registro_contagem + 1):
    mala_direta.DataSource.ActiveRecord = i
    mala_direta.DataSource.FirstRecord = i
    mala_direta.DataSource.LastRecord = i

    mala_direta.Destination = 0
    mala_direta.Execute(False)

    #Pegando o valor guardado
    nome_base = mala_direta.DataSource.DataFields('Nome da primeira Coluna raiz'.replace(' ', '_')).Value  #Nome da coluna PrimaryKey
    
    docAlvo = wordApp.ActiveDocument

    """ 
    Salvando arquivos no Word
    """ 
    
    docAlvo.SaveAs2(os.path.join(pasta_destino, nome_base + '.docx'), 16) #Salvando docx
    docAlvo.ExportAsFixedFormat(os.path.join(pasta_destino, nome_base), exportformat:=17) #Salvando PDF

    """ 
    Fechando documento alvo
    """

    docAlvo.Close(False)
    docAlvo = None

docRaiz.MailMerge.MainDocumentType = -1
