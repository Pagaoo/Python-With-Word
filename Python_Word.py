from docx import Document

documento = Document()  # Criando um documento

# Abrindo arquivo já existente
arquivo = open('arquivo_existente.docx', 'rb')

documento = Document(arquivo)

arquivo.close()

documento = Document('arquivo_existente.docx')

# Salvando arquivo novo
documento.save('arquivo_salvo.docx')