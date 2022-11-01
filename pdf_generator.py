import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet

lime = pd.read_excel("Z:/CGETNO/CGETNO/Estágio/Davi/lime.xlsx") 
terrasIndigenas = pd.read_excel("Z:/CGETNO/CGETNO/Estágio/Davi/terrasindigenas.xlsx")
etnias = pd.read_excel("Z:/CGETNO/CGETNO/Estágio/Davi/etnias.xlsx")

id_resposta = 166

lime.set_index('ID da resposta',inplace=True)
terrasIndigenas.set_index('ID',inplace=True)
etnias.set_index('ID',inplace=True)

my_Style=ParagraphStyle('My Para style',
fontName='Times-Bold',
fontSize = 16,
leading = 20,
alignment = 1)

styleSheet = getSampleStyleSheet()
bt = styleSheet['BodyText']

btL = ParagraphStyle('BodyTextTTLower',parent=bt, textTransform='uppercase')

titulo = lime['Nome do Projeto'][id_resposta]

c = canvas.Canvas('projetos_etno.pdf')
# draw = Drawing(500, 200)
c.saveState()
c.setLineWidth(1.5)
c.rect(10,690,575,110, stroke=1, fill=0)
c.restoreState()
c.setFont('Times-Bold', 16)
c.drawCentredString(300,770, lime['Coordenação Regional ou Frente de Proteção Etnoambiental'][id_resposta])
c.drawImage('logo-png.png', 500, 770, 50, 70)
# c.saveState()
c.setFont('Times-Bold', 15) 
if len(titulo) < 70:
    paragrafo = Paragraph(lime['Nome do Projeto'][id_resposta], style=my_Style)
    paragrafo.wrapOn(c, 530, 400)
    paragrafo.drawOn(c, 30, 730)
else:    
    paragrafo = Paragraph(lime['Nome do Projeto'][id_resposta], style=my_Style)
    paragrafo.wrapOn(c, 530, 400)
    paragrafo.drawOn(c, 30, 700)

# c.restoreState()
c.saveState()
c.setLineWidth(1.5)
c.rect(10,510,575,150)
c.rect(10,320,575,165)
c.rect(10,130,575,165)
c.restoreState()
c.drawString(250,640, "Identificação")
c.drawString(250,465, "Descrição Geral")
c.drawString(250,275, "Terras Indígenas")
c.setFont('Times-Bold', 14)
c.drawString(30,610,'Técnico Responsável:')
c.drawString(30,580, 'Lotação:')
c.drawString(30,550, 'CTLs Envolvidas:')
c.drawString(30, 415, 'Objetivo Geral:')
c.drawString(30, 370, 'Qtde de Comunidades/Aldeias diretamente atendidas:')
c.drawString(30, 350, 'Qtde de Famílias diretamente atendidas:')
c.drawString(30, 330, 'Esse projeto visa: ')
c.drawString(30, 230, 'Terras índigenas: ')
c.drawString(30, 180, 'Etnias contempladas: ')
c.setFont('Times-Roman', 13)
c.drawString(162,612, lime['Técnico Responsável'][id_resposta])
c.drawString(162,582, lime['Lotação do Técnico Responsável'][id_resposta])
c.drawString(162, 552, lime['CTL 1'][id_resposta])
objetivo = Paragraph(lime['Objetivo Geral'][id_resposta][:240], style= btL)
objetivo.wrapOn(c, 420, 20)
objetivo.drawOn(c, 162, 400)
c.drawString(360, 370, str(lime['Quantidade de Comunidades/Aldeias diretamente atendidas'][id_resposta]))
c.drawString(360, 350, str(lime['Quantidade de Famílias diretamente atendidas'][id_resposta]))
if lime['Este projeto visa .... [Implantar nova(s) atividade(s) produtiva(s)]'][id_resposta] == 'Sim':
    c.drawString(162, 330, 'Implantar nova(s) atividade(s) produtiva(s)')
else:
    c.drawString(162, 330, 'Fortalecer atividade(s) produtiva(s) pré-existente(s)')


terras = []
for indice, linha in terrasIndigenas.iterrows():
    if indice == id_resposta:
        terras.append(linha['Terras Indígenas'])

print(terras)
    
terras_indigenas = str(terras).strip('[]') 
print(terras_indigenas)
terras = Paragraph(terras_indigenas, style= btL)
print(terras)
terras.wrapOn(c, 420, 20)
terras.drawOn(c, 162, 227)


valores = []
for indice, linha in etnias.iterrows():
    if indice == id_resposta:
        valores.append(linha['Etnias'])
 
etnias_id = str(valores).strip('[]') 
etn = Paragraph(etnias_id, style= btL)
print(etn)
etn.wrapOn(c, 420, 20)
etn.drawOn(c, 162, 180)
  
# c.drawString(30,640,'Coordenação Regional:')
# c.drawString(160,643, str (col3))
# c.drawString(30,615,'Data e hora:')
# c.drawString(160,618, str (col4))


# # c.drawString(30,600, str (col1))
# # c.drawString(100,600, str (col2))
# # c.drawString(350,600, str (col3))              
           
c.save()

