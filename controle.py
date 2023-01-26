from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QMessageBox

import pandas as pd
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
import gspread
import os

CODE = '1IMVaXrRNAuUA_MsHHRVbAY5uzuIyB4Z0-1JUbSDcRg0'

credencial = {
  
}

gc = gspread.service_account(filename='key.json')

sh = gc.open_by_key(CODE)

ws = sh.worksheet('Formulário de Projetos de Etnod')
ti = sh.worksheet('Terras')

lime = pd.DataFrame(ws.get_all_records())
lista_terras = pd.DataFrame(ti.get_all_records())

lime.set_index('ID da resposta', inplace=True)

import string

letras = list(string.ascii_lowercase)

for i in letras:
    try:
        terrasIndigenas = pd.read_excel(f"{i}:/CGETNO/CGETNO/Estágio/Davi/terrasindigenas.xlsx")
        etnias = pd.read_excel(f"{i}:/CGETNO/CGETNO/Estágio/Davi/etnias.xlsx")
    except FileNotFoundError:
        pass
   
my_Style=ParagraphStyle('My Para style',
fontName='Times-Bold',
fontSize = 16,
leading = 20,
alignment = 1)

styleSheet = getSampleStyleSheet()
bt = styleSheet['BodyText']

btL = ParagraphStyle('BodyTextTTLower',parent=bt, textTransform='uppercase')

def gerar_pdf():
    try:    
        id = generator.lineEdit.text()
        id_resposta = int(id)

        pdf_name = f'projeto_etno{id}.pdf'
        save_name = os.path.join(os.path.expanduser("~"), "Desktop/", pdf_name)

        c = canvas.Canvas(save_name)
        titulo = lime['Nome do Projeto'][id_resposta]

        c.saveState()
        c.setLineWidth(1.5)
        c.rect(10,690,575,110, stroke=1, fill=0)
        c.restoreState()
        c.setFont('Times-Bold', 16)
        c.drawCentredString(300,770, lime['Coordenação Regional ou Frente de Proteção Etnoambiental'][id_resposta])
        c.drawImage('./adicionais/logo.png', 500, 770, 50, 70)
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
        c.drawString(30, 370, 'Qtde de Comunidades/Aldeias diretamente atendidas:')
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
        c.drawString(360, 370, str(lime['Quantidade de Comunidades/Aldeias diretamente atendidas'][id_resposta]))
        c.drawString(360, 350, str(lime['Quantidade de Famílias diretamente atendidas'][id_resposta]))

        if lime['Este projeto visa .... [Implantar nova(s) atividade(s) produtiva(s)]'][id_resposta] == 'Sim':
            c.drawString(162, 330, 'Implantar nova(s) atividade(s) produtiva(s)')
        else:
            c.drawString(162, 330, 'Fortalecer atividade(s) produtiva(s) pré-existente(s)')


        for indice, linha in lime.iterrows():
            if indice == id_resposta:
                terras.append(linha['Terra Indígena 1'])
                terras.append(linha['Terra Indígena 2'])
                terras.append(linha['Terra Indígena 3'])
                terras.append(linha['Terra Indígena 4'])
                terras.append(linha['Terra Indígena 5'])
                terras.append(linha['Terra Indígena 6'])
                terras.append(linha['Terra Indígena 7'])
                terras.append(linha['Terra Indígena 8'])
                terras.append(linha['Terra Indígena 9'])
                terras.append(linha['Terra Indígena 10'])
                terras.append(linha['Terra Indígena 11'])
                terras.append(linha['Terra Indígena 12'])
                terras.append(linha['Terra Indígena 13'])
                terras.append(linha['Terra Indígena 14'])
                terras.append(linha['Terra Indígena 15'])
                terras.append(linha['Terra Indígena 16'])
                terras.append(linha['Terra Indígena 17'])
                terras.append(linha['Terra Indígena 18'])
                terras.append(linha['Terra Indígena 19'])
                terras.append(linha['Terra Indígena 20'])

        terras = list(filter(None, terras))
        terrasIndigenas = pd.DataFrame(terras, columns=['terrai_cod'])
        terrasIndigenas = terrasIndigenas.merge(lista_terras, how = 'inner')

        terrasIndigenas.drop(['gid', 'terrai_cod', 'etnia_nome', 'municipio_', 'uf_sigla', 'superficie', 'fase_ti', 'modalidade', 'reestudo_t', 
                            'cr', 'faixa_fron', 'undadm_cod', 'undadm_nom', 'undadm_sig', 'dominio_un'], axis=1, inplace=True)

        terra = terrasIndigenas['terrai_nom'].tolist()                   
        terras_indigenas = str(terra).strip('[]')    

        terras = Paragraph(terras_indigenas, style= btL)
        terras.wrapOn(c, 420, 20)
        terras.drawOn(c, 162, 227)

        valores = []
        for indice, linha in lime.iterrows():
            if indice == id_resposta:
                valores.append(linha['Etnia 1'])
                valores.append(linha['Etnia 2'])
                valores.append(linha['Etnia 3'])
                valores.append(linha['Etnia 4'])
                valores.append(linha['Etnia 5'])
                valores.append(linha['Etnia 6'])
                valores.append(linha['Etnia 7'])
                valores.append(linha['Etnia 8'])
                valores.append(linha['Etnia 9'])
                valores.append(linha['Etnia 10'])
                valores.append(linha['Etnia 11'])
                valores.append(linha['Etnia 12'])
                valores.append(linha['Etnia 13'])
                valores.append(linha['Etnia 14'])
                valores.append(linha['Etnia 15'])
                valores.append(linha['Etnia 16'])
                valores.append(linha['Etnia 17'])
                valores.append(linha['Etnia 18'])
                valores.append(linha['Etnia 19'])
                valores.append(linha['Etnia 20'])
        
        etnias_id = list(filter(None, valores))
        etnias_id = str(etnias_id).strip('[]')  
        etn = Paragraph(etnias_id, style= btL)
        etn.wrapOn(c, 420, 20)
        etn.drawOn(c, 162, 180)
                
        c.showPage() 

        for i in range(1):      
                
            c.saveState()
            c.setLineWidth(1.5)
            c.rect(10,690,575,110, stroke=1, fill=0)
            c.restoreState()
            c.setFont('Times-Bold', 16)
            c.drawCentredString(300,770, lime['Coordenação Regional ou Frente de Proteção Etnoambiental'][id_resposta])
            c.drawImage('./adicionais/logo.png', 500, 770, 50, 70)    
            
            if len(titulo) < 70:
                paragrafo = Paragraph(lime['Nome do Projeto'][id_resposta], style=my_Style)
                paragrafo.wrapOn(c, 530, 400)
                paragrafo.drawOn(c, 30, 730)
            else:    
                paragrafo = Paragraph(lime['Nome do Projeto'][id_resposta], style=my_Style)
                paragrafo.wrapOn(c, 530, 400)
                paragrafo.drawOn(c, 30, 700)
            
            data = [
                # ['SUBELEMENTO DE DESPESA', 'VALOR'],
                [lime['1º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 1º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['2º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 2º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['3º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 3º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['4º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 4º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['5º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 5º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['6º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 6º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['7º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 7º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['8º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 8º subelemento de despesa selecionado ']
                    [id_resposta]],
                [lime['9º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 9º subelemento de despesa selecionado ']
                    [id_resposta]], 
                [lime['10º Subelemento de Despesa'][id_resposta], lime[' Valor solicitado para o 10º subelemento de despesa selecionado ']
                    [id_resposta]],
            ]
            
            df = pd.DataFrame(
                data, columns=['SubElemento de despesa', 'Valor'])
            
            #df['Valor'] = df['Valor'].astype(float)
            # df = df.drop(columns = ['Valor'])
            #df['Valor'] = df['Valor'].map("{:.2f}".format)
            #df['Valor'] = df['Valor'].map("R$ {}".format)
    
            data2 = df.dropna()
            data2 = df.loc[~(df['Valor'] == ' R$  -   ')]
            # print(data2)

            lista = data2.values.tolist()
            
            style = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                                ('VALIGN',(0,0),(-1,-1),'TOP'),
                                ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                                ('BOX',(0,0),(-1,-1),1,colors.black),
                                ('BOX',(0,0),(0,-1),1,colors.black),
                                # ('FONTNAME', (0,0), (-1,-1), 'Calibri'),
                                ('FONTSIZE', (0,0), (-1,-1), 12),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER')])    
            style.add('BACKGROUND',(0,0),(1,0),colors.lightblue)
            style.add('BACKGROUND',(0,1),(-1,-1),colors.white)
            
            t=Table(lista, rowHeights=(30))
            t.setStyle(style)
            
            c.saveState()
            c.setFont('Times-Bold', 16)
            if len(data2.index) >= 7:
                t.wrapOn(c, 800, 800)
                t.drawOn(c, 50, 300)
            else: 
                t.wrapOn(c, 800, 800)
                t.drawOn(c, 50, 400)
                
            c.restoreState()

            c.showPage()    
        
        c.save()
        QMessageBox.about(generator, "Alerta", "PDF gerado")

    except KeyError:
        QMessageBox.about(generator, "Alerta", "ID não existente")

    except ValueError:
        QMessageBox.about(generator, "Alerta", "Digite apenas números!")  


def deletar():
    generator.lineEdit.clear()

sys._excepthook = sys.excepthook 
def exception_hook(exctype, value, traceback):
    print(exctype, value, traceback)
    sys._excepthook(exctype, value, traceback) 
    sys.exit(1) 
sys.excepthook = exception_hook     
    
app = QtWidgets.QApplication([])
generator = uic.loadUi("./adicionais/gui.ui")
generator.pushButton.clicked.connect(deletar)
generator.pushButton_2.clicked.connect(gerar_pdf)


generator.show()
app.exec()
