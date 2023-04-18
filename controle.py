from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QFileDialog
import pandas as pd
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
import gspread
import os
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

CODE = ''

credencial = {
  
}

gc = gspread.service_account_from_dict(credencial)

sh = gc.open_by_key(CODE)

ws = sh.worksheet('Projetos 2023')

ti = sh.worksheet('Terras Indigenas')

lime = pd.DataFrame(ws.get_all_records())
lista_terras = pd.DataFrame(ti.get_all_records())
    
my_Style=ParagraphStyle('My Para style',
fontName='Times-Bold',
fontSize = 16,
leading = 20,
alignment = 1)

styleSheet = getSampleStyleSheet()
bt = styleSheet['BodyText']

btL = ParagraphStyle('BodyTextTTLower',parent=bt, textTransform='uppercase')

#lista_cr.rename(columns={"ID": "Coordenação Regional ou Frente de Proteção Etnoambiental"},  inplace=True)
#lista_cr.to_excel('teste2.xlsx')
#lime = lime.merge(lista_cr)
#lime['Coordenação Regional ou Frente de Proteção Etnoambiental'] = lime['CR']

def escolherCR():

    cr = generator.comboBox_2.currentText()
    final_df = lime[lime['Coordenação Regional ou Frente de Proteção Etnoambiental'] == cr]

    generator.frame_5.close()
    projetos = []
    for indice, linha in final_df.iterrows():
            projetos.append(linha['Nome do Projeto'])

    generator.comboBox.clear()
    
    if len(projetos) == 0:
        projetos.append('Nenhum projeto enviado.')
    
    generator.comboBox.addItems(projetos)

def gerar_pdf():

        projeto = generator.comboBox.currentText()
        
        if projeto == 'Nenhum projeto enviado.':
            QMessageBox.about(generator, "Alerta", "CR sem projeto ainda.")
            
        else: 
            final_df = lime[lime['Nome do Projeto'] == projeto]
            
            id_resposta = 0
            
            for indice, linha in final_df.iterrows():
                id_resposta = linha['ID da resposta']
        
            lime2 = lime.set_index('ID da resposta')
            pdf_name = f'projeto_etno{id_resposta}.pdf'
            response = QFileDialog.getExistingDirectory(caption='Select a folder')
            print(response)
            save_name = os.path.join(response, pdf_name)
            c = canvas.Canvas(save_name)
            titulo = lime2['Nome do Projeto'][id_resposta]

            c.saveState()
            c.setLineWidth(1.5)
            c.rect(10,690,575,110, stroke=1, fill=0)
            c.restoreState()
            c.setFont('Times-Bold', 16)
            c.drawCentredString(300,770, lime2['Coordenação Regional ou Frente de Proteção Etnoambiental'][id_resposta])
            c.drawImage('logo.png', 500, 770, 50, 70)
            # c.saveState()
            c.setFont('Times-Bold', 15) 

            if len(titulo) < 70:
                paragrafo = Paragraph(lime2['Nome do Projeto'][id_resposta], style=my_Style)
                paragrafo.wrapOn(c, 530, 400)
                paragrafo.drawOn(c, 30, 730)
            else:    
                paragrafo = Paragraph(lime2['Nome do Projeto'][id_resposta], style=my_Style)
                paragrafo.wrapOn(c, 530, 400)
                paragrafo.drawOn(c, 30, 700)

            # c.restoreState()
            c.saveState()
            c.setLineWidth(1.5)
            c.rect(10,510,575,150)
            c.rect(10,10,575,472)
            #c.rect(10,130,575,165)
            c.restoreState()
            c.drawString(250,640, "Identificação")
            c.drawString(250,465, "Descrição Geral")
            #c.drawString(250,275, "Terras Indígenas")
            c.setFont('Times-Bold', 14)
            c.drawString(30,610,'Técnico Responsável:')
            c.drawString(30,580, 'Lotação:')
            c.drawString(30,550, 'CTLs Envolvidas:')
            c.drawString(30, 440, 'Objetivo Geral:')
            c.drawString(30, 335, 'Metodologia')
            c.drawString(30, 180, 'Justificativa')
            c.setFont('Times-Roman', 13)
            c.drawString(162,612, lime2['Técnico Responsável'][id_resposta])
            c.drawString(162,582, lime2['Lotação do Técnico Responsável'][id_resposta])
            c.drawString(162, 552, lime2['CTL 1'][id_resposta])

            objetivoTam = lime2['Objetivo Geral'][id_resposta]
            objetivo = Paragraph(lime2['Objetivo Geral'][id_resposta][:500], style= btL)

            if len(objetivoTam) <= 200:
                objetivo.wrapOn(c, 540, 20)
                objetivo.drawOn(c, 30, 400)
            elif len(objetivoTam) <= 330:
                objetivo.wrapOn(c, 540, 20)
                objetivo.drawOn(c, 30, 375)
            else:
                objetivo.wrapOn(c, 540, 20)
                objetivo.drawOn(c, 30, 365)        

            metodologiaTam = lime2['Metodologia'][id_resposta]
            metodologia = Paragraph(lime2['Metodologia'][id_resposta][:900], style= btL)

            if len(metodologiaTam) <= 200:
                metodologia.wrapOn(c, 540, 20)
                metodologia.drawOn(c, 30, 295)
            elif len(metodologiaTam) <= 330:
                metodologia.wrapOn(c, 540, 20)
                metodologia.drawOn(c, 30, 275)
            elif len(metodologiaTam) <= 500:
                metodologia.wrapOn(c, 540, 20)
                metodologia.drawOn(c, 30, 235)     
            elif len(metodologiaTam) <= 700:
                metodologia.wrapOn(c, 540, 20)
                metodologia.drawOn(c, 30, 215)         
            else:    
                metodologia.wrapOn(c, 540, 20)
                metodologia.drawOn(c, 30, 195)

            justificativaTam = lime2['Justificativa'][id_resposta][:900]
            justificativa = Paragraph(lime2['Justificativa'][id_resposta][:900], style= btL)
    
            if len(justificativaTam) <= 200:
                justificativa.wrapOn(c, 540, 20)
                justificativa.drawOn(c, 30, 125)
            elif len(justificativaTam) <= 250:
                justificativa.wrapOn(c, 540, 20)
                justificativa.drawOn(c, 30, 120)
            elif len(justificativaTam) <= 330:
                justificativa.wrapOn(c, 540, 20)
                justificativa.drawOn(c, 30, 155)
            elif len(justificativaTam) <= 500:
                justificativa.wrapOn(c, 540, 20)
                justificativa.drawOn(c, 30, 90)  
            elif len(justificativaTam) <= 700:
                justificativa.wrapOn(c, 540, 20)
                justificativa.drawOn(c, 30, 70)                
            else:    
                justificativa.wrapOn(c, 540, 20)
                justificativa.drawOn(c, 30, 20)

            c.showPage()

            for i in range(1):  

                c.saveState()
                c.setLineWidth(1.5)
                c.rect(10,690,575,110, stroke=1, fill=0)
                c.restoreState()
                c.setFont('Times-Bold', 16)
                c.drawCentredString(300,770, lime2['Coordenação Regional ou Frente de Proteção Etnoambiental'][id_resposta])
                c.drawImage('logo.png', 500, 770, 50, 70)    
                
                if len(titulo) < 70:
                    paragrafo = Paragraph(lime2['Nome do Projeto'][id_resposta], style=my_Style)
                    paragrafo.wrapOn(c, 530, 400)
                    paragrafo.drawOn(c, 30, 730)
                else:    
                    paragrafo = Paragraph(lime2['Nome do Projeto'][id_resposta], style=my_Style)
                    paragrafo.wrapOn(c, 530, 400)
                    paragrafo.drawOn(c, 30, 700)


                c.setLineWidth(1.5)
                c.rect(10,260,575,400)    
                c.setFont('Times-Bold', 14)
                c.drawString(250, 640, "Descrição Geral")
                c.drawString(30, 612, 'Qtde de Comunidades/Aldeias diretamente atendidas:')
                c.drawString(30, 582, 'Qtde de Famílias diretamente atendidas:')
                c.drawString(30, 552, 'Esse projeto visa: ')
                c.drawString(30, 512, 'Terras índigenas: ')
                c.drawString(30, 472, 'Etnias contempladas: ')
                c.drawString(30, 422, 'Parcerias: ')
                
                c.setFont('Times-Roman', 13)
                c.drawString(360, 612, str(lime2['Quantidade de Comunidades/Aldeias diretamente atendidas'][id_resposta]))
                c.drawString(360, 582, str(lime2['Quantidade de Famílias diretamente atendidas'][id_resposta]))

                if lime2['Este projeto visa .... [Implantar nova(s) atividade(s) produtiva(s)]'][id_resposta] == 'Sim':
                    c.drawString(162, 552, 'Implantar nova(s) atividade(s) produtiva(s)')
                elif lime2['Este projeto visa .... [Fortalecer atividade(s) produtiva(s) pré-existente(s)]'][id_resposta] == 'Sim':
                    c.drawString(162, 552, 'Fortalecer atividade(s) produtiva(s) pré-existente(s)')
                else:
                    c.drawString(162, 552, 'Realizar diagnósticos e levantamento de dados')


                terras = []

                for indice, linha in lime2.iterrows():
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

                
                terrasLista = list(filter(None, terras))
                #terrasIndigenas = pd.DataFrame(terras, columns=['terrai_cod'])
                #terrasIndigenas = terrasIndigenas.merge(lista_terras, how = 'inner')

                #terrasIndigenas.drop(['gid', 'terrai_cod', 'etnia_nome', 'municipio_', 'uf_sigla', 'superficie', 'fase_ti', 'modalidade', 'reestudo_t', 
                #                   'cr', 'faixa_fron', 'undadm_cod', 'undadm_nom', 'undadm_sig', 'dominio_un'], axis=1, inplace=True)

                #terra = terrasIndigenas['terrai_nom'].tolist()                   
                terras_indigenas = str(terrasLista).strip('[]')  

                terras = Paragraph(terras_indigenas, style= btL)
                terras.wrapOn(c, 420, 20)
                terras.drawOn(c, 162, 512)

                valores = []

                for indice, linha in lime2.iterrows():
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
                etn.drawOn(c, 164, 472)
                
                parceiros = []
                
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGGAM/DPDS]'][id_resposta] == 'Sim':
                        parceiros.append('CGGAM/DPDS')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGLIC/DPDS]'][id_resposta] == 'Sim':
                        parceiros.append('CGLIC/DPDS')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGPC/DPDS]'][id_resposta] == 'Sim':
                        parceiros.append('CGPC/DPDS')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGPDS/DPDS]'][id_resposta] == 'Sim':
                        parceiros.append('CGPDS/DPDS')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGAF/DPT]'][id_resposta] == 'Sim':
                        parceiros.append('CGAF/DPT')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGGEO/DPT]'][id_resposta] == 'Sim':
                        parceiros.append('CGGEO/DPT')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGID/DPT]'][id_resposta] == 'Sim':
                        parceiros.append('CGID/DPT')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGMT/DPT]'][id_resposta] == 'Sim':
                        parceiros.append('CGMT/DPT')
                if lime2['Com qual ou quais Coordenações-Gerais o projeto é compartilhado? [CGIIRC/DPT]'][id_resposta] == 'Sim':
                        parceiros.append('CGIIRC/DPT')

                
                listaParceiros = list(parceiros)
                listaDeParceiros = str(listaParceiros).strip('[]') 
                
                terras = Paragraph(listaDeParceiros, style= btL)
                terras.wrapOn(c, 420, 20)
                terras.drawOn(c, 162, 422)
                c.showPage()
                
            for i in range(1):      
                    
                c.saveState()
                c.setLineWidth(1.5)
                c.rect(10,690,575,110, stroke=1, fill=0)
                c.restoreState()
                c.setFont('Times-Bold', 16)
                c.drawCentredString(300,770, lime2['Coordenação Regional ou Frente de Proteção Etnoambiental'][id_resposta])
                c.drawCentredString(300, 670, 'Quadro Orçamentário')
                c.drawImage('logo.png', 500, 770, 50, 70)    
                
                if len(titulo) < 70:
                    paragrafo = Paragraph(lime2['Nome do Projeto'][id_resposta], style=my_Style)
                    paragrafo.wrapOn(c, 530, 400)
                    paragrafo.drawOn(c, 30, 730)
                else:    
                    paragrafo = Paragraph(lime2['Nome do Projeto'][id_resposta], style=my_Style)
                    paragrafo.wrapOn(c, 530, 400)
                    paragrafo.drawOn(c, 30, 700)
                    
                data = [
                    [lime2['1º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 1º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['2º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 2º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['3º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 3º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['4º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 4º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['5º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 5º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['6º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 6º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['7º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 7º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['8º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 8º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['9º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 9º subelemento de despesa selecionado ']
                        [id_resposta]], 
                    [lime2['10º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 10º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['11º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 11º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['12º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 12º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['13º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 13º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['14º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 14º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['15º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 15º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['16º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 16º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['17º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 17º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['18º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 18º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['19º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 19º subelemento de despesa selecionado ']
                        [id_resposta]],
                    [lime2['20º Subelemento de Despesa'][id_resposta], lime2['Valor solicitado para o 20º subelemento de despesa selecionado ']
                        [id_resposta]],
                ]
                
                df = pd.DataFrame(
                    data, columns=['SubElemento de despesa', 'Valor'])
                    
                data2 = df.loc[~(df['Valor'] == ' R$  -   ')]
                data2 = df.loc[~(df['Valor'] == '0')]
                data2 = df.loc[~(df['Valor'] == '')]
            
                valorElementos = []
                valorNumerico = []
                
                for indice, linha in data2.iterrows():
                    tamanho_valor = str(linha['Valor'])
                    if len(tamanho_valor) > 6:
                        if tamanho_valor[-2:] == "00":
                            linha['Valor'] = int(tamanho_valor[:-2]) 
                
                for indice, linha in data2.iterrows():
                    valorNumerico.append(linha['Valor'])
                    linha['Valor'] = locale.currency(linha['Valor'], grouping=True)
                    valorElementos.append(linha['Valor'])
                
                
                int_list = list(map(float, valorNumerico))
                somaValores = sum(int_list)
                somaValores = locale.currency(somaValores, grouping=True)
                
                data2.loc[len(data2) + 1] = ['Total', somaValores]
                
                lista = data2.values.tolist()
                style = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                                    ('VALIGN',(0,0),(-1,-1),'TOP'),
                                    ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                                    ('BOX',(0,0),(-1,-1),1,colors.black),
                                    ('BOX',(0,0),(0,-1),1,colors.black),
                                    ('FONTSIZE', (0,0), (-1,-1), 12),
                                    ('ALIGN', (0, 0), (-1, -1), 'CENTER')])    
                style.add('BACKGROUND',(0,1),(-1,-1),colors.white)

                t=Table(lista, rowHeights=(30))
                t.setStyle(style)
                
                c.saveState()
                c.setFont('Times-Bold', 16)
                
                if len(data2.index) <= 6:
                    t.wrapOn(c, 800, 800)
                    t.drawOn(c, 50, 400)
                elif len(data2.index) <= 11:
                    t.wrapOn(c, 800, 800)
                    t.drawOn(c, 50, 100)
                elif len(data2.index) >= 13: 
                    t.wrapOn(c, 800, 800)
                    t.drawOn(c, 50, 25)
                    
                c.restoreState()

                c.showPage()   
                    
            c.save()
                
            QMessageBox.about(generator, "Alerta", f"PDF gerado, salvo em {response}")

 
def deletar():
    generator.lineEdit.clear()
    generator.frame_4.close()
    generator.frame_2.close()
    generator.frame_3.close()
    generator.pushButton_3.close()


sys._excepthook = sys.excepthook 
def exception_hook(exctype, value, traceback):
    print(exctype, value, traceback)
    sys._excepthook(exctype, value, traceback) 
    sys.exit(1) 
sys.excepthook = exception_hook     
    
app = QtWidgets.QApplication([])
generator = uic.loadUi("gui3.ui")
generator.comboBox_2.addItems(['CR ALTO SOLIMÕES', 'CR ALTO PURUS', 'CR AMAPÁ E NORTE DO PARÁ', 'CR ARAGUAIA E TOCANTINS', 'CR BAIXO SÃO FRANCISCO', 'CR BAIXO TOCANTINS', 'CR CACOAL',
 'CR CAMPO GRANDE', 'CR CENTRO LESTE DO PARÁ', 'CR CUIABÁ', 'CR DOURADOS', 'CR GUAJARÁ-MIRIM', 'CR GUARAPUAVA', 'CR INTERIOR SUL', 'CR JI-PARANÁ', 'CR JOÃO PESSOA', 'CR JURUÁ', 
 'CR KAYAPÓ SUL DO PARÁ', 'CR LITORAL SUDESTE', 'CR LITORAL SUL', 'CR MADEIRA', 'CR MANAUS', 'CR MARANHÃO', 'CR MÉDIO PURUS', 'CR MINAS GERAIS E ESPÍRITO SANTO', 'CR NORDESTE I', 
 'CR NORDESTE II', 'CR NOROESTE DO MATO GROSSO', 'CR NORTE DO MATO GROSSO', 'CR PASSO FUNDO', 'CR PONTA PORÃ', 'CR RIBEIRÃO CASCALHEIRA', 'CR RIO NEGRO', 'CR RORAIMA', 'CR SUL DA BAHIA', 'CR TAPAJÓS', 
 'CR VALE DO JAVARI', 'CR XAVANTE', 'CR XINGU', 'FPE AWA GUAJÁ', 'FPE CUMINAPANEMA', 'FPE ENVIRA', 'FPE GUAPORÉ', 'FPE MADEIRINHA-JURUENA', 'FPE MADEIRA-PURUS', 'FPE YANOMAMI/YE KUANA',
 'FPE MÉDIO XINGU', 'FPE VALE DO JAVARI', 'FPE WAIMIRI-ATROARI'])
 
generator.pushButton_2.clicked.connect(gerar_pdf)
generator.pushButton_3.clicked.connect(escolherCR)
generator.show()
app.exec()

