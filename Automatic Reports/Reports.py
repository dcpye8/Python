import os
import time
import numpy as np
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph, Frame, PageTemplate
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from PIL import Image
from datetime import date

PATH = r'C:\Users\diogo.carneiro\Desktop\Nova pasta'

def join_reports(PATH):
    # If there's more than one item in the order joins all .xlsx files. returns a dataframe
    df_report = pd.DataFrame(data=None)
    for files in os.listdir(PATH):
        if files[:3] == 'rel':
            report = pd.read_excel(files)
            df_report = pd.concat([df_report, report], ignore_index=True)

    return df_report

def internal_report(df_report, color_intensity):
    # produces a final dataframe with all necessary internal report columns
    columns = ['Lotes', 'Rem', 'Metros Brutos', 'Metros Líq.', 'Bonus %', '1P', '2P', '3P', '4P', 'Total Def.', 'Def. Calc.', 'OK/NOK']
    final_data = pd.DataFrame(data=None, columns=columns)
    final_data['Lotes'] = df_report['Lote']
    final_data['Rem'] = df_report['Remessa']
    final_data['Metros Brutos'] = (df_report['Metros Brutos']).round(decimals=2)
    final_data['Metros Líq.'] = (df_report['Estoque total']).round(decimals=2)
    final_data['Bonus %'] = (df_report['Bonus']).round(decimals=2)
    final_data['1P'] = df_report['Def. 1P']
    final_data['2P'] = df_report['Def. 2P']
    final_data['3P'] = df_report['Def. 3P']
    final_data['4P'] = df_report['Def. 4P']
    final_data['Total Def.'] = df_report['TTL Defect.']
    max_def = 12 if color_intensity == '0-Clara' else 8
    final_data['Def. Calc.'] = final_data['Total Def.']-((max_def*final_data['Metros Brutos'])/60)
    final_data['Def. Calc.'] = final_data['Def. Calc.'].round(decimals=0)
    final_data['OK/NOK'] = np.where(final_data['Def. Calc.'] <= 0.49, 'OK', 'NOK')

    return final_data

def client_report(df_report):
    columns = ['Roll nº', 'Batch', 'Gross Meters', 'Net Meters', 'Bonus %', 'Total Faults']
    final_data = pd.DataFrame(data=None, columns=columns)
    final_data['Roll nº'] = df_report['Lote']
    final_data['Batch'] = df_report['Remessa']
    final_data['Gross Meters'] = (df_report['Metros Brutos']).round(decimals=2)
    final_data['Net Meters'] = (df_report['Estoque total']).round(decimals=2)
    final_data['Bonus %'] = (df_report['Bonus']).round(decimals=2)
    final_data['Total Faults'] = df_report['TTL Defect.']
    
    return final_data

def order_information(df_report, df_analise, order, item):
    # Sums up all the order information to later populate the report with it
    
    filter_rows = df_analise.loc[(df_analise['Documento de vendas']==order) & (df_analise['ItmCmM']==item)]
    color_tone, color_intensity = filter_rows[['Tonalidade', 'Intensidade']].iloc[0]
    material = filter_rows['Material'].iloc[0]
    material_name = filter_rows['Denominação'].iloc[0]
    client_name = filter_rows['Nome Cliente Final'].iloc[0]
    color_number = filter_rows['Val.caract.'].iloc[0]
    width_ft = df_report['Larg.'].iloc[0]
    gross_meters = (df_report['Metros Brutos'].sum()).round(decimals=2)
    net_meters = (df_report['Estoque total'].sum()).round(decimals=2)
    bonif = (((gross_meters-net_meters)/gross_meters)*100).round(decimals=2)

    order_info = {'Material': material,
                  'Família': material_name,
                  'Cliente': client_name,
                  'Tonalidade': color_tone,
                  'Intensidade': color_intensity,
                  'Nº Cor': color_number,
                  'Larg. Teórica': width_ft,
                  'Metros Brutos': gross_meters,
                  'Metros Líq.': net_meters,
                  'Bonificação %': bonif,
                  'Encomenda': order,
                  'Item': item,
                  }

    return order_info

def report_layout(order_info, final_data, report_type):
    # Creates a layout (Internal Report) and formatting for the report with a header, a table and a footer.
    if report_type == 'INTERNAL':
        filename = str(order_info['Encomenda']) + '-' + str(order_info['Item']) + '-INTERNAL-REPORT' + '.pdf'
    elif report_type == 'CLIENT':
        filename = str(order_info['Encomenda']) + '-' + str(order_info['Item']) + '-CLIENT-REPORT' + '.pdf'
            
    pdf = SimpleDocTemplate(
        filename,
        pagesize=landscape(A4),
        )
    
    def header(canvas, pdf):
        # Gives all the order information as a header
        canvas.saveState()
        IMAGE_PATH = r'C:\Users\diogo.carneiro\Desktop\Nova pasta\logo.jpg'
        
        if report_type == 'INTERNAL':
            canvas.setFont('Helvetica-Bold', 11)
            canvas.drawString(25, 570, 'Artigo:')
            canvas.drawString(25, 555, 'Nome Artigo:')
            canvas.drawString(25, 540, 'Ton/Inten:')
            canvas.drawString(25, 525, 'Cor:')
            canvas.drawString(300,570,'Larg. Teórica:')
            canvas.drawString(300,555,'Qtd. Brutos (m):')
            canvas.drawString(300,540,'Qtd. Líq. (m)')
            canvas.drawString(300,525,'Bonificação (%):')
            canvas.drawString(600,570,'Encomenda:')
            canvas.drawString(600,555,'Item:')
            canvas.drawString(600,540,'Cliente:')
            canvas.drawString(600,525,'Data:')
            
            canvas.setFont('Helvetica', 11)
            canvas.drawString(100, 570, order_info['Material'])
            canvas.drawString(100, 555, order_info['Família'])
            canvas.drawString(100, 540, order_info['Tonalidade'])
            canvas.drawString(100, 525, order_info['Nº Cor'])
            canvas.drawString(390,570, order_info['Larg. Teórica'])
            canvas.drawString(390,555, str(order_info['Metros Brutos']))
            canvas.drawString(390,540, str(order_info['Metros Líq.']))
            canvas.drawString(390,525, str(order_info['Bonificação %']))
            canvas.drawString(675,570, str(order_info['Encomenda']))
            canvas.drawString(675,555, str(order_info['Item']))
            canvas.drawString(675,540, order_info['Cliente'])
            canvas.drawString(675,525, str(date.today()))
            
        elif report_type == 'CLIENT':
            canvas.setFont('Helvetica-Bold', 11)
            canvas.drawString(25, 570, 'Article:')
            canvas.drawString(25, 555, 'Article Name:')
            canvas.drawString(25, 540, 'Color:')
            canvas.drawString(300,570,'Client:')
            canvas.drawString(300,555,'Order:')
            canvas.drawString(300,540,'Date:')
            
            canvas.setFont('Helvetica', 11)
            canvas.drawString(100, 570, order_info['Material'])
            canvas.drawString(100, 555, order_info['Família'])
            canvas.drawString(100, 540, order_info['Nº Cor'])
            canvas.drawString(390,570, order_info['Cliente'])
            canvas.drawString(390,555, str(order_info['Encomenda']))
            canvas.drawString(390,540, str(date.today()))

        canvas.rect(25, 515, 800, 1, fill=1)
        canvas.drawImage(IMAGE_PATH, 600, 0)
        canvas.restoreState()

    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height - 2*mm, id='normal')
    pdf.addPageTemplates([PageTemplate(id='header', frames=[frame], onPage=header), PageTemplate(id='normal', frames=[frame])])

    # Transforms a DataFrame to a List to populate the table
    data = np.array(final_data).tolist()
    # Inserts the columns names as the first list inside the data list
    data.insert(0,final_data.columns.tolist())

    table = Table(data)
    # Highlights rows when i%2=0
    for i in range(1, len(data)):
        if i % 2 == 0:
            bc = colors.lightgrey
        else:
            bc = colors.white
        ts = TableStyle([('BACKGROUND',(0, i), (-1, i), bc),
                         ('ALIGN', (0,0), (-1,-1), 'CENTER')])
        table.setStyle(ts)
    # Table Style (columns name)
    style = TableStyle([
        ('BACKGROUND', (0,0), (13,0), colors.white),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 14),])

    table.setStyle(style)
    elems = []
    elems.append(table)
    
    pdf.build(elems)

def main():
    # Reads the orders database (.xlsx file)
    df_analise = pd.read_excel('analise.xlsx')
    # Calls join_reports() function to join multiple files if it's necessary (returns a dataframe with all rolls and it's information)
    df_report = join_reports(PATH) 
    # Gets the order/item number to filter the order database using the order_information() function
    order, item = df_report[['Encomenda', 'Item']].iloc[0] 
    order_info = order_information(df_report, df_analise, order, item)
    color_intensity = order_information(df_report, df_analise, order, item)['Intensidade']
    # Asks the user what kind or report it should export and calls the right function accordingly
    report_type = (input('Choose the kind of report to export (Internal or Client): ')).upper()
    if report_type == 'INTERNAL':
        final_data = internal_report(df_report, color_intensity)
    elif report_type == 'CLIENT':
        final_data = client_report(df_report)
    else:
        print('Type of report not found!')
    # Calls the designed layout through the report_layout function after having all the data
    report_layout(order_info, final_data, report_type)
     
if __name__ == '__main__':
    main()
    print(time.process_time())
