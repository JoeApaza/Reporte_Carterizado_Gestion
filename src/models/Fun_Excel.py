
import logging
import xlwings as xw
import os
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path
from jinja2 import Template
import win32com.client as win32
import locale
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
 
# Obtener el nombre del día actual en español
nombre_dia = datetime.now().strftime('%A').capitalize()
nombre_mes = datetime.now().strftime('%B').capitalize() 


fecha_actual = datetime.now()
año = fecha_actual.strftime('%Y')
mes = fecha_actual.strftime('%m')
dia = fecha_actual.strftime('%d')
fecha_usar=fecha_actual.strftime('%d/%m/%Y')



def Macros(ruta_libro_formato,hoja,rango_inicio,dataframe,Nombre_Macro):
        logging.info('Iniciando proceso para ejecucion de macro')
        #dataframe=dataframe.to_pandas()
        app = xw.App(visible=False)
        logging.info('Ejecutando macro sin hacer visible que se abra excel')
        wb = xw.Book(ruta_libro_formato)
        logging.info('Se abrio archivo excel')
        sheet = wb.sheets[hoja]
        sheet.range(rango_inicio).value = dataframe.values
        try:
            wb.macro(Nombre_Macro).run()
            logging.info(f"La macro '{Nombre_Macro}' se ha ejecutado con éxito.")
        except Exception as e:
            logging.info(f"Error al ejecutar la macro: {e}")
        wb.close()
        logging.info('Se cerro archivo excel')
        app.quit()
        logging.info('hacer visible que se abra excel')


     

def Eliminar_Excel(ruta_libro):
    if os.path.exists(ruta_libro):
        os.remove(ruta_libro)
        logging.info(f"El archivo {ruta_libro} se ha eliminado con éxito.")
    else:
        logging.info(f"El archivo {ruta_libro} no existe.")

     

def leer_html(ruta_html,dataframe1,dataframe2,var1,var2):
    ruta_html=Path(ruta_html)
    with open(ruta_html,'r',encoding='utf-8') as file:
         template_html=file.read()
         template=Template(template_html)
         return template.render(columns1=dataframe1.columns,data1=dataframe1,columns2=dataframe2.columns,data2=dataframe2,var1=var1,var2=var2)
    
def enviar_correo(html):
    outlook= win32.Dispatch('outlook.application')
    mail=outlook.createitem(0)
    mail.subject=f"Reporte de Gestión - Cobranza Corporativa- {nombre_dia} {dia}.{mes}.{año}"
    #attachment = os.path.abspath(ruta_libro)
    #mail.Attachments.Add(attachment)
    #mail.to='joel.maita@claro.com.pe'
    mail.to='cobranzacorporativa@claro.com.pe'
    mail.CC='joel.maita@claro.com.pe;rlavado@claro.com.pe;maria.chumpitaz@claro.com.pe;liset.flores@claro.com.pe'
    mail.HTMLBody=html
    #mail.SentOnBehalfOfName = 'joel.maita@claro.com.pe'
    #$mail.SentOnBehalfOfName = 'recuperacorp@claro.com.pe'
    mail.GetInspector 
    mail.Send()
    logging.info(f"Informe de recaudacion Corporativa enviado correctamente" )


