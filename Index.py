import logging
import os
import pandas as pd
import shutil
from datetime import datetime, timedelta
from dotenv import load_dotenv
from src.database.db_oracle import close_connection_db,read_database_db,leer_sql,get_connection,dtypes
from src.routes.Rutas import ruta_Detalle_Deuda_Carterizado,ruta_env,ruta_html,ruta_libro_Formato
from src.models.Fun_Excel import Macros,Eliminar_Excel,leer_html,enviar_correo
import locale

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

fecha_actual = datetime.now()
año = fecha_actual.strftime('%Y')
mes = fecha_actual.strftime('%m')
dia = fecha_actual.strftime('%d')
fecha_usar=fecha_actual.strftime('%d/%m/%Y')

fecha_anteayer = fecha_actual - timedelta(days=2)
año_1 = fecha_anteayer.strftime('%Y')
mes_1 = fecha_anteayer.strftime('%m')
dia_1 = fecha_anteayer.strftime('%d')

logging.basicConfig(format="%(asctime)s::%(levelname)s::%(message)s",   
                    datefmt="%d-%m-%Y %H:%M:%S",    
                    level=10,   
                    filename='.//src//utils//log//app.log',filemode='w')


load_dotenv(ruta_env)
Conexion_Opercom=get_connection(os.getenv('USER_DB'),os.getenv('PASSWORD_DB'),os.getenv('DNS_DB'))


Df_Detalle_Deuda_Carterizado=read_database_db(leer_sql(ruta_Detalle_Deuda_Carterizado),Conexion_Opercom,dtypes)
Df_Detalle_Deuda_Carterizado=Df_Detalle_Deuda_Carterizado.to_pandas()


Macros(ruta_libro_Formato,'Detalle_Deuda','A2',Df_Detalle_Deuda_Carterizado,'Reporte_Deuda_Carterizado_Gestión')


df_1=Df_Detalle_Deuda_Carterizado.loc[(Df_Detalle_Deuda_Carterizado['OPERADORES'] == 'NO OPERADOR')]

df_1_i = df_1.pivot_table(index=['GESTOR'], columns='TRAMO_VENCIMIENTO', values='DEUDA_SOLES',aggfunc='sum').reset_index().fillna(0)
df_1_t = df_1.pivot_table(index=['GESTOR'], values='DEUDA_SOLES',aggfunc='sum').reset_index().fillna(0)
df_1_f = pd.merge(df_1_i, df_1_t, on=['GESTOR'], how='outer').sort_values(by='DEUDA_SOLES', ascending=False)
df_1_f['1. Por vencer']=df_1_f['1. Por vencer'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['2. 1 a 30']=df_1_f['2. 1 a 30'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['3. 31 a 60']=df_1_f['3. 31 a 60'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['4. 61 a 90']=df_1_f['4. 61 a 90'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['5. 91 a 120']=df_1_f['5. 91 a 120'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['6. 121 a 150']=df_1_f['6. 121 a 150'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['7. 151 a 210']=df_1_f['7. 151 a 210'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['8. 211 a 364']=df_1_f['8. 211 a 364'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['9. 365 a mas']=df_1_f['9. 365 a mas'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f['DEUDA_SOLES']=df_1_f['DEUDA_SOLES'].apply(lambda x:'{:,.0f}'.format(x))
df_1_f=df_1_f.apply(lambda x: x.astype(str).str.capitalize())



df_2_i = df_1.pivot_table(index=['GESTOR'], columns='TRAMO_VENCIMIENTO', values='RUC',aggfunc='nunique').reset_index().fillna(0)
df_2_t = df_1.pivot_table(index=['GESTOR'], values='RUC',aggfunc='nunique').reset_index().fillna(0)
df_2_f = pd.merge(df_2_i, df_2_t, on=['GESTOR'], how='outer').sort_values(by='RUC', ascending=False)
df_2_f['1. Por vencer']=df_2_f['1. Por vencer'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['2. 1 a 30']=df_2_f['2. 1 a 30'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['3. 31 a 60']=df_2_f['3. 31 a 60'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['4. 61 a 90']=df_2_f['4. 61 a 90'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['5. 91 a 120']=df_2_f['5. 91 a 120'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['6. 121 a 150']=df_2_f['6. 121 a 150'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['7. 151 a 210']=df_2_f['7. 151 a 210'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['8. 211 a 364']=df_2_f['8. 211 a 364'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['9. 365 a mas']=df_2_f['9. 365 a mas'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f['RUC']=df_2_f['RUC'].apply(lambda x:'{:,.0f}'.format(x))
df_2_f=df_2_f.apply(lambda x: x.astype(str).str.capitalize())


ruta_libro = "./src/models/"+año+""+mes+""+dia+"_Deuda_Carterizado_Gestión.xlsx"  # Reemplaza con la ruta y nombre de tu archivo Excel

# Definir el nombre de la carpeta
carpeta_año=f"RUTA\{año}"
carpeta_mes_año = f"RUTA\{año}\{fecha_actual.strftime('%m %B %Y').title()}"
carpeta_dia_fecha = f"RUTA\{año}\{fecha_actual.strftime('%m %B %Y').title()}\{fecha_actual.strftime('%A %d%m%Y').title()}"
ruta_final=f"RUTA\{año}\{fecha_actual.strftime('%m %B %Y').title()}\{fecha_actual.strftime('%A %d%m%Y').title()}\{año}{mes}{dia}_Deuda_Carterizado_Gestión.xlsx"

# Verificar si la carpeta existe
if not os.path.exists(carpeta_año):
    # Si no existe, crear la carpeta
    os.makedirs(carpeta_año)
else:
    logging.info(f"La carpeta '{carpeta_año}' ya existe.")
if not os.path.exists(carpeta_mes_año):
# Si no existe, crear la carpeta
    os.makedirs(carpeta_mes_año)
else:
    logging.info(f"La carpeta '{carpeta_mes_año}' ya existe.")
if not os.path.exists(carpeta_dia_fecha):
# Si no existe, crear la carpeta
    os.makedirs(carpeta_dia_fecha)
    shutil.copy(ruta_libro, carpeta_dia_fecha)
else:
    logging.info(f"La carpeta '{carpeta_dia_fecha}' ya existe.")
    shutil.copy(ruta_libro, carpeta_dia_fecha)
    


html=leer_html(ruta_html,df_1_f,df_2_f,fecha_usar,ruta_final)
Name_File_1 = "Reporte Deuda Carterizado al "+dia_1+"."+mes_1+"."+año_1+""  # Reemplaza con la ruta y nombre de tu archivo Excel


Eliminar_Excel(f'{Name_File_1}.html')
Eliminar_Excel(f'{Name_File_1}.pdf')

Name_File = "Reporte Deuda Carterizado al "+dia+"."+mes+"."+año+""  # Reemplaza con la ruta y nombre de tu archivo Excel

with open(f'./{Name_File}.html', 'w') as f:
    f.write(html)

import pdfkit
 
path_to='D:\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
path_file = f'{Name_File}.html'
config=pdfkit.configuration(wkhtmltopdf=path_to)
 
pdf_options = {
    'page-size': 'Letter',
    'margin-top': '0in',
    'margin-right': '0in',
    'margin-bottom': '0in',
    'margin-left': '0in'
}

pdfkit.from_file(path_file,output_path=f'./{Name_File}.pdf',options=pdf_options,configuration=config)


enviar_correo(html)
#Eliminar_Excel(ruta_libro)
close_connection_db(Conexion_Opercom)