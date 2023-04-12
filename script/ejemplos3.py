import docx
import pandas as pd
import numpy as np
import nums_from_string
import os #Este paquete permite crear actividades dependientes del sistema operativo, por ejemplo crear carpetas, conocer sobre un proceso, finalizar un proceso, etc.
import getpass
import glob
import matplotlib.pyplot as plt
from datetime import datetime
#from pyprojroot import here
import pyodbc
from janitor import clean_names # pip install pyjanitor
from pathlib import Path
from docx.shared import Pt
from docx.shared import Inches
import re #Nos porporciona opciones de coincidencia.
from docx.shared import Cm # para incluir imagenes en el documento Word
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
import matplotlib.ticker as mtick
from docxtpl import DocxTemplate
from docxtpl import InlineImage
import win32com.client as win32
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ruta= r"C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GESTION DE LA INFORMACIÓN\UPP\Am Automatizada v2\AM_Automatizada\dataset"

## Cargamos nombres de regiones
nombre_regiones = pd.read_excel(r"C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GESTION DE LA INFORMACIÓN\UPP\Am Automatizada v2\AM_Automatizada\dataset\nombre_regiones.xlsx")

data_cas_nexus = pd.read_excel(proyecto / r"C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GESTION DE LA INFORMACIÓN\UPP\Am Automatizada v2\AM_Automatizada\dataset/bd_plazas_nexus/ReporteNexusCas_IAP_30032023.xlsx", sheet_name='Reporte_PEAS_2022')   
data_cas_nexus = clean_names(data_cas_nexus) # Normalizamos nombres

data_cas_anexo2=data_cas_nexus[['cod_pliego','peas_programadas','contratado','cod_tipo_cargo','cargo','cod_int','intervencion_nombre_corto']].groupby(by=['cod_pliego','cod_int','intervencion_nombre_corto','cod_tipo_cargo','cargo'], as_index=False).sum()
data_cas_anexo2=data_cas_anexo2.merge(right=nombre_regiones,how="left", on='cod_pliego')

#Amazonas
region_seleccionada = data_cas_anexo2['region'] == 'AMAZONAS' #Seleccionar region
tabla_intervenciones = data_cas_anexo2[region_seleccionada]   
   
#Ejecución
tabla_intervenciones['ejecucion']= tabla_intervenciones["contratado"]/tabla_intervenciones["peas_programadas"] 
tabla_intervenciones['ejecucion'] = tabla_intervenciones['ejecucion'].replace(np.inf, 0)
tabla_intervenciones['ejecucion'] = tabla_intervenciones['ejecucion'].fillna(0)

ejec_2=tabla_intervenciones["contratado"].sum()/tabla_intervenciones["peas_programadas"].sum()  if tabla_intervenciones["contratado"].sum()/tabla_intervenciones["peas_programadas"].sum()>0 else 0

# TOTAL
tabla_intervenciones_formato_10 = tabla_intervenciones #data_cas_2[region_seleccionada]
tabla_intervenciones_formato_10 = tabla_intervenciones.sort_values(by=['intervencion_nombre_corto'])

categorias = tabla_intervenciones_formato_10['intervencion_nombre_corto'].unique()

# Tabla con los subtotales
df_list = []
for i in range(0,len(categorias)):
    filas_categoria = tabla_intervenciones_formato_10.loc[tabla_intervenciones_formato_10['intervencion_nombre_corto'] == categorias[i]]
    total_int = filas_categoria.groupby(by = ["region"], as_index=False).sum()
    tabla_intervenciones_formato_11 = filas_categoria.append(total_int, ignore_index=True)
    
    ejec_3=tabla_intervenciones_formato_11["contratado"].sum()/tabla_intervenciones_formato_11["peas_programadas"].sum()  if tabla_intervenciones_formato_11["contratado"].sum()/tabla_intervenciones_formato_11["peas_programadas"].sum()>0 else 0

    tabla_intervenciones_formato_11['peas_programadas'] = tabla_intervenciones_formato_11['peas_programadas'].fillna("0").astype(float)
    tabla_intervenciones_formato_11.iloc[-1, tabla_intervenciones_formato_11.columns.get_loc('ejecucion')] = ejec_3
    #Incluimos palabra "total" y "-" en vez de NaN
    tabla_intervenciones_formato_11['intervencion_nombre_corto'] = tabla_intervenciones_formato_11['intervencion_nombre_corto'].fillna("Total")
    tabla_intervenciones_formato_11['cargo'] = tabla_intervenciones_formato_11['cargo'].fillna("")
    #tabla_intervenciones_formato_7['ejecucion'] = tabla_intervenciones_formato_7['ejecucion'].fillna("0").astype(float)
    tabla_intervenciones_formato_11=tabla_intervenciones_formato_11[['intervencion_nombre_corto','cargo','peas_programadas','contratado','ejecucion']]
    
    df_list.append(tabla_intervenciones_formato_11)
    
# Concatenar los DataFrames en uno solo
df_concatenado = pd.concat(df_list, ignore_index=True)
    


'''







# Generamos fila total
total_int = tabla_intervenciones_formato_10.groupby(by = ["region"], as_index=False).sum()

# Realizamos append del total en la tabla
tabla_intervenciones_formato_10 = tabla_intervenciones_formato_10.append(total_int, ignore_index=True)

# Reemplazamos % de avance correctos en fila total
#tabla_intervenciones_formato_10['costo_upp_total'] = tabla_intervenciones_formato_9['costo_upp_total'].fillna("0").astype(float)
tabla_intervenciones_formato_10['peas_programadas'] = tabla_intervenciones_formato_10['peas_programadas'].fillna("0").astype(float)
tabla_intervenciones_formato_10.iloc[-1, tabla_intervenciones_formato_10.columns.get_loc('ejecucion')] = ejec_2
#Incluimos palabra "total" y "-" en vez de NaN
tabla_intervenciones_formato_10['intervencion_nombre_corto'] = tabla_intervenciones_formato_10['intervencion_nombre_corto'].fillna("Total")
tabla_intervenciones_formato_10['cargo'] = tabla_intervenciones_formato_10['cargo'].fillna("")
#tabla_intervenciones_formato_7['ejecucion'] = tabla_intervenciones_formato_7['ejecucion'].fillna("0").astype(float)

# Formato para la tabla
formato_tabla_intervenciones = {
"intervencion_nombre_corto" : "{}",
"cargo" : "{}",
"peas_programadas" : "{:,.0f}",
"contratado": "{:,.0f}",
"ejecucion": "{:,.1%}"
#    "costo_actual": "{:,.0f}",
#    "ejecucion": "{:,.1%}",
}

tabla_intervenciones_formato_10 = tabla_intervenciones_formato_10.transform({k: v.format for k, v in formato_tabla_intervenciones.items()})
