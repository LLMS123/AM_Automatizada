#%reset -f

import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from janitor import clean_names # pip install pyjanitor
from docx.shared import Cm # para incluir imagenes en el documento Word
#lectura
#

fecha_corte_disponibilidad = "20220215"
nyear_disponibilidad=fecha_corte_disponibilidad[0:4]
nmes_disponibilidad=fecha_corte_disponibilidad[4:6]
ndia_disponibilidad=fecha_corte_disponibilidad[6:]
fecha_corte_disponibilidad_format=ndia_disponibilidad + "/" + nmes_disponibilidad + "/" + nyear_disponibilidad

df1 = pd.read_excel (
     os.path.join(r"C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GITHUB\Graficos_180",f"Disponibilidad_Presupuestal_{fecha_corte_disponibilidad}.xlsx"),
     engine='openpyxl',
)

df2 = df1.clean_names()

#filtrado por región
filtro_region=df2['region']=="CUSCO"
base_region=df2[filtro_region]

base_filtrada=base_region[['unidad_ejecutora',"pia",f"pim_reporte_siaf_{fecha_corte_disponibilidad}",f"devengado_reporte_siaf_{fecha_corte_disponibilidad}"]]

pia=[base_filtrada["pia"].sum()]
pim=[base_filtrada[f"pim_reporte_siaf_{fecha_corte_disponibilidad}"].sum()]
devengado=[base_filtrada[f"devengado_reporte_siaf_{fecha_corte_disponibilidad}"].sum()]
fecha=[fecha_corte_disponibilidad_format]


df_collaps=base_filtrada[['pia','unidad_ejecutora']].groupby(by="unidad_ejecutora", as_index=False).sum()

df_sorted = df_collaps.sort_values('pia', ascending=False)

plt.barh(df_sorted["unidad_ejecutora"], df_sorted["pia"],color=['#C00000'])

for i, v in enumerate(df_sorted['pia']):
    plt.text(v + 2, i - 0.1, str(v), color='black', fontsize=12, fontweight='normal', fontstyle='normal', fontfamily='sans-serif')

#plt.axvline(x=31357871, color='blue', linestyle='--', linewidth=1)
#plt.axvline(x=31367871, color='green', linestyle='--', linewidth=1)

plt.grid(axis='x', color='gray', linestyle=':', linewidth=1, zorder=-1000)

plt.show()

# Agregar títulos y etiquetas de los ejes
#plt.title('Gráfico de barras invertidas')
#plt.xlabel('Valores')
#plt.ylabel('Categorias')





import pandas as pd
import matplotlib.pyplot as plt

# Crear el DataFrame
df = pd.DataFrame({
    'Categorias': ['A', 'B', 'C', 'D'],
    'Valores': [25.3, 50.7, 75.2, 100.9]
})

# Ordenar el DataFrame por los valores de manera descendente
df_sorted = df.sort_values('Valores', ascending=False)

# Crear el gráfico de barras invertidas con las barras de color rojo
plt.barh(df_sorted['Categorias'], df_sorted['Valores'], color='red')

# Agregar títulos y etiquetas de los ejes
plt.title('Gráfico de barras invertidas')
plt.xlabel('Valores')
plt.ylabel('Categorias')

# Colocar las etiquetas al costado de las barras con formato adecuado
for i, v in enumerate(df_sorted['Valores']):
    plt.text(v + 2, i - 0.1, '{:.1f}%'.format(v/df_sorted['Valores'].sum()*100), color='black', fontsize=12, fontweight='normal', fontstyle='normal', fontfamily='sans-serif')

# Agregar líneas verticales
plt.axvline(x=50, color='blue', linestyle='--', linewidth=1)
plt.axvline(x=75, color='green', linestyle='--', linewidth=1)

# Agregar líneas de cuadrícula en el eje x detrás de las barras del gráfico
plt.grid(axis='x', color='gray', linestyle=':', linewidth=1, zorder=-1)

# Mostrar el gráfico
plt.show()


picPath = r'C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GESTION DE LA INFORMACIÓN\UPP\Am Automatizada v2\AM_Automatizada\graficos\Gráfico_1.png'
plt.savefig(picPath)

plt.show()







import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# Crear el DataFrame
df = pd.DataFrame({
    'Categorias': ['A', 'B', 'C', 'D'],
    'Valores': [25.3, 50.7, 75.2, 100.9]
})

# Ordenar el DataFrame por los valores de manera descendente
df_sorted = df.sort_values('Valores', ascending=False)

# Crear el gráfico de barras invertidas con las barras de color rojo
plt.barh(df_sorted['Categorias'], df_sorted['Valores'], color='red', zorder=2)

# Agregar títulos y etiquetas de los ejes
plt.title('Gráfico de barras invertidas')
plt.xlabel('Valores')
plt.ylabel('Categorias')

# Colocar las etiquetas al costado de las barras con formato adecuado
for i, v in enumerate(df_sorted['Valores']):
    plt.text(v + 2, i - 0.1, '{:.1f}%'.format(v), color='black', fontsize=12, fontweight='normal', fontstyle='normal', fontfamily='sans-serif')

# Agregar líneas verticales
plt.axvline(x=50, color='blue', linestyle='--', linewidth=1, zorder=1)
plt.axvline(x=75, color='green', linestyle='--', linewidth=1, zorder=1)

# Agregar líneas de cuadrícula en el eje x detrás de las barras del gráfico
plt.grid(axis='x', color='gray', linestyle=':', linewidth=1, zorder=0)

# Mostrar el eje x en formato porcentaje
fmt = '%.0f%%'
xticks = mtick.StrMethodFormatter(fmt)
plt.gca().xaxis.set_major_formatter(xticks)

# Mostrar el gráfico
plt.show()









#etiquetas=['PIA','PIM','Devengado']

etiquetas=[f'Corte {fecha_corte_disponibilidad_format}']
valores=pia
valores2=pim
valores3=devengado

co=np.arange(len(valores))  #
an=0.30

fig, ax=plt.subplots()

g1=ax.bar(co-an/2,valores,an-0.20,color='#B22600',label='PIA')
g2=ax.bar(co+an-0.35,valores2,an-0.20,color='#C27B6E',label='PIM')
g3=ax.bar(co+an-0.25,valores3,an-0.20,color='#A6A6A6',label='Devengado')

#g1=ax.bar(co-an/2,valores,an-0.15,color='#B22600',label='PIA')
#g2=ax.bar(co+an/2,valores2,an+0.15,color='#C27B6E',label='PIM')
#g3=ax.bar(co+1.5*an,valores3,an-0.15,color='#A6A6A6',label='Devengado')

'''
for i,j in zip(co,valores):
    ax.annotate(j, xy=(i-0.2,j+100000))

for i,j in zip(co,valores2):
    ax.annotate(j, xy=(i+0.1,j+100000))
    
for i,j in zip(co,valores3):
    ax.annotate(j, xy=(i+ 0.4,j+100000))  
'''

#definimos los ejes
lista_y=pia + pim + devengado #colocar todos los valores de las variables
lista_y_max=max(lista_y) #colocar todos los valores de las variables
lista_y_num=np.array(lista_y_max)
lista_y_num_mill=lista_y_num/1000000
lista_y_limite_sup=round(lista_y_num_mill.max()+5,1) #se suma 0.05 para redondear

eje_y_interval=round(lista_y_limite_sup/3,1)

eje_y_limite_inf=0
eje_y_limite_2=eje_y_limite_inf+eje_y_interval
eje_y_limite_3=eje_y_limite_2+eje_y_interval

axtick=[int(eje_y_limite_inf*1000000),int(eje_y_limite_2*1000000),int(eje_y_limite_3*1000000),int(lista_y_limite_sup*1000000)]
#axtick_entero=[int(eje_y_limite_inf*1000000),int(eje_y_limite_2*1000000),int(eje_y_limite_3*1000000),int(lista_y_limite_sup*1000000)]
axtick_decimal=[eje_y_limite_inf,eje_y_limite_2,eje_y_limite_3,lista_y_limite_sup]

y_values=[]
for word in axtick_decimal:
     string_eje=str(word)+" mill."
     y_values.append(string_eje)

Axes=plt.gca() #Para quitar las las lineas del recuadro donde esta el gráfico
Axes.spines['top'].set_visible(False)
Axes.spines['right'].set_visible(False)
Axes.spines['bottom'].set_linestyle('dashed') 
Axes.spines['left'].set_linestyle('dashed') 
#plt.ylabel('Montos en S/ (Millones)')
plt.yticks(axtick, y_values)

#ax.set_title('Información SIAF')
ax.set_ylabel('Montos (en millones de S/)')
#ax.set_xticks(co)
#ax.set_xticklabels(etiquetas)
#plt.legend()
plt.yticks(axtick, y_values)

ax.set_yticks(axtick)
ax.set_ylim([0,int(lista_y_limite_sup*1000000)])

x=0
for p in g1:
    valores_1=str(round(valores[x]/1000000,2))+" mill."
    height = p.get_height()

    ax.annotate('{}'.format(valores_1),
      xy=(p.get_x() + p.get_width() / 2, height),
      xytext=(0, 3), # 3 points vertical offset
      textcoords="offset points",
      ha='center', va='bottom')

x=0
for p in g2:
    valores_2=str(round(valores2[x]/1000000,2))+" mill."
    height = p.get_height()

    ax.annotate('{}'.format(valores_2),
      xy=(p.get_x() + p.get_width() / 2, height),
      xytext=(0, 3), # 3 points vertical offset
      textcoords="offset points",
      ha='center', va='bottom')

x=0
for p in g3:
    valores_3=str(round(valores3[x]/1000000,2))+" mill."
    height = p.get_height()

    ax.annotate('{}'.format(valores_3),
      xy=(p.get_x() + p.get_width() / 2, height),
      xytext=(0, 3), # 3 points vertical offset
      textcoords="offset points",
      ha='center', va='bottom')


#ax.legend(handles=[g1], loc='upper left')

#new_array = np.array(devengado, dtype=int)
#new_array2 = np.array(PIM, dtype=int)

#plt.arrow(0,0,1,1)
#ax2=plt.subplots(ncols=1, sharey=True)
#plt.xticks(rotation=45)
#ax2.set_box_aspect(1)
#plt.twinx()
ax.set_xticks(co)
ax.set_xticklabels(etiquetas)
plt.legend()

picPath = r'C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GESTION DE LA INFORMACIÓN\UPP\Am Automatizada v2\AM_Automatizada\graficos\Gráfico_1.png'
plt.savefig(picPath)

plt.show()



'''


valores=pia + pim + devengado
#------------------------------------------------







#------------------------------------------------






variables_SIAF
Valores_SIAF



#definimos los ejes
lista_y=pia + pim + devengado #colocar todos los valores de las variables
lista_y_max=max(lista_y) #colocar todos los valores de las variables
lista_y_num=np.array(lista_y_max)
lista_y_num_mill=lista_y_num/1000000
lista_y_limite_sup=round(lista_y_num_mill.max()+1,1) #se suma 0.05 para redondear

eje_y_interval=round(lista_y_limite_sup/3,1)

eje_y_limite_inf=0
eje_y_limite_2=eje_y_limite_inf+eje_y_interval
eje_y_limite_3=eje_y_limite_2+eje_y_interval

axtick=[int(eje_y_limite_inf*1000000),int(eje_y_limite_2*1000000),int(eje_y_limite_3*1000000),int(lista_y_limite_sup*1000000)]
#axtick_entero=[int(eje_y_limite_inf*1000000),int(eje_y_limite_2*1000000),int(eje_y_limite_3*1000000),int(lista_y_limite_sup*1000000)]
axtick_decimal=[eje_y_limite_inf,eje_y_limite_2,eje_y_limite_3,lista_y_limite_sup]

y_values=[]
for word in axtick_decimal:
     string_eje=str(word)+" mill."
     y_values.append(string_eje)


fig, ax2 = plt.subplots()
ppss=ax2.bar(variables_SIAF, Valores_SIAF,color=['green','red','black'], label=('PIA','PIM','Devengado'))
ax2.set_title('Gráfico información SIAf')
Axes=plt.gca()
Axes.spines['top'].set_visible(False)
Axes.spines['right'].set_visible(False)
Axes.spines['bottom'].set_visible(False)
Axes.spines['left'].set_visible(False)
plt.ylabel('Montos en S/ (Millones)')
plt.yticks(axtick, y_values)

#-----------------------



#definimos los ejes
lista_y=pia + pim + devengado #colocar todos los valores de las variables
lista_y_max=max(lista_y) #colocar todos los valores de las variables
lista_y_num=np.array(lista_y_max)
lista_y_num_mill=lista_y_num/1000000
lista_y_limite_sup=round(lista_y_num_mill.max()+0.05,1) #se suma 0.05 para redondear

eje_y_interval=round(lista_y_limite_sup/3,1)

eje_y_limite_inf=0
eje_y_limite_2=eje_y_limite_inf+eje_y_interval
eje_y_limite_3=eje_y_limite_2+eje_y_interval

axtick=[int(eje_y_limite_inf*1000000),int(eje_y_limite_2*1000000),int(eje_y_limite_3*1000000),int(lista_y_limite_sup*1000000)]
#axtick_entero=[int(eje_y_limite_inf*1000000),int(eje_y_limite_2*1000000),int(eje_y_limite_3*1000000),int(lista_y_limite_sup*1000000)]
axtick_decimal=[eje_y_limite_inf,eje_y_limite_2,eje_y_limite_3,lista_y_limite_sup]


ax = df2.plot.bar(rot=0)

y_values=[]
for word in axtick_decimal:
     string_eje=str(word)+" mill."
     y_values.append(string_eje)

#creación del gráfico
fig, ax=plt.subplots()
ppss=ax.bar(mes_fin, graficar,color=['green','red','black'], label=('PIA','PIM','Devengado'))
Axes=plt.gca()
Axes.spines['top'].set_visible(False)
Axes.spines['right'].set_visible(False)
Axes.spines['bottom'].set_visible(False)
Axes.spines['left'].set_visible(False)
plt.ylabel('Devengado')
plt.yticks(axtick, y_values)

ax.set_yticks(axtick)
ax.set_ylim([0,int(lista_y_limite_sup*1000000)])


x=0
for p in ppss:
    graficars=str(round(graficar[x]/1000000,2))+" mill."
    height = p.get_height()

    ax.annotate('{}'.format(graficars),
      xy=(p.get_x() + p.get_width() / 2, height1),
      xytext=(0, 3), # 3 points vertical offset
      textcoords="offset points",
      ha='center', va='bottom')

ax.legend(handles=[ppss], loc='upper left')
new_array = np.array(devengado, dtype=int)
new_array2 = np.array(PIM, dtype=int)

#plt.arrow(0,0,1,1)
#ax2=plt.subplots(ncols=1, sharey=True)
plt.xticks(rotation=45)
#ax2.set_box_aspect(1)
#plt.twinx()
plt.show()


ppss2=ax.bar(mes_fin, pim,color='red', label='Devengado')
Axes=plt.gca()
Axes.spines['top'].set_visible(False)
Axes.spines['right'].set_visible(False)
Axes.spines['bottom'].set_visible(False)
Axes.spines['left'].set_visible(False)
plt.ylabel('PIM')
plt.yticks(axtick, y_values)

ax.set_yticks(axtick)
ax.set_ylim([0,int(lista_y_limite_sup*1000000)])
#creación del gráfico
x=0
for p in ppss2:
    pims=str(round(pim[x]/1000000,2))+" mill."
    height = p.get_height()
    x=x+1
    
    ax.annotate('{}'.format(pims),
      xy=(p.get_x() + p.get_width() / 2, height),
      xytext=(0, 3), # 3 points vertical offset
      textcoords="offset points",
      ha='center', va='bottom')
   
ax.legend(handles=[ppss], loc='upper left')
new_array = np.array(devengado, dtype=int)
new_array2 = np.array(PIM, dtype=int)

#plt.arrow(0,0,1,1)

#ax2=plt.subplots(ncols=1, sharey=True)
plt.xticks(rotation=45)

plt.legend()
plt.show()










Galaxy= plt.scatter(mes_fin, costo,  color='blue', s=150, label='Costo')
x=0
for word in costo:
    word=str(round(costo[x]/1000000,2))+" mill."
    plt.text(mes_fin[x],costo[x] , word, horizontalalignment='center',
     verticalalignment='bottom')
    x=x+1
    
plt.ylim([0,int(lista_y_limite_sup*1000000)])




Axes=plt.gca()
Axes.spines['top'].set_visible(False)
Axes.spines['right'].set_visible(False)
Axes.spines['bottom'].set_visible(False)
Axes.spines['left'].set_visible(False)

Axes.set_yticks([])

plt.legend()
plt.show()

'''
