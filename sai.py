# -*- coding: utf-8 -*-
"""
Created on Mon Aug 22 16:08:34 2022

@author: braviles
"""

import pandas as pd
import numpy as np

#%% Archivos y datos de entrada
archivo_sai = 'Junio-2022_CRDM_v2.xls'
archivo_population = 'población_provincia.xlsx'

#%% Procesamiento del archivo trimestral de Control
df_sai = pd.read_excel(archivo_sai)
df_sai.drop(['No'], inplace=True, axis=1)

prestadoras = np.unique(df_sai['PRESTADOR'])
print('No. Prestadoras', len(prestadoras))
provincias = np.unique(df_sai['PROVINCIA'])
df_provincial = pd.DataFrame(columns=(provincias))
# Mover la columna 'PRESTADOR' a la posición 0
df_provincial.insert(0, 'PRESTADOR', '')

# cálculo de conexiones por provincia y prestador
prestadores_provincia = dict()
for provincia in provincias:
    conteo = 0
    for i in range(len(prestadoras)):
        df_prestador = df_sai[(df_sai['PRESTADOR'] == prestadoras[i]) & (df_sai['PROVINCIA'] == provincia)]
        conexiones = np.sum(df_prestador['TOTAL CUENTAS'])
        df_provincial.loc[i, provincia] = conexiones
        df_provincial.loc[i, 'PRESTADOR'] = prestadoras[i]
        if not df_prestador.empty:
            conteo +=1
            prestadores_provincia[provincia] = conteo
        
df_provincial['TOTAL'] = df_provincial.iloc[:, 1:27].sum(axis=1)       

total_conexiones = df_provincial.loc[:, 'TOTAL'].sum(axis=0)
print(total_conexiones)

#porc_provincias = '%' + provincias

# import the population file
df_population = pd.read_excel(archivo_population)
dic_population = dict([(i, int(a)) for i, a in zip(df_population.PROVINCIA, df_population.POBLACIÓN)])
print(dic_population)

# Cálculo Cuota de mercado por provincia y prestador
for provincia in provincias:
    for i in range(len(prestadoras)):
        try:
            if df_provincial.loc[:, provincia].sum(axis=0) !=0:
                df_provincial.loc[i, '%'+ provincia] = df_provincial.loc[i, provincia] / df_provincial.loc[:, provincia].sum(axis=0)
        except:
            pass
# Cálculo de la densidad por provincia y prestador
for provincia in provincias:
    for i in range(len(prestadoras)):
        try:
            if dic_population[provincia] !=0:
                df_provincial.loc[i, 'Densidad_'+provincia] = (df_provincial.loc[i, provincia] / dic_population[provincia])*100
        except:
            pass

columnas = list(df_provincial.columns.values)

# Cálculo del IHH
IHH = []
total_provincia = []
densidad_provincia = []
for i in range(28, 53):
    IHH.append(np.sum(np.square(df_provincial.iloc[:,i]))*100**2)
for i in range(1, 28):
    total_provincia.append(np.sum(df_provincial.iloc[:, i]))
for i in range(53, len(columnas)):
    densidad_provincia.append(np.sum(df_provincial.iloc[:, i]))

fila_end = ['Totales, IHH, densidades']
fila_end.extend(total_provincia)
fila_end.extend(IHH)
fila_end.extend(densidad_provincia)
df_provincial.loc[len(prestadoras)] = fila_end

df_provincial['densidad_TOTAL'] = (df_provincial['TOTAL'] / df_population.loc[:, 'POBLACIÓN'].sum(axis=0))*100

# Data Frame Medio de TX
mediostx = np.unique(df_sai['TIPO ENLACE'])
df_mediostx = pd.DataFrame(columns=mediostx)
for medio in mediostx:
    for i in range(len(provincias)):
        df_medio = df_sai[(df_sai['TIPO ENLACE'] == medio) & (df_sai['PROVINCIA'] == provincias[i])]
        conexiones = np.sum(df_medio['TOTAL CUENTAS'])
        df_mediostx.loc[i, medio] = conexiones
        df_mediostx.loc[i, 'PROVINCIA'] = provincias[i]

#%% Export
# prestadores_provincia
# df_provincial
# df_mediostx
def exportar(df, nombre):
    df.to_excel(nombre, index=True, header=True)

exportar(df_provincial, 'final_provincial.xlsx')
df_prestadores_provincia = pd.DataFrame.from_dict(prestadores_provincia, orient ='index')
exportar(df_prestadores_provincia, 'prestadores_provincia.xlsx')
exportar(df_mediostx, 'final_mediostx.xlsx')
