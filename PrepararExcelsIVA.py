import pandas as pd
import numpy as np
import os
import LIB.formatos as fmt
from tkinter.messagebox import showinfo

# Leer el Excel "WP IVA IIBB.xlsx"
df = pd.read_excel('WP IVA IIBB.xlsx', engine='openpyxl' , sheet_name="Clientes")

# Filtrar los que no diga "SI" o "si" en la columna 'Importar'
df['Importar'] = df['Importar'].fillna('NO')
df = df[df['Importar'].str.contains('SI|si')]
df['Ubicación IVA'] = df['Ubicación IVA'].fillna('NO').str.replace("\\", "/", regex=True)
df['Archivo'] = df['Ubicación IVA'] + df['Archivo IVA'] + '.xlsx'

# Por cada linea del Excel, abrir los excels resultantes de la concatenación entre 'Ubicación IVA' y 'Archivo IVA' '.xlsx'
for i in range(len(df)):
    archivo = df['Archivo'].iloc[i]
    Df_T_Compras = pd.read_excel(archivo, engine='openpyxl', sheet_name='Compras')
    Df_T_NCCompras = pd.read_excel(archivo, engine='openpyxl', sheet_name='NCCompras')
    Df_T_Ventas = pd.read_excel(archivo, engine='openpyxl', sheet_name='Ventas')
    Df_T_NCVentas = pd.read_excel(archivo, engine='openpyxl', sheet_name='NCVentas')

    # Eliminnar la ultima fila de cada hoja
    Df_T_Compras = Df_T_Compras.iloc[:-1]
    Df_T_NCCompras = Df_T_NCCompras.iloc[:-1]
    Df_T_Ventas = Df_T_Ventas.iloc[:-1]
    Df_T_NCVentas = Df_T_NCVentas.iloc[:-1]

    # Arreglar Compras
    Df_T_Compras['Compras por Agrupación de Crédito Fiscal'].fillna('Compra de bienes en el mercado local' , inplace=True)
    # Eliminar las columnas que no se usan que contienen 'Url'
    Df_T_Compras.drop(Df_T_Compras.filter(regex='Url').columns, axis=1, inplace=True)
    # Agrupar las compras por 'Compras por Agrupación de Crédito Fiscal' y 'Tasa IVA'
    Df_T_Compras = Df_T_Compras.groupby(['Compras por Agrupación de Crédito Fiscal' , 'Tasa IVA']).sum().reset_index()
    # Reemplazar 'Compra de bienes en el mercado local' por 'Compras de bienes (excepto bienes de uso)'
    Df_T_Compras['Compras por Agrupación de Crédito Fiscal'] = Df_T_Compras['Compras por Agrupación de Crédito Fiscal'].replace('Compra de bienes en el mercado local' , 'Compras de bienes (excepto bienes de uso)')


    # Arreglar NCCompras
    Df_T_NCCompras['Notas de Créd. Recibidas'].fillna('Compra de bienes en el mercado local' , inplace=True)
    # Eliminar las columnas que no se usan que contienen 'Url'
    Df_T_NCCompras.drop(Df_T_NCCompras.filter(regex='Url').columns, axis=1, inplace=True)
    # Agrupar las compras por 'Compras por Agrupación de Crédito Fiscal' y 'Tasa IVA'
    Df_T_NCCompras = Df_T_NCCompras.groupby(['Notas de Créd. Recibidas' , 'Tasa IVA']).sum().reset_index()
    # Reemplazar 'Compra de bienes en el mercado local' por 'Compras de bienes (excepto bienes de uso)'
    Df_T_NCCompras['Notas de Créd. Recibidas'] = Df_T_NCCompras['Notas de Créd. Recibidas'].replace('Compra de bienes en el mercado local' , 'Compras de bienes (excepto bienes de uso)')

    # Reemplazos de operaciones con...
    # los valores que estan entre parentesis, 'Cons. Finales' y 'Operaciones gravadas al 0%'
    reemplazosOperaciones = {
        r"\(.*\)" : "" ,
        'Cons. Finales, Exentos y No Alcanzados' : 'Consumidores finales, Exentos y No alcanzados' ,
        'Operaciones gravadas al 0%' : 'Operaciones no gravadas y exentas' }


    # Arreglar Ventas
    # Realizar los reemplazos de 'operaciones con... '
    Df_T_Ventas['Operaciones con...'] = Df_T_Ventas['Operaciones con...'].replace(reemplazosOperaciones , regex=True)    
    # Eliminar de la columna 'Operaciones con...' todo lo anteior al primer punto
    Df_T_Ventas['Operaciones con...'] = Df_T_Ventas['Operaciones con...'].str.split('.').str[1].str.strip()    
    # Eliminar las columnas que no se usan que contienen 'Url'
    Df_T_Ventas.drop(Df_T_Ventas.filter(regex='Url').columns, axis=1, inplace=True)
    
    # Arreglar NCVentas
    # Realizar los reemplazos de 'operaciones con... '
    Df_T_NCVentas['Operaciones con...'] = Df_T_NCVentas['Operaciones con...'].replace(reemplazosOperaciones , regex=True)
    # Eliminar de la columna 'Operaciones con...' todo lo anteior al primer punto
    Df_T_NCVentas['Operaciones con...'] = Df_T_NCVentas['Operaciones con...'].str.split('.').str[1].str.strip()
    # Eliminar las columnas que no se usan que contienen 'Url'
    Df_T_NCVentas.drop(Df_T_NCVentas.filter(regex='Url').columns, axis=1, inplace=True)


    # Listar los Dataframes
    Dataframes = [Df_T_Compras, Df_T_NCCompras, Df_T_Ventas, Df_T_NCVentas]

    # Reemplazos de valores
    reemplazos = {
        0.21: 'I.V.A. al 21,00%',
        0.105: 'I.V.A. al 10,50%',
        0.27: 'I.V.A. al 27,00%',
        0.05: 'I.V.A. al 5,00%',
        0.025: 'I.V.A. al 2,50%'}

    # Por cada dataframe en la lista 'Dataframes'
    for dataframe in Dataframes:
        dataframe['Tasa IVA'] = dataframe['Tasa IVA'].replace(reemplazos)

    OutputFile = df["Archivo IVA 2002"].str.replace("\\" , "/" , regex=True).iloc[i]

    # Exportación a Excel
    with pd.ExcelWriter(OutputFile) as writer:
        Df_T_Compras.to_excel(writer, sheet_name='Compras', index=False)
        Df_T_NCCompras.to_excel(writer, sheet_name='NCCompras', index=False)
        Df_T_Ventas.to_excel(writer, sheet_name='Ventas', index=False)
        Df_T_NCVentas.to_excel(writer, sheet_name='NCVentas', index=False)

    

#Mostrar mensaje de finalización
showinfo("Finalizado", "Proceso finalizado con éxito")