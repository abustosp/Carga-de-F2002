import pandas as pd
import numpy as np
import os
import LIB.formatos as fmt
from tkinter.messagebox import showinfo
import openpyxl

# Leer el Excel "WP IVA IIBB.xlsx"
df = pd.read_excel('WP IVA IIBB.xlsx', engine='openpyxl' , sheet_name="Clientes")

# Filtrar los que no diga "SI" o "si" en la columna 'Importar'
df['Importar'] = df['Importar'].fillna('NO')
df = df[df['Importar'].str.contains('SI|si')]
#df['Ubicación IVA'] = df['Ubicación IVA'].str.replace("\\", "/", regex=True)
df['Archivo'] = df['Ubicación IVA'] + df['Archivo IVA'] + '.xlsx'

# Por cada linea del Excel, abrir los excels resultantes de la concatenación entre 'Ubicación IVA' y 'Archivo IVA' '.xlsx'
for i in range(len(df)):
    archivo = df['Archivo'].iloc[i]

    # si el archivo no existe continuar con el siguiente
    if not os.path.isfile(archivo):
        continue

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
    # Reemplazar 'Compra de bienes en el mercado local' por 'Compras de bienes (excepto bienes de uso)'
    Df_T_Compras['Compras por Agrupación de Crédito Fiscal'] = Df_T_Compras['Compras por Agrupación de Crédito Fiscal'].replace('Compra de bienes en el mercado local' , 'Compras de bienes (excepto bienes de uso)')
    Df_T_Compras['Compras por Agrupación de Crédito Fiscal'] = Df_T_Compras['Compras por Agrupación de Crédito Fiscal'].replace('Otros conceptos' , 'Otros Conceptos')
    Df_T_Compras['Compras por Agrupación de Crédito Fiscal'] = Df_T_Compras['Compras por Agrupación de Crédito Fiscal'].replace('Inversiones en bienes de uso' , 'Inversiones de Bienes de Uso')
    # Si 'Compras por Agrupación de Crédito Fiscal' es igual a 'Otros Conceptos' reemplazar 'Tasa IVA' por 'Consolidado'
    Df_T_Compras.loc[Df_T_Compras['Compras por Agrupación de Crédito Fiscal'] == 'Otros Conceptos', 'Tasa IVA'] = 'Consolidado'
    # Agrupar las compras por 'Compras por Agrupación de Crédito Fiscal' y 'Tasa IVA'
    Df_T_Compras = Df_T_Compras.groupby(['Compras por Agrupación de Crédito Fiscal' , 'Tasa IVA']).sum().reset_index()


    # Arreglar NCCompras
    Df_T_NCCompras['N. Créd. Recibidas - Crédito Fiscal a restituir'].fillna('Compra de bienes en el mercado local' , inplace=True)
    # Eliminar las columnas que no se usan que contienen 'Url'
    Df_T_NCCompras.drop(Df_T_NCCompras.filter(regex='Url').columns, axis=1, inplace=True)
    # Agrupar las compras por 'Compras por Agrupación de Crédito Fiscal' y 'Tasa IVA'
    Df_T_NCCompras = Df_T_NCCompras.groupby(['N. Créd. Recibidas - Crédito Fiscal a restituir' , 'Tasa IVA' , 'Column-1']).sum(numeric_only=False).reset_index()
    # Reemplazar 'Compra de bienes en el mercado local' por 'Compras de bienes (excepto bienes de uso)'
    Df_T_NCCompras['Column-1'] = Df_T_NCCompras['Column-1'].replace('Compra de bienes en el mercado local' , 'compras de bienes en el mercado local (excepto bienes de uso)')

    # Reemplazos de operaciones con...
    # los valores que estan entre parentesis, 'Cons. Finales' y 'Operaciones gravadas al 0%'
    reemplazosOperaciones = {
        r"\(.*\)" : "" ,
        'Cons. Finales, Exentos y No Alcanzados' : 'Consumidores finales, Exentos y No alcanzados' ,
        'Operaciones gravadas al 0%' : 'Operaciones no gravadas y exentas',
        'Operaciones no gravadas y exentas excepto exportaciones' : 'Operaciones no gravadas y exentas',
        }
    Texto_a_reemplazar = ' **** SIN CAE ni PEM ****'


    # Arreglar Ventas
    # Realizar los reemplazos de 'operaciones con... '
    Df_T_Ventas['Operaciones con...'] = Df_T_Ventas['Operaciones con...'].replace(reemplazosOperaciones , regex=True)
    # Reemplazar el texto ' **** SIN CAE ni PEM ****' por ''
    Df_T_Ventas['Operaciones con...'] = Df_T_Ventas['Operaciones con...'].str.replace(Texto_a_reemplazar , '' , regex=False)
    
    # Eliminar de la columna 'Operaciones con...' todo lo anteior al primer punto
    Df_T_Ventas['Operaciones con...'] = Df_T_Ventas['Operaciones con...'].str.split('.').str[1].str.strip()    
    # Eliminar las columnas que no se usan que contienen 'Url'
    Df_T_Ventas.drop(Df_T_Ventas.filter(regex='Url').columns, axis=1, inplace=True)
    
    # # sumar las filas cuyas 'Operaciones con...' son 'Operaciones no gravadas y exentas' y 'Operaciones no gravadas y exentas NC' en una sola fila con el concepto 'Operaciones no gravadas y exentas'
    # si el largo del dataframe es 0, continuar con el siguiente
    if len(Df_T_Ventas) > 0:
        Df_T_Ventas['Operaciones con...' ] = Df_T_Ventas['Operaciones con...'].replace('Operaciones no gravadas y exentas NC' , 'Operaciones no gravadas y exentas')
        # Agregar los montos y reemplazar en el DataFrame original
        Df_T_Ventas = Df_T_Ventas.groupby(['Ventas x Cód. Actividad', 'Operaciones con...' , 'Tasa IVA'])[['Monto Neto', 'Monto IVA', 'Monto Total', 'Copiar F2002']].sum().reset_index()

    
    # Arreglar NCVentas
    # Realizar los reemplazos de 'operaciones con... '
    Df_T_NCVentas['Operaciones con...'] = Df_T_NCVentas['Operaciones con...'].replace(reemplazosOperaciones , regex=True)
    Df_T_NCVentas['Operaciones con...'] = Df_T_NCVentas['Operaciones con...'].str.replace(Texto_a_reemplazar , '' , regex=False)
    # reemplazar 'Consumidores finales, Exentos y No alcanzados' y 'Monotributistas' por 'Sujetos Exentos, No Alcanzados, Monotributistas y Consumidores Finales'
    reemplazosncv ={
        'Consumidores finales, Exentos y No alcanzados' : 'Sujetos Exentos, No Alcanzados, Monotributistas y Consumidores Finales',
        'Monotributistas' : 'Sujetos Exentos, No Alcanzados, Monotributistas y Consumidores Finales'
    } 
    Df_T_NCVentas['Operaciones con...'] = Df_T_NCVentas['Operaciones con...'].replace(reemplazosncv , regex=True)
    # Eliminar de la columna 'Operaciones con...' todo lo anteior al primer punto
    Df_T_NCVentas['Operaciones con...'] = Df_T_NCVentas['Operaciones con...'].str.split('.').str[1].str.strip()
    # Eliminar las columnas que no se usan que contienen 'Url'
    Df_T_NCVentas.drop(Df_T_NCVentas.filter(regex='Url').columns, axis=1, inplace=True)
    # Sumar las operaciones que tengan el mismo concepto y tasa de IVA
    Df_T_NCVentas = Df_T_NCVentas.groupby(['Operaciones con...' , 'Tasa IVA']).sum(numeric_only=False).reset_index()


    # Listar los Dataframes
    Dataframes = [Df_T_Compras, Df_T_NCCompras, Df_T_Ventas, Df_T_NCVentas]

    # Reemplazos de valores
    reemplazos = {
        0.00: 'I.V.A. al 0,00%',
        0.21: 'I.V.A. al 21,00%',
        0.105: 'I.V.A. al 10,50%',
        0.27: 'I.V.A. al 27,00%',
        0.05: 'I.V.A. al 5,00%',
        0.025: 'I.V.A. al 2,50%'}

    # Por cada dataframe en la lista 'Dataframes'
    for dataframe in Dataframes:
        dataframe['Tasa IVA'] = dataframe['Tasa IVA'].replace(reemplazos)

    OutputFile = df["Ubicación IVA"].iloc[i] + "Procesado/" + df['Archivo IVA'].iloc[i] + '.xlsx'

    Directorio = df["Ubicación IVA"].iloc[i] + "Procesado/"

    # Si el directorio no existe, crearlo
    if not os.path.exists(Directorio):
        os.mkdir(df["Ubicación IVA"].iloc[i] + "Procesado/")

    # Exportación a Excel
    with pd.ExcelWriter(OutputFile , engine='openpyxl') as writer:
        Df_T_Compras.to_excel(writer, sheet_name='Compras', index=False)
        Df_T_NCCompras.to_excel(writer, sheet_name='NCCompras', index=False)
        Df_T_Ventas.to_excel(writer, sheet_name='Ventas', index=False)
        Df_T_NCVentas.to_excel(writer, sheet_name='NCVentas', index=False)

    # Aplicar los Formatos de Títulos y filtros
    workbook = openpyxl.load_workbook(OutputFile)
    h1 = workbook['Compras']
    h2 = workbook['NCCompras']
    h3 = workbook['Ventas']
    h4 = workbook['NCVentas']

    Hojas = [h1, h2, h3, h4]

    # Loop en todas las hojas
    for hoja in Hojas:
        fmt.Aplicar_formato_encabezado(hoja)
        # Autoajustar columnas
        fmt.Autoajustar_columnas(hoja)
        # Agregar Filtros
        fmt.Agregar_filtros(hoja)

    # Guardar el archivo
    workbook.save(OutputFile)

#Mostrar mensaje de finalización
showinfo("Finalizado", "Proceso finalizado con éxito")