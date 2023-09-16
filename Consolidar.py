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
df['Ubicación IVA'] = df['Ubicación IVA'].fillna('NO').str.replace("\\", "/", regex=True)
df['Archivo'] = df['Ubicación IVA'] + "Procesado/" + df['Archivo IVA'] + '.xlsx'

Ventas = pd.DataFrame()
Compras = pd.DataFrame()
NCVentas = pd.DataFrame()
NCCompras = pd.DataFrame()

# Por cada linea del Excel, abrir los excels resultantes de la concatenación entre 'Ubicación IVA' y 'Archivo IVA' '.xlsx'
for i in range(len(df)):
    archivo = df['Archivo'].iloc[i]
    ubicacion = df['Ubicación IVA'].iloc[i] + "Procesado/"
    nombreArchivo = df['Archivo IVA'].iloc[i]

    # Si el archivo no existe continuar con el siguiente
    if not os.path.exists(archivo):
        continue

    Df_T_Compras = pd.read_excel(archivo, sheet_name='Compras')
    Df_T_NCCompras = pd.read_excel(archivo, sheet_name='NCCompras')
    Df_T_Ventas = pd.read_excel(archivo, sheet_name='Ventas')
    Df_T_NCVentas = pd.read_excel(archivo, sheet_name='NCVentas')

    # Agregar en todos los DataFrames una columna con la ubicación del archivo
    Df_T_Compras['Ubicación'] = ubicacion
    Df_T_NCCompras['Ubicación'] = ubicacion
    Df_T_Ventas['Ubicación'] = ubicacion
    Df_T_NCVentas['Ubicación'] = ubicacion

    # Agregar a todos los DataFrames una columna con el nombre del archivo
    Df_T_Compras['Archivo'] = nombreArchivo
    Df_T_NCCompras['Archivo'] = nombreArchivo
    Df_T_Ventas['Archivo'] = nombreArchivo
    Df_T_NCVentas['Archivo'] = nombreArchivo

    # Concatenar los DataFrames
    Compras = pd.concat([Compras, Df_T_Compras], ignore_index=True)
    NCCompras = pd.concat([NCCompras, Df_T_NCCompras], ignore_index=True)
    Ventas = pd.concat([Ventas, Df_T_Ventas], ignore_index=True)
    NCVentas = pd.concat([NCVentas, Df_T_NCVentas], ignore_index=True)

# Exportar los DataFrames a Excel
with pd.ExcelWriter('WP IVA Consolidado.xlsx') as writer:
    Compras.to_excel(writer, sheet_name='Compras', index=False)
    NCCompras.to_excel(writer, sheet_name='NCCompras', index=False)
    Ventas.to_excel(writer, sheet_name='Ventas', index=False)
    NCVentas.to_excel(writer, sheet_name='NCVentas', index=False)

# Aplicar formato de Titulos y Filtros
workbook = openpyxl.load_workbook('WP IVA Consolidado.xlsx')
h1 = workbook['Compras']
h2 = workbook['NCCompras']
h3 = workbook['Ventas']
h4 = workbook['NCVentas']

Hojas = [h1, h2, h3, h4]

for hoja in Hojas:
    fmt.Aplicar_formato_encabezado(hoja)
    fmt.Agregar_filtros(hoja)
    fmt.Autoajustar_columnas(hoja)

workbook.save('WP IVA Consolidado.xlsx')

showinfo("Consolidar", "Se consolidaron los archivos en 'WP IVA Consolidado.xlsx'")


