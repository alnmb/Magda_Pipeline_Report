import pandas as pd
from datetime import datetime
import numpy as np
import os
#####################################
"""
Reporte de formateo de archivo .xls usando python
"""

"""
Version: 1.0
"""
######################################

archivo_fuente_ruta = "C:/Users/ALANMART/OneDrive - Capgemini/Desktop/test/"

def LeerArchivo(ruta):
    files = os.listdir(ruta)
    for file_name in files:
        if file_name.endswith('.xlsx'):
            df = pd.read_excel(ruta+file_name,index_col=None, dtype={'Opportunity ID':str})
            os.remove(ruta+file_name)
    print(df)
    return df

def Transformaciones(df):
    #borrar columnas Account y Customer Relationship
    columnas_eliminar = ['Account','Customer Relationship','Services Revenue Currency']
    df = df.drop(columnas_eliminar,axis=1)
    print(df.columns.to_list())
    #crear columna
    df.insert(0,'ECSD Month','')
    #renombrar columnas
    nuevos_nombres = {'Proposal Submission Date':'PSD','Close Date (Expected Signing or Date Won/Lost/WD)':'ECSD','Consulting Start Date':'Contract Start Date',
                      'Consulting End Date':'Contract End Date','Win Probability%':'Win Prob%','Services Revenue':'Revenue'}
    df = df.rename(columns=nuevos_nombres)
    print(df.columns.to_list())
    #cambiar formato de columas
    df['Revenue'] = df['Revenue'].astype(int)
    df['Revenue'] = df['Revenue'].apply(lambda x: "${:,.0f}".format(x))

    #añadir a las columnas de fechas formato datetime
    df['PSD'] = pd.to_datetime(df['PSD'])
    #df['PSD'] = df['PSD'].dt.strftime('%m/%d/%Y')
    
    df['ECSD'] = pd.to_datetime(df['ECSD'])
    #df['ECSD'] = df['ECSD'].dt.strftime('%m/%d/%Y')
    
    df['Contract Start Date'] = pd.to_datetime(df['Contract Start Date'])
    #df['Contract Start Date'] = df['Contract Start Date'].dt.strftime('%m/%d/%Y')

    df['Contract End Date'] = pd.to_datetime(df['Contract End Date'])
    #df['Contract End Date'] = df['Contract End Date'].dt.strftime('%m/%d/%Y')
    
    df['Managed Services Start Date'] = pd.to_datetime(df['Managed Services Start Date'])
    #df['Managed Services Start Date'] = df['Managed Services Start Date'].dt.strftime('%m/%d/%Y')
    
    df['Managed Services End Date'] = pd.to_datetime(df['Managed Services End Date'])
    #df['Managed Services End Date'] = df['Managed Services End Date'].dt.strftime('%m/%d/%Y')

    print(df.head())
    #ordenar columnas
    df = df.sort_values('ECSD', ascending=True)
    print(df)
    #df.to_excel('salida.xlsx',index=False)
    return df
def CombinarFechas(df):
    print(df)
    newdf = pd.DataFrame(columns=df.columns)
    #crear variable de fecha nula
    for index, row in df.iterrows():
        print('primer print')
        print(row)
        contract_start_date = row['Contract Start Date']
        contract_end_date = row['Contract End Date']
        managed_start_date = row['Managed Services Start Date']
        managed_end_date = row['Managed Services End Date']

        if pd.isna(contract_start_date) == True and pd.isna(managed_start_date) == False:
            row['Contract Start Date'] = row['Managed Services Start Date']
        if pd.isna(contract_end_date) == True and pd.isna(managed_end_date) == False:
            row['Contract End Date'] = row['Managed Services End Date']
        
        if pd.isna(contract_start_date) == False and pd.isna(managed_start_date) == False:
            if row['Contract Start Date'] != row['Managed Services Start Date']:
                raise
        if pd.isna(contract_end_date) == False and pd.isna(managed_end_date) == False:
            if row['Contract End Date'] != row['Contract End Date']:
                raise
        print("*******************************************")
        print('segundo print')
        print(row)
        print(df.columns.tolist())
        new_row = pd.DataFrame({'ECSD Month':[row['ECSD Month']], 'Stage':[row['Stage']], 'Opportunity ID':[row['Opportunity ID']], 'Opportunity Name':[row['Opportunity Name']], 'Deal Description':[row['Deal Description']], 
                                'PSD':[row['PSD']], 'ECSD':[row['ECSD']], 'Contract Start Date':[row['Contract Start Date']], 
                                'Contract End Date':[row['Contract End Date']], 'Managed Services Start Date':[row['Managed Services Start Date']], 
                                'Managed Services End Date':[row['Managed Services End Date']], 'Sub Service Group':[row['Sub Service Group']], 'Win Prob%':[row['Win Prob%']], 'Revenue':[row['Revenue']]})
        newdf = pd.concat([newdf, new_row], ignore_index=True)
        print(newdf)

    
    columnas_eliminar = ['Managed Services Start Date','Managed Services End Date']
    newdf = newdf.drop(columnas_eliminar,axis=1)

    #newdf['Contract Start Date'] = newdf['Contract Start Date'].dt.strftime('%m/%d/%Y')
    #newdf['Contract End Date'] = newdf['Contract End Date'].dt.strftime('%m/%d/%Y')
    print(newdf)


    writer = pd.ExcelWriter(archivo_fuente_ruta+'salida/output.xlsx', engine='xlsxwriter')
    newdf.to_excel(writer, index=False, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})
    worksheet.set_column('F:F', None, date_format)
    worksheet.set_column('G:G', None, date_format)
    worksheet.set_column('H:H', None, date_format)
    worksheet.set_column('I:I', None, date_format)

    writer.save()
    

def main():
    df = LeerArchivo(archivo_fuente_ruta)
    df = Transformaciones(df)
    CombinarFechas(df)

         

if __name__ == "__main__":
    main()