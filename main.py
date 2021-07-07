import os
import urllib.request
from psycopg2 import connect
import xlwings as xw
import pandas as pd
from sqlalchemy import create_engine

# Def Connection to postgres container 
def connect_db():
    CONNECT = "postgresql://postgres:docker@localhost:5432/test_raizen"
    ENGINE = create_engine(CONNECT)

    return ENGINE

# Def Insert to postgres container 
def insert_postgres(df,engine):
    df.to_sql(
        'anp',
        con=engine,
        index=False,
        if_exists='replace'
    )

# Def Extract PivotCache
def pivotcache(macro, local, file):
    # Load Files
    WORKBOOK_EXTRACT = xw.Book(macro)
    WORKBOOK = xw.Book(file)

    # Run Macro
    extractmacro = WORKBOOK_EXTRACT.macro('click')
    extractmacro()

    # Remove old Files
    try:
        os.remove(local)
    except:
        print("Empty Folder - Missing Data File")

    # Save Response and Close Excel
    WORKBOOK.save(local)
    xw.apps.active.quit()

# Def Load Input DB
def loaddf(local):
    xl = pd.ExcelFile(local)
    xl = len(xl.sheet_names)

    DF_2 = pd.DataFrame(columns=['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE', 'Jan',
                                 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'TOTAL'])
    for num in range(0, xl - 1):
        DF_1 = pd.read_excel(local, sheet_name=num)
        DF_2 = DF_2.append(DF_1, ignore_index=True)

    DF_LOAD = pd.melt(DF_2, id_vars=['COMBUSTÍVEL', 'ANO', 'ESTADO'], value_vars=DF_2.columns[DF_2.columns.get_loc(
        'Jan'):DF_2.columns.get_loc('TOTAL')], var_name=['month'], value_name="volume")

    DF_LOAD['unit'] = DF_LOAD['COMBUSTÍVEL'].str.extract("\(([^)]+)")
    DF_LOAD['year_month'] = DF_LOAD['ANO'].astype('str') + '_' + DF_LOAD['month']
    DF_LOAD['volume'] = DF_LOAD['volume'].replace(',','.').astype('float')
    DF_LOAD['COMBUSTÍVEL'] = DF_LOAD['COMBUSTÍVEL'].str.replace('\(m3\)', '').str.rstrip()
    DF_LOAD.rename({'ESTADO':'uf','COMBUSTÍVEL': 'product', 'ANO': 'year'}, axis="columns", inplace=True)
    DF_LOAD.drop(columns=['year', 'month'], inplace=True)
    
    return DF_LOAD

if __name__ == "__main__":
    # File Name
    DATABASE_DERIVADOS_LOCATION = "./dist/vendas-combustiveis-m3-derivados.xls"
    DATABASE_DIESEL_LOCATION = "./dist/vendas-combustiveis-m3-diesel.xls"

    # Macro Name
    MACRO_DERIVADOS_LOCATION = "./dist/extractpivotcachederivados.xlsm"
    MACRO_DIESEL_LOCATION = "./dist/extractpivotcachediesel.xlsm"

    # URL File
    URL = "http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls"
    FILE_NAME, headers = urllib.request.urlretrieve(URL)

    pivotcache(MACRO_DERIVADOS_LOCATION, DATABASE_DERIVADOS_LOCATION, FILE_NAME)
    pivotcache(MACRO_DIESEL_LOCATION, DATABASE_DIESEL_LOCATION, FILE_NAME)
    DF_DER = loaddf(DATABASE_DERIVADOS_LOCATION)
    DF_DIE = loaddf(DATABASE_DIESEL_LOCATION)
    DF = DF_DER.append(DF_DIE)

    ENGINE = connect_db()
    insert_postgres(DF,ENGINE)

    print('Successfully Inserted')