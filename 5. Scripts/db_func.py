import pyodbc as pyd, os, sqlalchemy as sa
from defined_func import getReportPeriod
from parameters import *
import pandas as pd

def loadToDataBase():
    removeExistingRecord()
    conx = sa.create_engine(f'mssql+pyodbc://{getServerName()}/{getDatabaseName()}?trusted_connection=yes&driver=ODBC Driver 17 for SQL Server')
    df= pd.read_excel(os.path.join(getEtlFolder(), "etl_file.xlsx"))

    for ind,rw in df.iterrows():
        print(rw[3])

    try:
        df.to_sql(getTableName(), con= conx, if_exists="append", index=False)
        value = "ETL was Successful"
    except:
        value = "There is a failure"
    
    return value


def removeExistingRecord():
    con_str = "Driver={SQL Server};Server=" + getServerName() + ";Database=" + getDatabaseName() + ";TrustedConnection=Yes;"
    query  = f"Delete from {getTableName()} where Report_Period = '{getReportPeriod()[2]}' "
    
    try :
        conx = pyd.connect(con_str)
        cursor = conx.cursor()
        cursor.execute(query)
        cursor.commit()
        
        conx.close()
        value= "Successful"
       
    except:
        value = "Failure"    
              
    #df = pd.io.sql.read_sql(query,conx)


    return  value


