import os, pyodbc, pandas as pd
import cx_Oracle as cx

conx = pyodbc.connect('Driver={SQL Server};Server=DESKTOP-022G187\SQLEXPRESS;Database=Retail;Trusted_Connection=yes;')
#conx = cx.connect('username/password@IP:Port/DatabaseName')
 
cursor = conx.cursor()
query ="Select * from tblsales"
df = pd.read_sql_query(query,conx)
df.to_csv(os.path.join(os.getcwd(),"total_sales.csv"))


