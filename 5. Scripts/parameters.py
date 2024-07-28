import pandas as pd, os 


#this returns the folder path where all excel files reside
def getDataSourceFolder():
    fol_path = os.path.join(os.getcwd(),"01.Parameters", 'parameters.xlsx')
    df = pd.read_excel(fol_path)
    filt = (df['Item'] == 'Data_Source')
    val = df[filt].iloc[0,1]

    return val



#this returns the number of previous months to locate
def getNoMonths():
    fol_path = os.path.join(os.getcwd(),"01.Parameters", 'parameters.xlsx')
    df = pd.read_excel(fol_path)
    filt = (df['Item'] == 'No_Months')
    val = df[filt].iloc[0,1]

    return val


def getLandingFolder():
    fol_path = os.path.join(os.getcwd(),"01.Parameters", 'parameters.xlsx')
    df = pd.read_excel(fol_path)
    filt = (df['Item'] == 'Landing_Folder')
    val = df[filt].iloc[0,1]

    return val


def getEtlFolder():
    fol_path = os.path.join(os.getcwd(),"01.Parameters", 'parameters.xlsx')
    df = pd.read_excel(fol_path)
    filt = (df['Item'] == 'ETL_Folder')
    val = df[filt].iloc[0,1]

    return val


def getServerName():
    fol_path = os.path.join(os.getcwd(),"01.Parameters", 'parameters.xlsx')
    df = pd.read_excel(fol_path)
    filt = (df['Item'] == 'Server_Name')
    val = df[filt].iloc[0,1]

    return val

def getDatabaseName():
    fol_path = os.path.join(os.getcwd(),"01.Parameters", 'parameters.xlsx')
    df = pd.read_excel(fol_path)
    filt = (df['Item'] == 'Database_Name')
    val = df[filt].iloc[0,1]

    return val



def getTableName():
    fol_path = os.path.join(os.getcwd(),"01.Parameters", 'parameters.xlsx')
    df = pd.read_excel(fol_path)
    filt = (df['Item'] == 'Table_Name')
    val = df[filt].iloc[0,1]

    return val

