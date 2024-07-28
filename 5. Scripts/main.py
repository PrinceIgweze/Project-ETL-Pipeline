import os 
from db_func import * 
from parameters import *
from defined_func import *



if __name__ == "__main__":


    try:
        getReportFile()
        getvalues()
        loadToDataBase()

        value = "ETL was successful"
    except:
        value= "There was a failure"




