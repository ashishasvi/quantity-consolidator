import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows



def process_excel(input_file,output_file,sumcoloumn):
    df=pd.DataFrame(openpyxl.load_workbook(input_file).active.values)
    df.columns=df.iloc[0]
    df=df[1:]

    
    oldcol=list(df.columns)
    gplist=[]

     #input less than one coloumn 
    popcol=oldcol.pop(sumcoloumn-1)
    oldcol.append(popcol)
    df.columns= range(1,len(df.columns)+1)


    for i in range(1,len(df.columns)+1):
        if i==sumcoloumn:
            pass
        else:
            df[i]=df[i].astype(str)
            gplist.append(i)



    
    # grouped.iloc[1:]=df.groupby(gplist)[sumcoloumn].sum().reset_index()
    grouped=df.groupby(gplist)[sumcoloumn].sum().reset_index()
    


    grouped.columns=oldcol
    excel_book=openpyxl.Workbook()
    excel_out=excel_book.active
    
    
    for row in dataframe_to_rows(grouped,index=False,header=True):
        excel_out.append(row)
        
        
    
    excel_book.save('downloads/consolidate.xlsx')
