# -*- coding: utf-8 -*-
"""
Created on Fri Jan 17 12:01:38 2020
"""

import os
import pandas as pd
import numpy as np
import re
from time import gmtime, strftime

directory=input("Please provide input Directory path : ")
fList=os.listdir(directory)

dtime=strftime("%Y%m%d%H%M%S",gmtime())
os.chdir(directory)
inputDir=os.path.normpath(os.getcwd() + os.sep + os.pardir)
outputDir='output' 
path=os.path.join(inputDir,outputDir,dtime)
os.makedirs(path) 
print(80 * '*' + "\n*"+ 24 * ' ' + "Welcome to csvMultiDimensional" 
          + 24 * ' ' +'*\n'+ 80 * '*')
print("List of Files in " + directory + "\n" + "\n".join(fList))
print(80*'*')
for filename in fList:
    fname=os.path.basename(filename).split('.')[0]
    print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) 
    + " INFO : Processing Started for " + fname)
    df=pd.read_excel(directory + '\/' + filename)
    df.columns=['col0','col1','col2','col3','col4']
    df[['col5','col6']]=df['col3'].str.split("_",
      expand=True)
    #df['col4']=
    df['Grade1']=np.where(df['col4'].str.contains("/",regex=True),
           df['col4'].str.split("/",expand=False).str[0].str.capitalize()
         ,df['col4']         
         )
    df['Climate']=np.where(df['col4'].str.contains("/",regex=True),
           df['col4'].str.split("/",expand=False).str[1].str.capitalize().str.split(" ",expand=False).str[1].str.capitalize()
         ,''       
         )    
   df['Profile']=np.where(df['col4'].str.contains("/",regex=True),
           df['col4'].str.split("/",expand=False).str[1].str.capitalize().str.split(" ",expand=False).str[2].str.capitalize()
         ,''       
         ) 
    
    df[['Grade1','Climate','Profile']]
    
    cli=df['col4'].str.split("/",expand=False).str[1]
    df['Grade']=df['col4'].str.split("/",expand=False).str[0].str.capitalize()
    df['Climate']=cli.str.split(" ",expand=False).str[-2].str.capitalize()
    df['Profile']=cli.str.split(" ",expand=False).str[-1].str.capitalize()
    df1=df[['col0','col1','col2','col5','col6','Grade','Climate','Profile']]
    print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) 
    + " INFO : Converting into Multi Dimensional")
    df2=df1.melt(id_vars=['col0','col1','col2',
                          'col5','col6'],
                 var_name='MultiDimensional',
                 value_name='Value')
    print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) 
    + " INFO : Converting into pivot table")
    dfinal=pd.pivot_table(
            df2,values='Value',
            index=['col0','col1','col2',
                   'col5','MultiDimensional'],
            columns='col6',
            aggfunc=np.sum)
    dfinal.replace(np.nan,0,inplace=True)
    dfinal.replace(0,'',inplace=True)
    print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) 
    + " INFO : Reseting Index on pivot table")
    dfinal.reset_index(inplace=True)
    dfinal.rename(columns={'col0':'ID',
                           'col1':'Class',
                           'col2':'Location Number',
                           'col5' : 'Root_Optimization_Year'
                           },inplace=True)
    flname= path +"\/" + fname + ".xlsx"
    wrt=pd.ExcelWriter(flname)
    print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) 
    + " INFO : Data writing into file")
    dfinal.to_excel(wrt,fname
                   ,index=False)
    wrt.save()
    print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) + " INFO : Process Completed" 
          + "\n" + 80 * '*')
