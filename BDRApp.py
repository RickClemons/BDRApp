#Chad's A/R Formatter

import pandas as pd 
import numpy as np 
from pandas import ExcelWriter
from textwrap import shorten
import getpass
import matplotlib
matplotlib.use('Agg')



fileLocation=input('Please enter document exact document name:')

readFile=pd.read_csv(r'/Users/rick.clemons/Downloads/' + str(fileLocation),engine='python')

#print(readFile.head())

drop_column=['Grant','Paid in full on','Payroll','Materials','Profit','Profit %','Crew', 'Payment','Super','Adjustments']
readFile.drop(drop_column, inplace=True, axis=1)

readFile=readFile.set_index('Invoice#')

df= pd.DataFrame(readFile)

df["Tuesday Notes"]=""

df["Friday Notes"]=""

fileFormat=df.rename(columns=lambda x: x.replace(' ',''))

print(fileFormat.head())

SalesRepBreakDown={k: v for k, v in fileFormat.groupby('SalesRep')} 

report=pd.ExcelWriter('ChadsARFromPY.xlsx')
for key in SalesRepBreakDown:
    if len(key)>31:
        shorten(key, width=15, placeholder="")
    
    else:
        SalesRepBreakDown[""+key+""].to_excel(report,""+key+"")
report.save()