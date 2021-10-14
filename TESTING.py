

""" Downloads data from Complaints Excel file and Sales Excel file
    Creates Monthly reports with tables and graphs before emailing """

import win32com.client as win32
import sys, os
import pandas as  pd
import datetime
from datetime import date
from datetime import timedelta
from pathlib import Path
from tabulate import tabulate
from IPython.display import HTML
import pyodbc

## This establishes the connection to SQL Server as I need realtime data to create the report rather than reading an Excel spreadsheet

cnxn = pyodbc.connect('''Driver={ODBC Driver 17 for SQL Server}; Server=SQL2014; Database=NWDComplaints; Trusted_Connection=yes;''')
cursor = cnxn.cursor()

query = '''SELECT CD.[ComplaintNum],CD.[Status],CD.[SoNum]
      ,CD.[SOItem],CD.[OpenDate],CD.[Product]
      ,CD.[Description],CD.[Expr9] As PartCost
      ,CD.[Qty],CD.[RC_Description]
      ,CD.[CS_Description],CD.[AffQty]
      ,CD.[AffQty]*CD.[Expr9] AS ComplaintCost
	  ,M.[RespDept] As Dept
  FROM [NwdComplaints].[dbo].[vComplaintDetails] AS CD
  JOIN [NwdComplaints].[dbo].[ComplaintMast] AS M
  ON (CD.[ComplaintNum] = M.[ComplaintNum])
  WHERE CD.[Status] IN ('Active', 'Closed') AND
        M.[RespDept] IN ('SHIPWH') AND
		CD.[OpenDate] >=DATEADD(day, DATEDIFF(day,1,GETDATE()),0) 
        AND CD.[OpenDate] < DATEADD(day, DATEDIFF(day,0,GETDATE()),0);'''
df = pd.read_sql(query, cnxn)


## Using df.loc to avoid the "SettingWithCopyWarning" in Pandas.  Had been using dfPrint=df[[x,y,z... names of columns]]
dfPrint = df.loc[:, ['Dept', 'ComplaintNum', 'SoNum', 'SOItem', 'Product', 'PartCost', 'RC_Description', 'AffQty', 'ComplaintCost']]
dfPrint.sort_values(by=['Dept'], inplace=True)

####################################################################################################
####################################################################################################

if __name__ == "__main__":
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # mail.To = 'gnetherton@northwestdoor.com; mmartin@northwestdoor.com; sjones@northwestdoor.com; jfrench@northwestdoor.com; choffman@northwestdoor.com; wbaer@northwestdoor.com'
    mail.To = 'gnetherton@northwestdoor.com'
    mail.Subject = 'Shipping Complaint Summary '+ datetime.datetime.now().strftime("%B %d, %Y")
    mail.Body = '''Please find data attached and below.\n\n
               {}'''.format(dfPrint.to_string())
    mail.HTMLBody = '''<h3>Please find yesterday's shipping complaint data below.</h3> {}'''.format(dfPrint.to_html())

    mail.Send()