

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

query = '''
BEGIN
  DECLARE @dayCount INT;
  DECLARE @wkDay VARCHAR(10);

  SELECT @wkDay = DATENAME(weekday, GETDATE() )

  -- Based on today's day of the week, set the offset for the query 
  SELECT @wkDay;
  IF @wkDay IN ('Tuesday', 'Wednesday', 'Thursday', 'Friday')
    BEGIN
	  SET @dayCount = 1;
	END
	ELSE
	BEGIN
	  SET @dayCount = 3;
	END
END;'''

df = pd.read_sql(query, cnxn)
print(df)
dayCount = df.to_string()

cnxn1 = pyodbc.connect('''Driver={ODBC Driver 17 for SQL Server}; Server=SQL2014; Database=NWDComplaints; Trusted_Connection=yes;''')
cursor1 = cnxn.cursor()

query1 = '''
WITH Complaints_CTE (DayOfWeek, ComplaintNum, Status, SoNum, SOItem, OpenDate, Product, Description, PartCost, 
     Qty, RC_Description, CS_Dscription, AffQty, ComplaintCost, RespDept)
AS
(
SELECT DATENAME(weekday, CD.[OpenDate]) AS DayOfWeek
      ,CD.[ComplaintNum],CD.[Status],CD.[SoNum]
      ,CD.[SOItem],CD.[OpenDate] ,CD.[Product]
      ,CD.[Description],CD.[Expr9] As PartCost
      ,CD.[Qty],CD.[RC_Description]
      ,CD.[CS_Description],CD.[AffQty]
	  ,CD.[AffQty]*CD.[Expr9] AS ComplaintCost
	  ,M.[RespDept]
  FROM [NwdComplaints].[dbo].[vComplaintDetails] AS CD
  LEFT OUTER JOIN [NwdComplaints].[dbo].[ComplaintMast] AS M
  ON (CD.[ComplaintNum] = M.[ComplaintNum])
  WHERE CD.[Status] IN ('Active', 'Closed')
)

	  SELECT ComplaintNum, SoNum, SOItem, Product, RC_Description, ComplaintCost, RespDept
	  FROM Complaints_CTE
	  WHERE  OpenDate >=DATEADD(day, DATEDIFF(day,? ,GETDATE()),0) AND OpenDate < DATEADD(day, DATEDIFF(day,0,GETDATE()),0);'''
parameters = ['dayCount']
#df1 = pd.read_sql(query1, cnxn1)
df1 = cursor1.execute(query1, parameters)
print(df1)


## Using df.loc to avoid the "SettingWithCopyWarning" in Pandas.  Had been using dfPrint=df[[x,y,z... names of columns]]
dfPrint = df1.loc[:, ['ComplaintNum', 'SoNum', 'SOItem', 'Product', 'RC_Description', 'ComplaintCost', 'RespDept']]
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