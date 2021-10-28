

""" Downloads data from Complaints Excel file and Sales Excel file
    Creates Monthly reports and emails """

import win32com.client as win32
import pandas as  pd
import datetime
import calendar
from datetime import date
from datetime import timedelta
# from datetime import datetime
from pathlib import Path
from tabulate import tabulate
from IPython.display import HTML
import pyodbc




## This establishes the connection to SQL Server as I need realtime data to create the report rather than reading an Excel spreadsheet
cnxn1 = pyodbc.connect('''Driver={ODBC Driver 17 for SQL Server}; Server=SQL2014; Database=NWDComplaints; Trusted_Connection=yes;''')
cursor1 = cnxn1.cursor()

query1 = '''
WITH Complaints_CTE (DayOfWeek, ComplaintNum, Status, SoNum, SOItem, OpenDate, Product, Description, PartCost, 
     Qty,  RC_Description, MfCustName, CS_Dscription, AffQty, ComplaintCost, RespDept)
AS
(
SELECT DATENAME(weekday, CD.[OpenDate]) AS DayOfWeek
      ,CD.[ComplaintNum],CD.[Status],CD.[SoNum]
      ,CD.[SOItem],CD.[OpenDate] ,CD.[Product]
      ,CD.[Description],CD.[Expr9] As PartCost
      ,CD.[Qty],CD.[RC_Description], CD.[MfCustName]
      ,CD.[CS_Description],CD.[AffQty]
	  ,CD.[AffQty]*CD.[Expr9] AS ComplaintCost
	  ,M.[RespDept]
  FROM [NwdComplaints].[dbo].[vComplaintDetails] AS CD
  LEFT OUTER JOIN [NwdComplaints].[dbo].[ComplaintMast] AS M
  ON (CD.[ComplaintNum] = M.[ComplaintNum])
  WHERE CD.[Status] IN ('Active', 'Closed') AND M.[RespDept] IN ('SHIPWH')
)

	  SELECT ComplaintNum, SoNum, SOItem, OpenDate, Product, RC_Description, ComplaintCost, RespDept, MfCustName
	  FROM Complaints_CTE
	  WHERE  OpenDate >=DATEADD(day, DATEDIFF(day,1 ,CAST(GETDATE() AS date)),0) AND OpenDate < DATEADD(day, DATEDIFF(day,0,CAST(GETDATE() AS date)),0);'''

query3 = '''
WITH Complaints_CTE (DayOfWeek, ComplaintNum, Status, SoNum, SOItem, OpenDate, Product, Description, PartCost, 
     Qty, RC_Description, MfCustName, CS_Dscription, AffQty, ComplaintCost, RespDept)
AS
(
SELECT DATENAME(weekday, CD.[OpenDate]) AS DayOfWeek
      ,CD.[ComplaintNum],CD.[Status],CD.[SoNum]
      ,CD.[SOItem],CD.[OpenDate] ,CD.[Product]
      ,CD.[Description],CD.[Expr9] As PartCost
      ,CD.[Qty],CD.[RC_Description], CD.[MfCustName]
      ,CD.[CS_Description],CD.[AffQty]
	  ,CD.[AffQty]*CD.[Expr9] AS ComplaintCost
	  ,M.[RespDept]
  FROM [NwdComplaints].[dbo].[vComplaintDetails] AS CD
  LEFT OUTER JOIN [NwdComplaints].[dbo].[ComplaintMast] AS M
  ON (CD.[ComplaintNum] = M.[ComplaintNum])
  WHERE CD.[Status] IN ('Active', 'Closed') AND M.[RespDept] IN ('SHIPWH')
)

	  SELECT ComplaintNum, SoNum, SOItem, OpenDate, Product, RC_Description, ComplaintCost, RespDept, MfCustName
	  FROM Complaints_CTE
	  WHERE  OpenDate >=DATEADD(day, DATEDIFF(day,3 ,CAST(GETDATE() AS date)),0) AND OpenDate < DATEADD(day, DATEDIFF(day,0,CAST(GETDATE() AS date)),0);'''

# This determines the day of the week for SQL offset reasons
curr_date = date.today()
dayOfWeek = (calendar.day_name[curr_date.weekday()])

if dayOfWeek in ['Tuesday', 'Wednesday', 'Thursday', 'Friday']:
  query = query1
else:
  query = query3

df1 = pd.read_sql(query, cnxn1)
print(df1.to_string() + "\n")
cursor1.close()
cnxn1.close()

#######################################################################################################################


dfPrint = df1.loc[:, ['RespDept','ComplaintNum', 'SoNum', 'SOItem', 'Product', 'MfCustName', 'ComplaintCost', 'RC_Description']]
print(dfPrint)


####################################################################################################
####################################################################################################

if __name__ == "__main__":
    #qcReport = generate_report('C:/temp/complaintsdept.pdf', "Complaints Report generated " + datetime.datetime.now().strftime("%B %d, %Y"), str(dfPrint))
    #qcReport = generate_report('C:/temp/complaintsdept.pdf', "Complaints Report generated " + datetime.datetime.now().strftime("%B %d, %Y"), (dfPrint.to_string()))
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # mail.To = 'gnetherton@northwestdoor.com; mmartin@northwestdoor.com; sjones@northwestdoor.com; jfrench@northwestdoor.com; choffman@northwestdoor.com; wbaer@northwestdoor.com'
    mail.To = 'gnetherton@northwestdoor.com; choffman@northwestdoor.com'
    mail.Subject = 'Shipping Complaint Summary '+ datetime.datetime.now().strftime("%B %d, %Y")
    mail.Body = '''Please find data attached and below.\n\n
               {}'''.format(dfPrint.to_string())
    mail.HTMLBody = '''<h3>Please find yesterday's shipping complaint data below.</h3> {}'''.format(dfPrint.to_html())

    # To attach a file to the email (optional):
    #attachment  = "C:/temp/complaintsdept.pdf"
   # mail.Attachments.Add(attachment)

    mail.Send()