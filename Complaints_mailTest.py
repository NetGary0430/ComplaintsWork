

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



#######################################################################################################################
import pyodbc
import pandas as pd

cnxn = pyodbc.connect('''Driver={ODBC Driver 17 for SQL Server}; Server=SQL2014; Database=NWDComplaints; Trusted_Connection=yes;''')
cursor = cnxn.cursor()

query = '''SELECT CD.[ComplaintNum],CD.[Status],CD.[SoNum]
      ,CD.[SOItem],CD.[OpenDate],CD.[ClassDesc]
      ,CD.[Description],CD.[Expr9] As PartCost
      ,CD.[Qty],CD.[RC_Description]
      ,CD.[CS_Description],CD.[AffQty]
	  ,M.[RespDept]
  FROM [NwdComplaints].[dbo].[vComplaintDetails] AS CD
  JOIN [NwdComplaints].[dbo].[ComplaintMast] AS M
  ON (CD.[ComplaintNum] = M.[ComplaintNum])
  WHERE CD.[Status] IN ('Active', 'Closed') AND
        CD.[ComplaintNum] > 50000 AND
		CD.[OpenDate] >=DATEADD(day, DATEDIFF(day,1,GETDATE()),0) 
        AND CD.[OpenDate] < DATEADD(day, DATEDIFF(day,0,GETDATE()),0);'''
df = pd.read_sql(query, cnxn)

cols = ['ClassDesc', 'SoNum','RC_Description', 'CS_Description']
df['ClassDesc'].str.strip()
df['SoNum'] = df['SoNum'].apply(int)
newMssgString = df[cols].to_markdown()
newMssgString2 = df[cols].to_html
tbl_Out = HTML(df.to_html(classes='table table-striped'))
print(newMssgString)
print('\n\n')
print(df)
#######################################################################################################################

today = date.today()
yesterday = today - timedelta(days = 1)
messageString = ""
complaints_today = {}

for i in range(df.shape[0]):
    if df['OpenDate'].values[i] == yesterday:        
        if df['ClassDesc'].values[i] in complaints_today:
            complaints_today[df['ClassDesc'].values[i]] += str(df['SoNum'].values[i]) + " :  " + str(df['RC_Description'].values[i]) + ",\n "
        else:
            complaints_today[df['ClassDesc'].values[i]] = str(df['SoNum'].values[i]) + " :  " + str(df['RC_Description'].values[i]) + ",\n "

for key, value in complaints_today.items():
    messageString += key + "\n" + value + "\n"


#messageString = messageString.replace('\n', '<br/>')
dfPrint = df[['ComplaintNum', 'Status', 'SoNum', 'SOItem', 'ClassDesc', 'PartCost', 'RC_Description', 'CS_Description', 'AffQty', 'RespDept']]
####################################################################################################
from reportlab.platypus import Paragraph, Spacer, Table, Image, SimpleDocTemplate
from reportlab.lib.styles import getSampleStyleSheet
import os, sys
from reportlab.pdfgen import canvas

def generate_report(attachment, title, paragraph):
    report = SimpleDocTemplate(attachment)

    print("--------------------------------")
    print(str(df[['ComplaintNum', 'Status', 'SoNum', 'SOItem', 'ClassDesc', 'PartCost', 'RC_Description', 'CS_Description', 'AffQty', 'RespDept']]))
    print(str(dfPrint))
    print("--------------------------------")

    flowables = []

    styles = getSampleStyleSheet()

    report_title = Paragraph(title, styles["h1"])
    paragraph1 = Paragraph(paragraph, styles["BodyText"])

    flowables.append(report_title)
    flowables.append(paragraph1)

    report.build(flowables)
####################################################################################################
####################################################################################################
####################################################################################################

if __name__ == "__main__":
    qcReport = generate_report('C:/temp/complaintsdept.pdf', "Complaints Report generated " + datetime.datetime.now().strftime("%B %d, %Y"), str(dfPrint))
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'gnetherton@northwestdoor.com'
    mail.Subject = 'Complaint Summary'
    mail.Body = messageString
    #mail.HTMLBody = tbl_Out #this field is optional

    # To attach a file to the email (optional):
    attachment  = "C:/temp/complaintsdept.pdf"
    mail.Attachments.Add(attachment)

    mail.Send()