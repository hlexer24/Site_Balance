import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import FORMULAE
from openpyxl import Workbook
import string
import re
from operator import *
import xlsxwriter
import xlwt
import streamlit as st

st.set_page_config(page_title='Site Balance')
st.title('Site Balance')
st.subheader("Instructions:")
st.subheader("In the first upload box, select the CND File. In the second upload box, select you downloaded file from Ops Tracker")
st.subheader("A File Name Site Automation Calculation will be generated to your User File, please locate the xlsx file in your user folder")
st.subheader("Now you can choose either to view the Excel file on your desktop or you can upload this file and it will be displayed on the website.")

uploaded_file1 = st.file_uploader('Choose a XLSX File', type='xlsx')
uploaded_file2 = st.file_uploader('Choose another XLSX File', type='xlsx')

i = 1


excel_file = pd.read_excel(uploaded_file1, sheet_name='FinalWithSectors')
excel_file[['Site Name','Tech','Mgr']]
Manager = excel_file[['Mgr','Tech']]
Tech = excel_file[['Tech']]

excel_file1 = pd.read_excel(uploaded_file2)
excel_file1[['Site Tech Name','Site Mgr. Name','Weight Call Volume','Weight Drive Time','Weight Equipment','Weight Site Access','Weight Site Type']]
SiteTotal = excel_file1[['Site Tech Name','Site Mgr. Name','Weight Call Volume','Weight Drive Time','Weight Equipment','Weight Site Access','Weight Site Type']]

writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')

Manager_Count = excel_file['Mgr'].value_counts()
print(Manager_Count)
Total = excel_file.groupby(['Tech','Mgr'])['Mgr'].value_counts()
print(Total)
Tech_Count = excel_file['Tech'].value_counts()
print(Tech_Count)



df = pd.DataFrame(Manager_Count)
df2 = pd.DataFrame(Manager)
df3 = pd.DataFrame(Total)
df4 = pd.DataFrame(SiteTotal)

df.to_excel(writer,sheet_name="Manager")
#df2.to_excel(writer,sheet_name="Tech")
df3.to_excel(writer,sheet_name="Total")
df4.to_excel(writer,sheet_name="Site_Totals")

writer.close()
wb = load_workbook('test.xlsx')

sheet1 = wb["Manager"] 
#sheet2 = wb["Tech"]
sheet3 = wb["Total"]
sheet4 = wb["Site_Totals"]

sheet1.column_dimensions['A'].width = 20
sheet1.column_dimensions['B'].width = 20
sheet1.column_dimensions['C'].width = 20
sheet1.column_dimensions['D'].width = 20
sheet1.column_dimensions['E'].width = 20
sheet1.column_dimensions['F'].width = 30
sheet1.column_dimensions['G'].width = 30
sheet1.column_dimensions['H'].width = 30
sheet1.column_dimensions['J'].width = 30
sheet1.column_dimensions['K'].width = 30
sheet1.column_dimensions['L'].width = 30
sheet1.column_dimensions['M'].width = 30

sheet3.column_dimensions['N'].width = 30
sheet3.column_dimensions['B'].width = 20
sheet3.column_dimensions['C'].width = 20
sheet3.column_dimensions['D'].width = 20
sheet3.column_dimensions['E'].width = 20
sheet3.column_dimensions['F'].width = 30
sheet3.column_dimensions['G'].width = 30
sheet3.column_dimensions['H'].width = 35
sheet3.column_dimensions['J'].width = 30
sheet3.column_dimensions['K'].width = 30
sheet3.column_dimensions['L'].width = 30
sheet3.column_dimensions['M'].width = 30
sheet3.column_dimensions['I'].width = 30

sheet4.column_dimensions['A'].width = 20
sheet4.column_dimensions['B'].width = 20
sheet4.column_dimensions['C'].width = 20
sheet4.column_dimensions['D'].width = 20
sheet4.column_dimensions['E'].width = 20
sheet4.column_dimensions['F'].width = 20
sheet4.column_dimensions['G'].width = 20
sheet4.column_dimensions['H'].width = 20

sheet1['B12'] = ('=SUM(B2:B11)')
sheet1['A12'] = 'Total'
sheet1['A1'] = 'Manager'
sheet1['B1'] = 'Current Site Count'
sheet1['A17'] = 'Target Avg.'
sheet1['B16'] = 'New Sites Per Tech'
sheet1['C16'] = 'New Points per Tech'
sheet1['D1'] = 'Current Techs'
sheet1['C1'] = 'Current Total Points'
sheet1['E1'] = 'Current Sites per Tech'
sheet1['F1'] = 'Current Points per Tech'
sheet1['G1'] = 'Suggested Change in Sites per Mgr'
sheet1['H1'] = 'Suggested Change in Points per Mgr'
sheet1['J1'] = 'Additional Headcount Needed'
sheet1['K1'] = 'New Total Headcount'
sheet1['L1'] = 'New Balance of Sites per Mgr'
sheet1['M1'] = 'New Balance of Points per Mgr'
sheet4['H1'] = 'Site Ranking Score'
sheet1['D1'].font = Font(bold=True)
sheet1['C1'].font = Font(bold=True)
sheet1['E1'].font = Font(bold=True)
sheet1['F1'].font = Font(bold=True)
sheet1['G1'].font = Font(bold=True)
sheet1['H1'].font = Font(bold=True)
sheet1['A12'].font = Font(bold=True)
sheet1['B1'].font = Font(bold=True)
sheet1['A17'].font = Font(bold=True)
sheet1['B16'].font = Font(bold=True)
sheet1['C16'].font = Font(bold=True)
sheet1['J1'].font = Font(bold = True)
sheet1['K1'].font = Font(bold = True)
sheet1['L1'].font = Font(bold = True)
sheet1['M1'].font = Font(bold = True)


sheet3['D1'].font = Font(bold=True)
sheet3['E1'].font = Font(bold=True)
sheet3['F1'].font = Font(bold=True)
sheet3['G1'].font = Font(bold=True)
sheet3['H1'].font = Font(bold=True)
sheet3['J1'].font = Font(bold = True)
sheet3['K1'].font = Font(bold = True)
sheet3['L1'].font = Font(bold = True)
sheet3['I1'].font = Font(bold = True)

sheet1['A20'] = 'New Target Avg.'
sheet1['B19'] = 'New Sites Per Tech'
sheet1['C19'] = 'New Points per Tech'
sheet1['J11'] = 'Total Head Add Count'
sheet1['K11'] = '=SUM(J2:J10)'
sheet1['B19'].font = Font(bold = True)
sheet1['C19'].font = Font(bold = True)
sheet3.auto_filter.ref = 'A1:B7584'

#sheet1['C19'].font = Font(bold=True)

sheet3['B1'] = 'Manager'
sheet3['C1'] = 'Current Site Count'
#sheet2['E16'] = 'New Sites Per Tech'
#sheet2['F16'] = 'New Points per Tech'
sheet3['D1'] = 'Current Total Points'
sheet3['E1'] = 'Current Sites per Tech'
sheet3['F1'] = 'Current Points per Tech'
sheet3['G1'] = 'Suggested Change in Sites per Tech'
sheet3['H1'] = 'Suggested Change in Points per Tech'
sheet3['I1'] = 'Additional Sites Added'
sheet3['J1'] = 'New Total Site Count'
sheet3['K1'] = 'New Balance of Sites per Tech'
sheet3['L1'] = 'New Balance of Points per Tech'


sheet1['B17'] =('=B12/D12')
sheet1['D12'] = ('=SUM(D2:D11)')
sheet1['B20'] = ('=B12/SUM(K2:K10)')
sheet1['C8'] = '=SUM(Site_Totals!G2:G1289)'
sheet1['C12'] = '=SUM(C2:C11)'
sheet1['C17'] = '=C12/D12'
sheet1['C20'] = '=C12/SUM(K2:K10)'



from_row = 2
to_row = 10

for i in range(from_row,to_row+1):
    sheet1[f"D{i}"] = f'=COUNTIF(Total!B2:B104,Manager!A{i})'
    sheet1[f"E{i}"] = f'=B{i}/D{i}'
    sheet1[f"G{i}"] = f'=(B17*D{i})-B{i}'
    sheet1[f"K{i}"] = f'=J{i}+D{i}'
    sheet1[f"L{i}"] = f'=(B20*K{i})-B{i}'
    sheet1[f"F{i}"] = f'=C{i}/D{i}'
    sheet1[f"H{i}"] = f'=(C17*D{i})-C{i}'
    sheet1[f"M{i}"] = f'=(C20*K{i})-C{i}'
    sheet1[f"C{i}"] = f'=SUMIF(Site_Totals!C2:C1289,Manager!A{i},Site_Totals!H2:H1289)'


from_row2 = 2
to_row2 = 104

for i in range(from_row2,to_row2+1):
    sheet3[f"E{i}"] = f'=Manager!B12/Total!C{i}'
    sheet3[f"G{i}"] = f'=(Manager!B17*1)-Total!C{i}'
    sheet3[f"J{i}"] = f'=I{i}+C{i}'
    sheet3[F"K{i}"] = f'=(Manager!B20*1)-Total!J{i}'
    sheet3[f"D{i}"] = f'=SUMIF(Site_Totals!B2:B1289,Total!A{i},Site_Totals!H2:H1289)'
    sheet3[f"F{i}"] = f'=D{i}/Manager!D{i}'
    sheet3[f"H{i}"] = f'=(Manager!C17*Total!D{i})-Total!C{i}'
    sheet3[f"L{i}"] = f'=(Manager!C20*K{i})-C{i}'

from_row3 = 2
to_row3 = 1289

for i in range(from_row3,to_row3+1):
    sheet4[F"H{i}"] = f'=AVERAGE(D{i}:G{i})'
sheet3.column_dimensions['A'].width = 20
sheet3.column_dimensions['B'].width = 20
sheet3.column_dimensions['C'].width = 20
sheet3.column_dimensions['D'].width = 20
sheet3['B106'] = 'Total'
sheet3['C106'] = '=SUM(C2:C105)'

wb.save('Site Automation Calculation.xlsx')
wb = Workbook("Site Automation Calculation.xlsx")

uploaded_file3 = st.file_uploader('Please Upload the Calculated Excel File', type='xlsx')
df5 = pd.DataFrame(uploaded_file3)
