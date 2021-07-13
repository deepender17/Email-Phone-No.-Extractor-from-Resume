#!/usr/bin/env python
# coding: utf-8

# In[8]:


#Dont change the r , change only addres path

input_folder=r"C:\Users\PC\Desktop\pdf\Input"
#input_folder=main folder where you stored all your resmume(like pdf)

result_file_folder=input_folder+"\\"+"Result.pdf"
#result_file_folder=The result merged pdf will created.

op_no_pdf_folder=input_folder+"\\"+"Numbers_pdf.xlsx"
#op_no_pdf_folder=The excle file will created with all the extracted phone no.

op_mail_pdf_folder=input_folder+"\\"+"Email_pdf.xlsx"
#op_mail_pdf_folder=The excle file will created with all the extracted email .


"""
pip install  pymupdf
pip install PyPDF2
pip install xlsxwriter
pip install pandas
"""
import os
from PyPDF2 import PdfFileMerger
import fitz #fitz is pymupdf
import re
import xlsxwriter 
import pandas as pd

maindir = input_folder
data = os.listdir(maindir)

merger = PdfFileMerger(strict=False)

for file in data:
    if file.endswith(".pdf"):
        path_with_file = os.path.join(input_folder, file)
        #print(path_with_file)
        merger.append(path_with_file,  import_bookmarks=False )
merger.write(result_file_folder)

merger.close()

with fitz.open(result_file_folder) as doc:
    text = ""
    for page in doc:
        text += page.getText()

with fitz.open(result_file_folder) as doc:
    text = ""
    for page in doc:
        text += page.getText()

email1 = re.findall('\S+@\S+.', text) 
number1 = re.findall(r'[\+\(]?[1-9][0-9.\-\(\)]{8,}[0-9]', text)

with fitz.open(result_file_folder) as doc:
    text = ""
    for page in doc:
        text += page.getText()
email1 = re.findall('\S+@\S+.', text) 
number1 = re.findall(r'[\+\(]?[1-9][0-9.\-\(\)]{8,}[0-9]', text)

number_list = [number1] 
email_list = [email1] 

def insert_data(listdata):
    wb = xlsxwriter.Workbook(op_no_pdf_folder)
    ws = wb.add_worksheet()
    row = 0
    col = 0
    for line in listdata:
        for item in line:
            ws.write(row, col, item)
            col += 1
            row += 1
            col = 0
 
    wb.close()

insert_data(number_list)

with fitz.open(result_file_folder) as doc:
    text = ""
    for page in doc:
        text += page.getText()

email1 = re.findall('\S+@\S+.', text) 
number1 = re.findall(r'[\+\(]?[1-9][0-9.\-\(\)]{8,}[0-9]', text)

number_list = [number1] 
email_list = [email1] 

def insert_data(listdata):
    wb = xlsxwriter.Workbook(op_mail_pdf_folder)
    ws = wb.add_worksheet()
    row = 0
    col = 0
    for line in listdata:
        for item in line:
            ws.write(row, col, item)
            col += 1
            row += 1
            col = 0
 
    wb.close()

insert_data(email_list)
print("DONE")

