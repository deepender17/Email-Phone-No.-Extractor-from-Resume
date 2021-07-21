import os
import glob
import pandas as pd
from pdfminer.high_level import extract_text  #pip install pdfminer.six
import re
import docx2txt

Email=[]
Number=[]

path= input("Enter Your Path: " )
os.chdir(path)

phone_reg=re.compile(r'[\+\(]?[1-9][0-9.\-\(\)]{8,}[0-9]')
email_reg=re.compile('\S+@\S+')

for file in glob.glob("*.pdf"):
    x=extract_text(file)
    
    email=re.findall(email_reg,x)
    phone=re.findall(phone_reg,x)
    
    Email.append(email)
    Number.append(phone)
    
for file1 in glob.glob("*.docx"):
    my_text=docx2txt.process(file1)
    
    email1=re.findall(email_reg,my_text)
    phone1=re.findall(phone_reg,my_text)
    
    Email.append(email1)
    Number.append(phone1)
    
email_address = [' '.join(ele) for ele in Email]
phone_address = [' '.join(ele) for ele in Number]

print(email_address)
print(phone_address)


data =pd.DataFrame(
    {"Number" : phone_address,
     "Email" : email_address,
    })

data.to_excel("Output.xlsx") 
