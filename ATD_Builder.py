import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import os
import sys



docDesktop = DocxTemplate("templates/Template_Desktop_Form.docx")
docMonitor = DocxTemplate("templates/Template_Monitor_Form.docx")
docCable = DocxTemplate("templates/Template_Cable_Form.docx")

today_date = datetime.today().strftime("%d %b, %Y")
print(today_date)


df = pd.read_excel('manual_hris.xlsx')

batch = df['Batch'].iloc[0]
prepby = df['IssuedBy'].iloc[0]

savepath = f"{batch}-{prepby}"

print(savepath)

#path = df['Batch'].iloc[0]
path = savepath

try:
    os.mkdir(path)
except OSError as error:
    print(error)
    sys.exit(1) 


for index, row in df.iterrows():
    
    context = {
        'Batch' :row['Batch'],
        'Type' :row['Type'],
        'DateOfIssuance' :row['DateOfIssuance'],
        'EID' : row['EID'],
        'LastName' : row['LastName'],
        'FirstName' : row['FirstName'],
        'Program' :row['Program'],
        'Description' :row['Description'],
        'Serial' :row['Serial'],
        'Cost' :row['Cost'],
        'PrepBy' :row['PrepBy'],
        'IssuedBy' :row['IssuedBy'],
        'DeptIssued' :row['DeptIssued'],
        'Remarks' :row['Remarks']
        }

    print(row['Type'])
    filename = f" {row['Batch']} {row['EID']}-{row['LastName']},{row['FirstName']} {row['Type']} {row['Serial']}"
    print(f"{filename}.docx")
   


    if row['Type'] == 'Desktop':
        docDesktop.render(context)
        docDesktop.save(f"{savepath}/{filename}.docx")
     
    elif row['Type'] == "Monitor":
        docMonitor.render(context)
        docMonitor.save(f"{savepath}/{filename}.docx")
     
    elif row['Type'] == "Cable":
        docCable.render(context)
        docCable.save(f"{savepath}/{filename}.docx")




