import pandas
import os 
import xlsxwriter

AgreeField = { 'Agree' : True, 'DisAgree': False, 'NoPaper': 'NoPaper' }
defaultValue = {'row' : 1 , 'col' : 0}

df = pandas.read_excel("sample.xlsx")
path = "Z:\\__論文審核專用\\華藝\\1102\\"

dirsArray = [] 
for root, dirs, files in os.walk(path):
    for folderName in dirs:
 
        if(len(folderName) == 8 and folderName.isdigit()):
            dirsArray.append(folderName)
            
        elif(folderName.find("_") != -1):
            tempFolder = folderName[0:8]
            if(tempFolder.isdigit()):
                 dirsArray.append(tempFolder)
                 
        elif(len(folderName) == 7 and folderName.isdigit()):
             dirsArray.append(folderName)
             
excelArray = []    

workbook = xlsxwriter.Workbook('export.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(defaultValue['row'] - 1, defaultValue['col'], '學號')
worksheet.write(defaultValue['row'] - 1, defaultValue['col'] + 1, '狀態')

for stuId in df["學號"]:
    if str(stuId) in dirsArray:
        worksheet.write(defaultValue['row'], defaultValue['col'], str(stuId))
        worksheet.write(defaultValue['row'], defaultValue['col'] + 1, "V")
    else:
        worksheet.write(defaultValue['row'], defaultValue['col'], str(stuId))
        worksheet.write(defaultValue['row'], defaultValue['col'] + 1, "X")
    defaultValue['row'] += 1
    
workbook.close()

