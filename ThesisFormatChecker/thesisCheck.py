# coding = utf-8 
import PyPDF2,os,stat
import pandas as pd
import pdfplumber
import re
from tika import parser
import pyinputplus as pyip
import shutil
import tkinter as tk
import time

def romanToInt(inputRoman):
    sum = 0
    inputRoman=inputRoman.replace(' ',"")
    inputRoman=inputRoman.replace('\t',"")
    inputRoman=inputRoman.replace('\n',"")
    convert={'X': 10,'V': 5,'I': 1,'v': 10,'x': 5,'i': 1} 
    for i in range(len(inputRoman)-1):            
        if convert[inputRoman[i]] < convert[inputRoman[i+1]]:                
            sum -= convert[inputRoman[i]]            
        else:                
            sum += convert[inputRoman[i]]        
    sum += convert[inputRoman[-1]]        
    return sum
    

def readContext( pdfReader ):    
    contextWithoutCover=[]
    checkV=True
    p0 = pdfReader.pages[0]
    text= p0.extract_text()
    contextWithoutCover.append(text)
    for i in range(1,len(pdfReader.pages)):
        pN = pdfReader.pages[i]
        text= pN.extract_text()
        contextWithoutCover.append(text)
        
    return contextWithoutCover
    
def ContentCheck(contextWithoutC,yourContentPageS,yourContentPageE,yourContextStart,fileName):
    checkBool = True
    correctNum = 0
    itemNum = 0
    uncorrect = 0
    uncorrectList = []
    uncorrectPageNum = []
    for i in range(yourContentPageS-1,yourContentPageE):    
        contextPage = contextWithoutC[i]
        needToCheckSet = contextPage.split('\n') #以換行區分
        numberRe = re.compile(r'\d+\s*$|[IXV]+\s*$|[ixv]+\s*$')
        digitNumber = False
        report = pd.DataFrame()
    
        for checkItem in needToCheckSet :
        
            tempCheckItem = checkItem
            checkItem = checkItem.replace('.',"")
            checkItem = checkItem.replace('…',"")
            checkItem = checkItem.replace('\n',"")
            PageNumber = numberRe.search(checkItem)
            if PageNumber != None:
                if str(PageNumber.group())[0] == 'I' or PageNumber.group()[0] == 'X' or PageNumber.group()[0] == 'V':
                    pageN=romanToInt(PageNumber.group() )
                    if digitNumber :
                        break
            
                else :
                    digitNumber = True
                    try:
                        pageN = int(PageNumber.group()) -1 +yourContextStart-1
                    except:
                        pageN=0
          #print("number is "+str(pageN))
          
                titleRe = re.compile(r'^.*\.|^.*\.\.\.')
                titleContextG = titleRe.search(tempCheckItem )
                if titleContextG != None:
                    titleContext = titleContextG.group()
                    titleContext = titleContext.replace('.',"")
                    titleContext = titleContext.replace('…',"")
                    titleContext = titleContext.replace('\n',"")
            
            #print(titleContext)
                    try :
                        titleContext=titleContext.replace('\t',"")
                        titleContext=titleContext.replace(' ',"")
                        titleCheckRe = re.compile(titleContext)
                        ContextDeleteEnter = contextWithoutC[pageN].replace('\n',"")
                        ContextDeleteEnter = ContextDeleteEnter.replace('\t',"")
                        ContextDeleteEnter = ContextDeleteEnter.replace(' ',"")
                        ContextDeleteEnter = ContextDeleteEnter.replace('.',"")
                        ContextDeleteEnter = ContextDeleteEnter.replace('...',"")
                        titleCheckG = titleCheckRe.search(ContextDeleteEnter)
                        itemNum = itemNum + 1
                        if titleCheckG != None:
                            correctNum = correctNum + 1
                        else :
                            uncorrect = uncorrect +1 
                            uncorrectList.append(titleContext)
                            uncorrectPageNum.append(pageN-yourContextStart+2)
                            print(titleContext)
                            checkBool = False 
                    except:
                        uncorrect = uncorrect +1 
                        uncorrectList.append("沒標題")
                        uncorrectPageNum.append(pageN-yourContextStart+2)
                        
                        checkBool = False 
            
                else :
                    print("something error")
                    
        #找出索引頁碼，對應contextWithoutC的index
    
    print("符合率"+str(correctNum/itemNum))
    print("檢驗的項目數"+str(itemNum))
    print("相符的項目數"+str(correctNum))
    print("錯誤的項目數"+str(uncorrect))
    report = pd.DataFrame((zip(uncorrectList, uncorrectPageNum)), columns = ['Ttitle', 'pageNum'])

    report.to_csv(fileName+"report.csv",encoding = "utf_8_sig")
    return checkBool
 


#列出所有PDF檔
#pdfFiles = []
#for filename in os.listdir('.'):
#    if filename.endswith('.pdf'):    
#        pdfFiles.append(filename)
    

#for filename in pdfFiles:  

while(True):
    studentID = input("請輸入學號")
    
    wkdir = os.get.getcwd()+"/"+studentID#改目錄要改這裡
    filename = "Full-Text.pdf"#檔名改這
    os.chdir(wkdir)
    try:
        pdfReader = pdfplumber.open(filename)
  
        contextWithoutC = readContext(pdfReader)
    
        yourContentPageS = pyip.inputNum("目錄開始於第幾頁",min=1,max=len(pdfReader.pages)) 
        yourContentPageE = pyip.inputNum("目錄結束於第幾頁",min=yourContentPageS,max=len(pdfReader.pages)) 
        yourContextStart = pyip.inputNum("正文於第幾頁開始",min=yourContentPageE,max=len(pdfReader.pages)) 
#yourContentPageS=6
#yourContentPageE=10
#yourContextStart=11
        correctOrNot=ContentCheck(contextWithoutC,yourContentPageS,yourContentPageE,yourContextStart,studentID)
        print(correctOrNot)
    except:
        print("something Error")
    
