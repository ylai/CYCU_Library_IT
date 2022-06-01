# coding = utf-8 
import os
import pandas as pd
import pdfplumber
import re
import tkinter as tk

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
    p0 = pdfReader.pages[0]
    text= p0.extract_text()
    contextWithoutCover.append(text)
    for i in range(1,len(pdfReader.pages)):
        pN = pdfReader.pages[i]
        text= pN.extract_text()
        contextWithoutCover.append(text)
    pdfReader.close()
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
                        pageN = int(PageNumber.group()) -1 + yourContextStart-1
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
                            uncorrect = uncorrect + 1 
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
    

    result = "符合率:"+ str(correctNum/itemNum) + "\n"
    result += "檢驗的項目數:" +str(itemNum) + "\n"
    result += "相符的項目數:" +str(correctNum) + "\n"
    result += "錯誤的項目數:" +str(uncorrect ) + "\n"

    state.set(result)

    report = pd.DataFrame((zip(uncorrectList, uncorrectPageNum)), columns = ['Title', 'pageNum'])
    report.to_csv(fileName+"report.csv",encoding = "utf_8_sig")

    open_button = tk.Button(
    window,
    text='查看結果',
    command=open_file)
    open_button.pack(pady=20)

    btn_refresh = tk.Button(window , text="refresh" ,command=refresh)
    btn_refresh.pack(pady=20)

    return checkBool

def main():
    try:
        wkdir = os.getcwd() + "/" + str(studentId.get("1.0","end").strip()) #改目錄要改這裡
        filename = "Full-Text.pdf"#檔名改這
        os.chdir(wkdir)
        pdfReader = pdfplumber.open(filename)

        contextWithoutC = readContext(pdfReader)

        yourContentPageS = int(ContentStart.get("1.0","end").strip())
        yourContentPageE = int(ContentEnd.get("1.0","end").strip())
        yourContextStart = int(ThesisStart.get("1.0","end").strip())
        
        correctOrNot=ContentCheck(contextWithoutC,yourContentPageS,yourContentPageE,yourContextStart,str(studentId.get("1.0","end").strip()))
        print(correctOrNot)
        pdfReader.close()
    except:
        print("something Error")

def refresh():
    studentId.delete("1.0", "end")
    ContentStart.delete("1.0", "end")
    ContentEnd.delete("1.0", "end")
    ThesisStart.delete("1.0", "end")
    os.chdir('../')

def open_file():
    os.system("start EXCEL.EXE " + str(studentId.get("1.0","end").strip()) +"report.csv")

if __name__ == '__main__':
    window = tk.Tk()
    window.title('論文檢查工具')
    window.geometry('300x500')
    window.resizable(0 , 0)
    studentIdLable = tk.Label(window, text='學號:')
    studentId = tk.Text(window, height=1)
    studentIdLable.pack()
    studentId.pack()
    ContentStartLable = tk.Label(window, text='目錄開始頁數:')
    ContentStart = tk.Text(window, height=1)
    ContentStartLable.pack()
    ContentStart.pack()
    ContentEndLable = tk.Label(window, text='目錄結束頁數:')
    ContentEnd = tk.Text(window, height=1)
    ContentEndLable.pack()
    ContentEnd.pack()
    ThesisStartLable = tk.Label(window, text='正文開始頁數:')
    ThesisStart = tk.Text(window, height=1)
    ThesisStartLable.pack()
    ThesisStart.pack()

    btn_start = tk.Button(window , text="開始檢查" , height=3 , width=15 , font=20 ,command=main)
    btn_start.pack(pady=20)
    state = tk.StringVar()
    state.set('')

    stateBar = tk.Label(window, textvariable=state , anchor='w' , font=('Arial', 12))
    stateBar.pack(pady=5)

    window.mainloop()
    