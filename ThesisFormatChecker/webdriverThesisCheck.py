from time import sleep
from selenium import webdriver 
import pandas as pd
import os
import re
import shutil
import pdfplumber

def FindStudentID(tb1_trlist):
  for item in tb1_trlist:
    if "學號" in item.text:
      studentId = item.text 
      studentId = re.sub("\D","",studentId)
      return studentId    

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
def FindS(contextWithoutC):
  
  for i in range(len(contextWithoutC)):
    contextNotWhiteS = contextWithoutC[i].replace(" ","")
    contextNotWhiteS = contextNotWhiteS.replace("\n","")
    contextNotWhiteS = contextNotWhiteS.replace("\t","")
    contextNotWhiteS = contextNotWhiteS.replace("-","")#特殊Case
    contextNotWhiteS = contextNotWhiteS.replace(".","") 
    if len(contextNotWhiteS) > 0:
      if contextNotWhiteS[-1] == '1':
        return i + 1
  

def ContentCheck(contextWithoutC,sId,content,yourContextStart):

    wkdir = "./"+sId + "./"
    sId = wkdir + sId
    checkBool = True
    correctNum = 0
    itemNum = 0
    uncorrect = 0
    uncorrectList = []
    uncorrectPageNum = []

    content #從網頁資訊抓下來的
    needToCheckSet = content.split('\n')#以換行區分
    numberRe = re.compile(r'\d+$|[IXV]+$|[ixv]+$') # 這裡可能需要改
    digitNumber = False
    report = pd.DataFrame()  #報表
    
    for checkItem in needToCheckSet:

        tempCheckItem = checkItem
        checkItem = checkItem.replace('.',"") 
        checkItem = checkItem.replace('…',"")
        checkItem = checkItem.replace('\n',"")
        PageNumber = numberRe.search(checkItem)

        if PageNumber != None:      
            if str(PageNumber.group())[0] == 'I' or PageNumber.group()[0] == 'X' or PageNumber.group()[0] == 'V' or str(PageNumber.group())[0] == 'i' or str(PageNumber.group())[0] == 'x' or str(PageNumber.group())[0] == 'v':
                pageN=romanToInt(PageNumber.group() )
                if digitNumber :
                    break
            else :
                digitNumber = True
                try:
                    pageN = int(PageNumber.group()) -1 +yourContextStart-1
                except:
                    pageN=0
                    
            print(str(checkItem) + "number is "+str(pageN))
          
            titleRe = re.compile(r'\D*')
            titleContextG = titleRe.search(checkItem )
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
    
    #print("符合率"+str(correctNum/itemNum))
    print("檢驗的項目數"+str(itemNum))
    print("相符的項目數"+str(correctNum))
    print("錯誤的項目數"+str(uncorrect))
    report = pd.DataFrame((zip(uncorrectList, uncorrectPageNum)), columns = ['Ttitle', 'pageNum'])

    report.to_csv(sId+"report.csv",encoding = "utf_8_sig")
    return checkBool
    
#下載位置設定


options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 0,'download.default_directory': os.getcwd()} #更改下載路徑
options.add_experimental_option('prefs', prefs)


driver = webdriver.Chrome(chrome_options=options)
driver.get("https://cloud.ncl.edu.tw/cycu/in.php?school_id=13")

sleep(1)
driver.find_element_by_name("FO_userid").send_keys("user_id")
driver.find_element_by_name("FO_passwd").send_keys("password")
driver.find_element_by_name("Image6").click()
#登入

sleep(1)
all_thesis = driver.find_elements_by_name("FB_查核資料")
home_page = driver.current_window_handle
for i in range(len(all_thesis)):
  all_thesis[i].click()
  sleep(1)
  
  all_handles = driver.window_handles

  for handle in all_handles:
    if handle != home_page:
      driver.switch_to.window(handle)



#通過標籤名獲取表格中所有行
  tb1_trlist = driver.find_elements_by_xpath("//table[@id='tab_1']//tr")
  sId=FindStudentID(tb1_trlist)
#print(tb1_trlist[20].text) 固定的會出錯

#使用find_element_by...找到學號
#利用os建檔
  path = os.getcwd() + "/"+str(sId)
  os.mkdir(path)
  driver.find_element_by_id("ui_6").click()
  driver.find_element_by_partial_link_text("Full-Text").click()#下載好了
  #driver.find_element_by_partial_link_text("查看").click()#下載好了、授權書檔名不一樣，沒辦法直接丟 ，應該有辦法記住它的名稱
  sleep(5)
#把檔案移過去
  source = os.getcwd()+ "/"+"/Full-Text.pdf"
  destination = os.getcwd()+ "/"+str(sId) +"/"+"/Full-Text.pdf"
# 切tab_3 的class!=td1的td的br
  driver.find_element_by_id("ui_3").click()
  test = driver.find_elements_by_xpath("//table[@id='tab_3']//tr//td")
  content = test[1].text
  #print(content) #拿到論文目次
  sleep(3)
  #找開始頁
  shutil.move(source,destination)
  filename = "Full-Text.pdf"#電子檔統一名稱
  wkdir = "./"+sId + "./"#按照學號放資料夾
  filename = wkdir + filename
  pdfReader = pdfplumber.open(filename)
  contextWithoutC = readContext(pdfReader)
  sPage = FindS(contextWithoutC)
  print(sPage)
  ContentCheck(contextWithoutC,sId,content,sPage)
  sleep(3)

  driver.close()
  driver.switch_to.window(home_page)







