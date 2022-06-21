from time import sleep
from selenium import webdriver 
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
  

def ContentCheck(contextWithoutC,sId,content):

    wkdir = "./"+sId + "./"
    sId = wkdir + sId
    checkBool = True
    correctNum = 0
    itemNum = 0
    uncorrect = 0
    uncorrectList = []
    uncorrectPageNum = []

    content #從網頁資訊抓下來的
    needToCheckSet = content.split('\n')
    for i in needToCheckSet:
      print(i)
    
    
#下載位置設定


options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 0,'download.default_directory': os.getcwd()} #更改下載路徑
options.add_experimental_option('prefs', prefs)


driver = webdriver.Chrome(chrome_options=options)
driver.get("https://cloud.ncl.edu.tw/cycu/in.php?school_id=13")

sleep(1)
driver.find_element_by_name("FO_userid").send_keys("userId")
driver.find_element_by_name("FO_passwd").send_keys("Password")
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
  sleep(10)
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
  ContentCheck(contextWithoutC,sId,content)
  sleep(3)

  driver.close()
  driver.switch_to.window(home_page)

# 接下來想辦法找到ContentCheck所需要的值
#(contextWithoutC,yourContentPageS,yourContentPageE,yourContextStart,fileName):





