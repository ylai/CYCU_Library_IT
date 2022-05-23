# -*- coding: utf-8 -*-

import pandas as pd 
import tkinter as tk

def main():
    try:
        studentData = pd.read_excel('sample.xlsx', sheetName.get("1.0","end").strip()) #讀取的excel

        print(studentData.shape[0])
        # 確定有幾位同學

        f = open( str(sheetName.get("1.0","end").strip()) + ".xml","w",encoding="utf-8") #輸出檔名

        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!--SQL Command:\n[College] = 全部學院\n[Year] = 89~110\n[ApprovalDate] = 2021-03-18~2021-11-01\n[CopyFullText] = \"N\"-->")
        f.write("\n<books>\n")

        for i in range(studentData.shape[0]):
            f.write("\t<book>\n\t\t<Creator>"+str(studentData.iloc[i,0])+"</Creator>\n")
            f.write("\t\t<Title>"+str(studentData.iloc[i,1]).replace('\n','').replace('\r','').replace('&','&amp;').replace('<','').replace('>','')+"</Title>\n")
            f.write("\t\t<Description.note.school>"+ str(schoolName.get("1.0","end").strip()) +"</Description.note.school>\n")
            f.write("\t\t<Description.note.department>"+str(studentData.iloc[i,2])+"</Description.note.department>\n")
            f.write("\t\t<Description.note.degree>"+str(studentData.iloc[i,3])+"</Description.note.degree>\n")
            f.write("\t\t<Description.note.year>"+str(studentData.iloc[i,4])+"</Description.note.year>\n")
            f.write("\t\t<DOI>"+str(studentData.iloc[i,6])+"</DOI>\n")
            f.write("\t</book>\n")
            
        f.write("</books>")
        f.close()
    
        state.set('完成轉檔！共 ' + str(studentData.shape[0]) + " 位學生")
    except Exception as e:
        print(e)
        state.set('轉檔失敗！')
    

if __name__ == '__main__':
    window = tk.Tk()
    window.title('論文清單轉XML工具')
    window.geometry('320x340')
    window.resizable(0 , 0)
    schoolNameLable = tk.Label(window, text='SchoolName:')
    schoolName = tk.Text(window, height=1)
    schoolNameLable.pack()
    schoolName.pack()
    sheetNameLable = tk.Label(window, text='SheetName:')
    sheetName = tk.Text(window, height=1)
    sheetNameLable.pack()
    sheetName.pack()
    
    workers = tk.Label(window, text='作者：賴卷狄 / 陳雋洋',font=('Arial', 12))
    workers.pack(side=tk.BOTTOM, pady=5)    

    copyRight = tk.Label(window, text='版權所有©:中原大學張靜愚圖書館',font=('Arial', 14))
    copyRight.pack(side=tk.BOTTOM, pady=5)    
    
    btn_start = tk.Button(window , text="開始轉檔" , height=3 , width=15 , font=20 ,command=main)
    btn_start.pack(pady=20)
    state = tk.StringVar()
    state.set('')
    
    stateBar = tk.Label(window, textvariable=state , anchor='w' , font=('Arial', 12))
    stateBar.pack(pady=5)
    
    window.mainloop()


