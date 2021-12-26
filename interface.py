import tkinter as tk
from tkinter.constants import ANCHOR, COMMAND, END
import tkinter.ttk as ttk
from typing import Literal, ValuesView
import sheet_data

win = tk.Tk()                          #定義視窗
win.title('簡單輸入')                  #視窗名稱
scr_width = win.winfo_screenwidth()    #視窗寬度
scr_height = win.winfo_screenheight()  #視窗高度
width = 1200
height = 700
x = int((scr_width - width)/2)
y = int((scr_height - height)/3)
win.geometry('{}x{}+{}+{}'.format(width, height, x, y)) # 大小以及位置
win.resizable(0,0)              #視窗大小固定
win.config(bg="#525252")        #視窗背景
win.attributes("-alpha",0.9)    #透明度
win.attributes("-topmost",0)    #視窗置頂

#grid控制變數
# label_wight = 130        #對應padx(x方向填空)
# label_wight_inner = 0    #對應ipadx(內部x方向填空)
# label_hight_inner = 0    #對應ipady(內部y方向填空)

#設定選項
zh_font = 'Regular 13'     #中文字體 大小 #Regular標楷體
label_pady = 20            #label y填空

#標題欄位建立
for col_index in range(1,7):
    lbl_1 = tk.Label(win, text=sheet_data.col_name[col_index-1], font='Regular 16')
    lbl_1.grid(column=col_index, row=1, padx=8 , pady=label_pady)

#空白首欄建立
firstCol = tk.Label(win, bg="#525252")
firstCol.grid(column=0, row=0,padx=16 , pady=6)

#Combobox被選擇後的行為
def combobox_selected(event):
    # print(event.widget) #範例
    ans = event.widget.get()
    index = sheet_data.inspect.index(ans) #尋找目標在陣列中的位置
    for i in range(1,6):
        TextVar['Txt'+str(event.widget.RSNum)+'_'+str(i)].delete(1.0,'end') #entry清空
        TextVar['Txt'+str(event.widget.RSNum)+'_'+str(i)].insert(1.0 ,sheet_data.presentVal[index][i-1]) #插入值
        TextVar['Txt'+str(event.widget.RSNum)+'_'+str(i)].tag_add('configure',1.0,'end')
        TextVar['Txt'+str(event.widget.RSNum)+'_'+str(i)].tag_configure('configure',justify='center',font=zh_font)
#讓Text可以正常使用Tab鍵轉移聚焦
def focus_next_widget(event):
    event.widget.tk_focusNext().focus()
    return("break")    

RowSerNum = 0  #新增列_每列的編號
ComboVar = {}  #存放Combobox變數及物件
StrVar = {}    #存放Combobox字串變數及值
TextVar = {}  #存放Entry變數及物件

#新增列
def Add_NewRow():
    global RowSerNum
    global GetVar
    StrVar['CoB'+'Get'+str(RowSerNum)] = tk.StringVar
    ComboVar['CoB'+str(RowSerNum)] = ttk.Combobox(win, 
                                                  state='readonly', #只能讀取
                                                  textvariable=StrVar['CoB'+'Get'+str(RowSerNum)], 
                                                  justify='center',   #'Left'靠左 ,'center'置中,'right'靠右
                                                  font=zh_font) #27行，字體變數
    ComboVar['CoB'+str(RowSerNum)]['values'] =  sheet_data.inspect
    ComboVar['CoB'+str(RowSerNum)].current(RowSerNum)
    ComboVar['CoB'+str(RowSerNum)].grid(column=1 ,row=RowSerNum+2 ,padx=6 ,ipady=16, ipadx=6)
    ComboVar['CoB'+str(RowSerNum)].bind('<<ComboboxSelected>>',combobox_selected)
    ComboVar['CoB'+str(RowSerNum)].RSNum = RowSerNum #自訂義方法儲存編號
    for Col in range(1,6):
        TextVar['Txt'+str(RowSerNum)+'_'+str(Col)] = tk.Text(win,height=3.5,width=23 )
        TextVar['Txt'+str(RowSerNum)+'_'+str(Col)].insert(1.0 ,sheet_data.presentVal[RowSerNum][Col-1]) #利用插入方法，設定預設值(第幾次元,插入文字)
        TextVar['Txt'+str(RowSerNum)+'_'+str(Col)].tag_add('configure',1.0,'end') 
        TextVar['Txt'+str(RowSerNum)+'_'+str(Col)].tag_configure('configure',justify='center',font=zh_font) 
        TextVar['Txt'+str(RowSerNum)+'_'+str(Col)].bind("<Tab>", focus_next_widget) #Tab鍵正常發揮函式
        TextVar['Txt'+str(RowSerNum)+'_'+str(Col)].grid(column=Col+1 ,row=RowSerNum+2 ,padx=6 )
        
    RowSerNum += 1
    return RowSerNum

for i in range(7): #建立7列
    Add_NewRow()
    print('row:',i)

#建立圖片物件
img1 = tk.PhotoImage(file="Pic/wordicon.png")
img2 = tk.PhotoImage(file="Pic/pdficon.png")
img3 = tk.PhotoImage(file="Pic/excelicon.png")
btn_width = 100
btn_height = 100 
y_gap = 30

import docx
doc = docx.Document('doc/慶育油品範例.docx')
tables = doc.tables  

#word轉檔
def scan_word():   
#讀取表格文字
    #找尋目標cell的x座標
    coord_X = 0  
    for i , n in enumerate(tables[0].rows):
        for cell in n.cells:
            if cell.text == sheet_data.col_name[0]: #找尋標題欄
                coord_X = i + 1 #標題欄下一欄
                print(cell.text,coord_X)
                break
        else:
            continue
        break 
    #修改表格內容
    for r in range(0,RowSerNum): #欄的數量
        print(r)
        #扣除首列之列的格子數量
        for index_x,cell_y in enumerate([3,4,7,9,10]): #修正(word內實際出現次數)
            repls_text = TextVar['Txt'+str(r)+'_'+str(index_x+1)].get('1.0','end-1c')
            #從視窗取得更新文字
            #end為換行符號，-1c指向前一字元
            cell_x = r+coord_X #標題欄位置+往後第幾欄
            tables[0].cell(cell_x,cell_y).text= repls_text
                
    doc.save('doc/scan.docx')

#按鈕>>轉成word檔
btn_word = tk.Button(text="轉成word檔")
btn_word.grid(columnspan=2,column=1,row=RowSerNum+3,pady=y_gap)
btn_word.config(image=img1)

#按鈕>>轉成pdf檔
btn_pdf = tk.Button(text="轉成pdf檔")
btn_pdf.grid(columnspan=2,column=3,row=RowSerNum+3,pady=y_gap)
btn_pdf.config(image=img2)

#按鈕>>轉成execel檔
btn_pdf = tk.Button(text="轉成excel檔",command=scan_word)
btn_pdf.grid(columnspan=2,column=5,row=RowSerNum+3,pady=y_gap)
btn_pdf.config(image=img3)

win.mainloop()

