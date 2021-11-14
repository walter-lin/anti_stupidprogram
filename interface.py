import tkinter as tk
from tkinter.constants import END
import tkinter.ttk as ttk
from typing import Literal, ValuesView
import sheet_data

win = tk.Tk()                   #定義視窗
win.title('簡單輸入')            #視窗名稱
win.geometry('1200x800')        #視窗大小
win.resizable(0,0)              #視窗大小固定
win.config(bg="#525252")        #視窗背景
win.attributes("-alpha",0.9)    #透明度
win.attributes("-topmost",0)    #視窗置頂
#grid控制變數
# label_wight = 130        #對應padx(x方向填空)
# label_wight_inner = 0    #對應ipadx(內部x方向填空)
# label_hight_inner = 0    #對應ipady(內部x方向填空)

#設定選項
zh_font = 'Regular 14'     #中文字體 大小 #Regular標楷體
label_pady = 20            #label y填空

#標題欄位建立
for col_index in range(1,7):
    lbl_1 = tk.Label(win, text=sheet_data.col_name[col_index-1], font='Regular 16')
    lbl_1.grid(column=col_index, row=1, padx=8 , pady=label_pady)

#空白首欄建立
# firstCol = tk.Label(win, bg="#525252")
# firstCol.grid(column=0, row=0,padx=50 , pady=8)

#Combobox被選擇後的行為
def combobox_selected(event):
    # print(event.widget) #範例
    ans = event.widget.get()
    index = sheet_data.inspect.index(ans)
    for i in range(1,6):
        EntryVar['Etr'+str(event.widget.RSNum)+'_'+str(i)].delete(0,END) #entry清空
        EntryVar['Etr'+str(event.widget.RSNum)+'_'+str(i)].insert(0 ,sheet_data.presentVal[index][i-1]) #插入值
    
RowSerNum = 0  #新增列_每列的編號
ComboVar = {}  #存放Combobox變數及物件
StrVar = {}    #存放Combobox字串變數及值
EntryVar = {}  #存放Entry變數及物件
GetVar = {}    #存放Entry值及儲存變數
#新增列
def Add_NewRow():
    global RowSerNum
    global GetVar
    StrVar['CoB'+'Get'+str(RowSerNum)] = tk.StringVar
    ComboVar['CoB'+str(RowSerNum)] = ttk.Combobox(win, state='readonly',
    textvariable=StrVar['CoB'+'Get'+str(RowSerNum)], justify='center', font=zh_font)
    ComboVar['CoB'+str(RowSerNum)]['values'] =  sheet_data.inspect
    ComboVar['CoB'+str(RowSerNum)].current(RowSerNum)
    ComboVar['CoB'+str(RowSerNum)].grid(column=1 ,row=(RowSerNum*2+2) ,padx=6 ,ipady=10, ipadx=10,rowspan=2)
    ComboVar['CoB'+str(RowSerNum)].bind('<<ComboboxSelected>>',combobox_selected)
    ComboVar['CoB'+str(RowSerNum)].RSNum = RowSerNum #自訂義方法儲存編號
    for Col in range(1,6):
        GetVar['Etr'+str(RowSerNum)+'_'+str(Col)+'Get'] = tk.StringVar
        EntryVar['Etr'+str(RowSerNum)+'_'+str(Col)] = tk.Entry(win , font=zh_font,
        textvariable = GetVar['Etr'+str(RowSerNum)+'_'+str(Col)+'Get'])
        EntryVar['Etr'+str(RowSerNum)+'_'+str(Col)].grid(column=Col+1 ,row=(RowSerNum*2+2) ,padx=6 ,ipady=10 ,ipadx=10,rowspan=2)
        EntryVar['Etr'+str(RowSerNum)+'_'+str(Col)].config(width=16)
        EntryVar['Etr'+str(RowSerNum)+'_'+str(Col)].insert(0 ,sheet_data.presentVal[RowSerNum][Col-1]) #利用插入方法，設定預設值(第幾次元,插入文字)
        
    RowSerNum += 1
    return RowSerNum

for i in range(7):
    Add_NewRow()
    print(i)

# #建立圖片物件
# img1 = tk.PhotoImage(file="wordicon.png")
# img2 = tk.PhotoImage(file="pdficon.png")
# #按鈕>>轉成word檔
# btn_word = tk.Button(text="轉成word檔")
# btn_word.grid(column=0,row=6,padx=60,pady=60)
# btn_word.config(image=img1)

# #按鈕>>轉成execel檔
# btn_pdf = tk.Button(text="轉成excel檔")
# btn_pdf.grid(column=1,row=6,padx=60,pady=60)
# btn_pdf.config(image=img2)


win.mainloop()

