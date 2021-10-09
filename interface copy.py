import tkinter as tk
import tkinter.ttk as ttk
from typing import ValuesView
import sheet_data
import time


win = tk.Tk()                   #定義視窗
win.title('簡單輸入')            #視窗名稱
win.geometry('1000x800')        #視窗大小
win.resizable(0,0)              #視窗大小固定
win.config(bg="#525252")        #視窗背景
win.attributes("-alpha",0.9)    #透明度
win.attributes("-topmost",0)    #視窗置頂
#grid控制變數
# label_wight = 130        #對應padx(x方向填空)
# label_wight_inner = 0    #對應ipadx(內部x方向填空)
# label_hight_inner = 0    #對應ipady(內部x方向填空)

#設定選項
zh_font = 'Regular 16'     #中文字體 大小 #Regular標楷體
label_pady = 20            #label y填空
###

##標題欄位建立
col_name = ['檢驗項目','檢驗方法' ,'物性指標' ,'檢驗結果' ,'單位'
,'檢驗人']

for col_index in range(0,6):
    lbl_1 = tk.Label(win, text=col_name[col_index], font='Regular 16')
    lbl_1.grid(column=col_index, row=0, padx=8 , pady=label_pady)
#設定選項
entry_hight = 5           #對應pady(y方向填空)
randomNum = 0

#標題變數名稱
comboList = ['combo1', 'combo2', 'combo3', 'combo4', 'combo5' ]
#文字變數名稱
textList = ['coboText1', 'coboText2', 'coboText3', 'coboText4', 'coboText5' ]

b = 0 
def add_cobobox(var,b):
    print('start :',b)
    textVar = tk.StringVar
    globals()[var] = ttk.Combobox(win,textvariable=textVar, state='readonly')
    globals()[var]['values'] = sheet_data.inspect
    globals()[var].current(b)
    globals()[var].grid(column=0 ,row=(b+1) ,padx=6 ,ipady=entry_hight)
    return print(b)






for i in comboList:        #建立欄
        add_cobobox(i,b)
        b += 1
       
win.mainloop()

