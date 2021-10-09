import tkinter as tk
import tkinter.ttk as ttk
from typing import ValuesView
import sheet_data


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

#標題欄位建立
for col_index in range(0,6):
    lbl_1 = tk.Label(win, text=sheet_data.col_name[col_index], font='Regular 16')
    lbl_1.grid(column=col_index, row=0, padx=8 , pady=label_pady)

#輸入列建立
# for row_index in list(range(0,8)):
#     for entry_col_index in range(1,6):
#         input_text = tk.Entry(win ,font='Regular 10')
#         input_text.grid(column=entry_col_index ,row=(row_index+1) ,padx=6 ,ipady=5)
#         input_text.config(width=16)
#         insert_text = sheet_data.preset_value[(entry_col_index-1)] #將文字取出
#         input_text.insert(0 ,sheet_data.preset_value[entry_col_index-1]) #利用插入方法，設定預設值(第幾次元,插入文字)
#         var = input_text.get()

#Combobox被選擇後的行為
def combobox_selected(event):
    print(event.widget) #範例
    print('entry：','good')
    
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
    textvariable=StrVar['CoB'+'Get'+str(RowSerNum)])
    ComboVar['CoB'+str(RowSerNum)]['values'] =  sheet_data.inspect
    ComboVar['CoB'+str(RowSerNum)].current(RowSerNum)
    ComboVar['CoB'+str(RowSerNum)].grid(column=0 ,row=(RowSerNum+1) ,padx=6 ,ipady=5)
    ComboVar['CoB'+str(RowSerNum)].bind('<<ComboboxSelected>>',combobox_selected)
    for i in range(1,6):
        EntryVar['Etr'] = tk.Entry(win ,font='Regular 10')
        EntryVar['Etr'].grid(column=i ,row=(RowSerNum+1) ,padx=6 ,ipady=5)
        EntryVar['Etr'].config(width=16)
        EntryVar['Etr'].insert(0 ,sheet_data.present_value[i-1]) #利用插入方法，設定預設值(第幾次元,插入文字)
        GetVar['Etr'+'Get'+str(i)] = EntryVar.get
        
    RowSerNum += 1
    return RowSerNum
#目標：幫entry獲取的值設計變數。修正變數名稱

for i in range(7):
    Add_NewRow()
print(ComboVar,StrVar,EntryVar,GetVar,sep='\n')






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

