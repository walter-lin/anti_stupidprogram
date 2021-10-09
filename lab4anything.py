###讀取word檔內容
# import docx

# doc = docx.Document('./doc/exampleWord.docx')
# print('段落數量： ', len(doc.paragraphs))

# for para in doc.paragraphs:
#     print(para,para.text)
###結果：成功

###寫入word檔
# import docx

# doc = docx.Document()

# doc.add_heading('AI 技術可以讓隱藏於暗處的物品現形', level=1)
# doc.add_heading('資料來源', level=2)
# doc.add_paragraph('Engadget中文版')
# doc.add_heading('內文', level=2)
# #para1 = "在幾乎完全沒有光線的環境中，相信任何人都無法清楚看到大部分的物體，但對 AI 來說或許不是件難事。麻省理工學院的科學家近來開發出一種技術，能透過深度神經網路在幾乎沒有光線的環境下，看見其中的物體。為了讓神經網路培養出這樣的能力，該團隊利用了一萬張黑暗、充滿噪點，甚至是沒有對焦的圖片，搭配上在其中存在的物品的圖片，對其進行訓練。而這樣的方式，不僅讓神經網路了解應該辨識出什麼物品，也能訓練其在微弱的光源下該如何從畫面中凸顯出該辨識的目標。除此之外，研究人員也給神經網路上了一堂物理課，讓它理解沒有對焦的相機是如何產生出模糊的圖片，藉此來應付這類情況下的畫面。"
# doc.add_paragraph(para1)

# doc.save('AI 技術可以讓隱藏於暗處的物品現形.docx')
# #測試動態變數
# for i in range(1,10):
#   globals()['number'+str(i)] = 2+i
#   print('Print In Once: ', globals()['number'+str(i)])

# print(globals()['number'+'1'])
# print(globals()['number1'])
# print(number1)
# print(type(number1))
###結果：發現讀取文字有字數限制，過長字串需指向一個變數可解決這個問題。

###變數建立方法1
# comboList = []
# print('資料生成中..')
# for row_index in list(range(1,9)):
#     comboList += ['mycombo'+str(row_index)]
#     globals()[comboList[row_index-1]] = row_index
#     comboList[row_index-1].bind('<<ComboboxSelected>>', combobox_selected()) #綁定事件
# print(comboList)
    
## 方法2_利用字典管理
# comboList = {}
# for row_index in list(range(1,9)):
#     comboList['mycombo'+str(row_index)] = 'a'+ str(row_index)
#     print(comboList)
#     print(comboList['mycombo'+str(row_index)])
# import tkinter as tk
# import tkinter.ttk as ttk
# from typing import ValuesView
# import sheet_data
###結果：成功

### 實驗
# win = tk.Tk()                   #定義視窗
# win.title('簡單輸入')            #視窗名稱
# win.geometry('1000x800')        #視窗大小
# win.resizable(0,0)              #視窗大小固定
# win.config(bg="#525252")        #視窗背景
# win.attributes("-alpha",0.9)    #透明度
# win.attributes("-topmost",0)    #視窗置頂

# #設定選項
# zh_font = 'Regular 16'     #中文字體 大小 #Regular標楷體
# label_pady = 20            #label y填空
# ###

# ##標題欄位建立
# col_name = ['檢驗項目','檢驗方法' ,'物性指標' ,'檢驗結果' ,'單位'
# ,'檢驗人']

# for col_index in range(0,6):
#     lbl_1 = tk.Label(win, text=col_name[col_index], font='Regular 16')
#     lbl_1.grid(column=col_index, row=0, padx=8 , pady=label_pady)
# #設定選項
# entry_hight = 5           #對應pady(y方向填空)
# randomNum = 0

# def combobox_selected1(event):
#     print(mycobo1.current(), mycobo1.get())

# def combobox_selected2(event):
#     print(mycobo2.current(), mycobo2.get())
# textVar = tk.StringVar
# mycobo1 = ttk.Combobox(win,textvariable=textVar, state='readonly')
# mycobo1['values'] = sheet_data.inspect
# mycobo1.current(0)
# mycobo1.grid(column=0 ,row=1 ,padx=6 ,ipady=8)
# mycobo1.bind('<<ComboboxSelected>>',combobox_selected1)

# textVar = tk.StringVar
# mycobo2 = ttk.Combobox(win,textvariable=textVar, state='readonly')
# mycobo2['values'] = sheet_data.inspect
# mycobo2.current(0)
# mycobo2.grid(column=0 ,row=2 ,padx=6 ,ipady=8)
# mycobo2.bind('<<ComboboxSelected>>',combobox_selected2)

# #輸入列建立
# for row_index in list(range(0,8)):
#     for entry_col_index in range(1,6):
#         input_text = tk.Entry(win ,font='Regular 10')
#         input_text.grid(column=entry_col_index ,row=(row_index+1) ,padx=6 ,ipady=entry_hight)
#         input_text.config(width=16)
#         insert_text = sheet_data.present_value[(entry_col_index-1)] #將文字取出
#         input_text.insert(0 ,sheet_data.present_value[entry_col_index-1]) #利用插入方法，設定預設值(第幾次元,插入文字)
#         var = input_text.get()

# win.mainloop()
###證明：textVar = tk.StringVar 重複沒差

# #實驗_函式_新增欄位
# import tkinter as tk
# import tkinter.ttk as ttk
# from typing import ValuesView
# import sheet_data

# win = tk.Tk()                   #定義視窗
# win.title('簡單輸入')            #視窗名稱
# win.geometry('1000x800')        #視窗大小
# win.resizable(0,0)              #視窗大小固定
# win.config(bg="#525252")        #視窗背景
# win.attributes("-alpha",0.9)    #透明度
# win.attributes("-topmost",0)    #視窗置頂

# #設定選項
# zh_font = 'Regular 16'     #中文字體 大小 #Regular標楷體
# label_pady = 20            #label y填空

# #標題欄位建立
# col_name = ['檢驗項目','檢驗方法' ,'物性指標' ,'檢驗結果' ,'單位'
# ,'檢驗人']

# for col_index in range(0,6):
#     lbl_1 = tk.Label(win, text=col_name[col_index], font='Regular 16')
#     lbl_1.grid(column=col_index, row=0, padx=8 , pady=label_pady)

# #輸入列建立
# for row_index in list(range(0,8)):
#     for entry_col_index in range(1,6):
#         input_text = tk.Entry(win ,font='Regular 10')
#         input_text.grid(column=entry_col_index ,row=(row_index+1) ,padx=6 ,ipady=5)
#         input_text.config(width=16)
#         insert_text = sheet_data.present_value[(entry_col_index-1)] #將文字取出
#         input_text.insert(0 ,sheet_data.present_value[entry_col_index-1]) #利用插入方法，設定預設值(第幾次元,插入文字)
#         var = input_text.get()
# ComSerNum = 0 #Combobox編號

# def combobox_selected(event):
#     print(event,event.widget,event.widget.current(),sep='\n') #可以用widget帶出對應的變數
#     print('2',event.widget.get(),sep='\n')


# def Add_cobobox(Var,StrVar):
#     global ComSerNum
#     ComSerNum += 1
#     StrVar = tk.StringVar
#     Var = ttk.Combobox(win,textvariable=StrVar, state='readonly')
#     Var['values'] =  sheet_data.inspect
#     Var.current(ComSerNum)
#     Var.grid(column=0 ,row=(ComSerNum+1) ,padx=6 ,ipady=5)
#     Var.bind('<<ComboboxSelected>>',combobox_selected)
    

# for i in range(5):
#     Add_cobobox(('a'+str(i)),i)
    

# win.mainloop()
##小結果：透過全域變數能將值帶出
##小結果：字典的Key的當作變數名稱，Valune能存放變數的值
##結論：(1).widget方法可以回傳對應變數(例如 .!cobobox2)
##     (2)使用lambda匿名函式可以增加傳入的參數

###實驗_區域變數
# def test(test):
#     global a
#     for i in range(1,3):
#         test = 1
#         a = test * i
#         print('test：',a)
#     print(a)
#     return(a)

# test(0)
# print(a)
##結論：1.一開始global宣告型態判斷不用理
##      2.有沒有return都可以傳出變數，但有無不良影響還待確認

#實驗_陣列空格是否會被輸出

# test = [1,2,3,]
# test2 = [1 , 2 ,3 ,' ' ]

# for i in range(3): 
#     print(test[i],sep='\n')
# for i in range(4):
#     print(test2[i],sep='\n')
## 結論：值得範圍是由逗號界定，但不會抓空值