##讀取word檔內容
# import docx 

# doc = docx.Document('doc/AI 技術可以讓隱藏於暗處的物品現形.docx')

# print('段落數量： ', len(doc.paragraphs))

# for para in doc.paragraphs:
#     print(para.text)
# print('para : ',type(para))
# print('para.text : ',type(para.text))
# print('doc.paragraphs : ',type(doc.paragraphs))
# print('doc : ',type(doc))

##尋找模組路徑
# import sys
# import docx
# print(sys.modules['docx'])
##結束

##讀取每個run的內容
# import docx 

# doc = docx.Document('doc/AI 技術可以讓隱藏於暗處的物品現形.docx')

# para = doc.paragraphs[4]
# print(para.text + '\n')
# print('run數量： ', len(para.runs))

# for i in range(0, len(para.runs)):
#     print(i, para.runs[i].text)
## 結束

## 寫入word 檔

#!/usr/bin/python
#-*- coding: utf-8 -*-
# import docx

# doc = docx.Document()

# doc.add_heading('AI 技術可以讓隱藏於暗處的物品現形', level=1)
# doc.add_heading('資料來源', level=2)
# doc.add_paragraph('Engadget中文版')
# doc.add_heading('內文', level=2)
# text1 = "在幾乎完全沒有光線的環境中，
# 相信任何人都無法清楚看到大部分的物體，但對 AI 來說或許不是件難事。麻省理工學院的科學家近來開發出一種技術，
# 能透過深度神經網路在幾乎沒有光線的環境下，看見其中的物體。為了讓神經網路培養出這樣的能力，該團隊利用了一萬張黑暗、
# 充滿噪點，甚至是沒有對焦的圖片，搭配上在其中存在的物品的圖片，對其進行訓練。而這樣的方式，
# 不僅讓神經網路了解應該辨識出什麼物品，也能訓練其在微弱的光源下該如何從畫面中凸顯出該辨識的目標。除此之外，
# 研究人員也給神經網路上了一堂物理課，讓它理解沒有對焦的相機是如何產生出模糊的圖片，藉此來應付這類情況下的畫面。"
# doc.add_paragraph(text1)

# doc.save('doc/AI 技術可以讓隱藏於暗處的物品現形.docx')
##結束

## 插入表格
# import docx
# from docx.shared import Cm  #加入可調整的 word 單位

# doc = docx.Document('doc/AI 技術可以讓隱藏於暗處的物品現形.docx')

# records = (
#     ('amos', '12345678', 'teacher'),
#     ('carol', '23456789', 'student'),
#     ('frank', '34567890', 'engineer')
# )


# table = doc.add_table(rows=1, cols=3)

# hdr_cells = table.rows[0].cells
# hdr_cells[0].text = '姓名'
# hdr_cells[1].text = '電話'
# hdr_cells[2].text = '職稱'


# for name, tel, title in records:
#     row_cells = table.add_row().cells
#     row_cells[0].text = name
#     row_cells[1].text = tel
#     row_cells[2].text = title


# doc.save('doc/AI 技術可以讓隱藏於暗處的物品現形.docx')

## 修改表格裡面的內容
# import docx

# doc = docx.Document('doc/AI 技術可以讓隱藏於暗處的物品現形.docx')
# tables = doc.tables
# tables[0].cell(1,0).text="貓糧"
# tables[0].cell(2,0).text="貓糧"
# tables[0].cell(3,0).text="貓糧"
# print(tables[0].cell(3,0).text)
# doc.save('doc/AI 技術可以讓隱藏於暗處的物品現形.docx')

##結束

## 解word範例的格式、修改表格內容

# import docx
# import sheet_data
# doc = docx.Document('doc/慶育油品範例.docx')
# tables = doc.tables
# for i in range(1,len(tables[0].rows)):
#     for j in range(1,len(tables[0].columns)):
#         # tables[0].cell(i,j).text="貓糧"
#         print(i,j)
#印出每格文字
# for col_index , n in enumerate(tables[0].rows):
#     row_index = 0 
#     for cell in n.cells:
#         row_index += 1
#         print(col_index,cell.text,'(',row_index,')')
# #找尋x座標
# coord_X = 0  
# for i , n in enumerate(tables[0].rows):
#     for cell in n.cells:
#         if cell.text == sheet_data.col_name[0]:
#             coord_X = i + 1
#             print(cell.text)
# #檢查x座標
# print(coord_X,tables[0].cell(coord_X,0).text,sep='\n')
  
#目標 1.讀取表格每格內容 2.找尋標題 3.修改每格內容
# doc.save('doc/scan.docx')
#結果: 資料寬度：2 1 3 2 2 1

#word表格值測試_關鍵覆蓋值
# import docx
# import sheet_data
# doc = docx.Document('doc/慶育油品範例.docx')
# tables = doc.tables
# tables[0].cell(7,0).text="貓糧0"
# tables[0].cell(7,1).text="貓糧1"
# tables[0].cell(7,2).text="貓糧2"
# tables[0].cell(7,3).text="貓糧3"
# tables[0].cell(7,4).text="貓糧4"
# tables[0].cell(7,5).text="貓糧5"
# tables[0].cell(7,6).text="貓糧6"
# tables[0].cell(7,7).text="貓糧7"
# tables[0].cell(7,8).text="貓糧8"
# tables[0].cell(7,9).text="貓糧9"
# tables[0].cell(7,10).text="貓糧10"

# doc.save('doc/scan.docx')

#結論: 最後一值(1 2 5 7 9 10)

#word表格值測試_1347910
# import docx
# import sheet_data
# doc = docx.Document('doc/慶育油品範例.docx')
# tables = doc.tables

# tables[0].cell(7,1).text="取代"
# tables[0].cell(7,0).text="貓糧1"
# tables[0].cell(7,2).text="貓糧3"
# tables[0].cell(7,3).text="貓糧4"
# # tables[0].cell(7,4).text="貓糧5"
# # tables[0].cell(7,5).text="貓糧6"
# tables[0].cell(7,6).text="貓糧7"
# # tables[0].cell(7,7).text="貓糧8"
# tables[0].cell(7,8).text="貓糧9"
# tables[0].cell(7,10).text="貓糧10"

# doc.save('doc/scan.docx')

#結果: 1.改變其中一值便能修改表格值
