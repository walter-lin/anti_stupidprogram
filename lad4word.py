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
##