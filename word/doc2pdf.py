import os
from win32com import client

path=r'C:\Users\jyh\PycharmProjects\pythonProject\gw'
word = client.Dispatch("Word.Application")  # 打开word应用程序

# 转换docx为pdf
def docx2pdf(fn):
    try:
        doc = word.Documents.Open(fn)  # 打开word文件
        doc.SaveAs("{}.pdf".format(fn[:-5]), 17)  # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    except:
        print(f"{fn}--转换失败")
    doc.Close()  # 关闭原来word文件

listdir= os.listdir(path)
for f in listdir:
    file_path=os.path.join(path,f)
    docx2pdf(file_path)
print("ok")
word.Quit()