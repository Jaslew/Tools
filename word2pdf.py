"""
wordpath: 根目录，其下只包含多级子目录和 doc 文件。
将根目录下所有子目录里的 doc 文件替换成 pdf 文件（会删除原 doc 文件）。

"""

import os
from win32com import client

def doc2pdf(wordpath):
    for root, dirs, files in os.walk(wordpath, topdown=False):
        ##后缀名为 doc,docx
        for name in [f for f in files if f.split('.')[-1] in ['docx', 'doc']]:
            doc_name = os.path.join(root, name)
            pdf_name = os.path.join(root, name.split('.')[-2])+'.pdf'
            try:
                word = client.DispatchEx('Word.Application')
                worddoc = word.Documents.Open(doc_name, ReadOnly = 1)
                worddoc.SaveAs(pdf_name, FileFormat = 17)
                worddoc.Close(True)
                word.Quit()
                os.remove(doc_name)
            except Exception as e:
                print(e)
                print("error")
                return
                
wordpath = "C:\\Users\\lauer\\Desktop\\word_root"
doc2pdf(wordpath)
