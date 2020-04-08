"""
wordpath: 根目录，其下只包含多级子目录和 doc 文件。
将根目录下所有子目录里的 doc 文件替换成 pdf 文件（会删除原 doc 文件）。

"""

import os
from win32com import client

def doc2pdf(wordpath):
    for root, dirs, files in os.walk(wordpath, topdown=False):
        for name in files:
            doc_name = os.path.join(root, name)
            # pdf_name = doc_name.split('.docx',)[0] + '.pdf'
            (filename, extension) = os.path.splitext(name)#filename文件名，extension后缀名
            #pdf_name为pdf文件名，此处不加.pdf也可以，但是word名中有‘.’的时候会发生转化失败
            pdf_name = os.path.join(root, filename)+'.pdf'
            try:
                word = client.DispatchEx('Word.Application')
                if os.path.exists(pdf_name):
                    os.remove(pdf_name)
                worddoc = word.Documents.Open(doc_name, ReadOnly = 1)
                worddoc.SaveAs(pdf_name, FileFormat = 17)
                worddoc.Close(True)
                word.Quit()#切记，这步必须加，要不然线程不会杀死，电脑会卡死
                os.remove(doc_name)
            except Exception as e:
                print(e)
                print("error")
                return 1
                
wordpath = "C:\\Users\\lauer\\Desktop\\word_root"
doc2pdf(wordpath)
