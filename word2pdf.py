"""
wordpath: 根目录，其下可包含多级子目录。
功能：将根目录下找到所有 .doc/.docx 文件替换成 .pdf 文件（默认将删除原文件）。

"""

import os
import re
from tqdm import tqdm
from win32com import client

def doc2pdf(wordpath):
    for root, dirs, files in os.walk(wordpath, topdown=False):
        ##后缀名为 doc,docx
        for name in tqdm([f for f in files if f.split('.')[-1] in ['docx', 'doc']]):
            doc_name = os.path.join(root, name)
            pdf_name = os.path.join(root, re.sub('\.(doc|docx)$', ".pdf", name))
            try:
                word = client.DispatchEx('Word.Application')
                worddoc = word.Documents.Open(doc_name, ReadOnly = 1)
                worddoc.SaveAs(pdf_name, FileFormat = 17)
                worddoc.Close(True)
                word.Quit()
                ##删除源文件
                os.remove(doc_name)
            except Exception as e:
                print(e)
                return
                
wordpath = "C:\\Users\\lauer\\Desktop\\word_root"
doc2pdf(wordpath)
