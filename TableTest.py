import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom
from Table import Table
import AddStyle
from Utildom import Util

dir = 'C:\\Users\\97498\\Desktop\\论文格式'




# def unzip():
#     os.chdir(dir)
#     f = zipfile.ZipFile("肖露露毕业论文.docx")  # 打开需要修改的docx文件
#     f.extractall('./tabletest')  # 提取要修改的docx文件里的所有文件到workfolder文件夹
#     f.close()
#
# unzip()
# doc = minidom.parse(dir + '/tabletest/word/document.xml')
# styles = minidom.parse(dir + '/tabletest/word/styles.xml')
# themes = minidom.parse(dir + '/tabletest/word/theme/theme1.xml')
# Table.doc = doc
# Table.styles = styles
Util.setUp()
Table.setup()
cnt = 0
for tbl in Table.tables_with_borders:
    print(Table.isThreeLineTable(tbl))
    cnt += 1
    text = ''
    for t in tbl.getElementsByTagName('w:t'):
        text += t.childNodes[0].data
    print(text)
# print(cnt)