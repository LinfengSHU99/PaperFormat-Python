import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom

import AddStyle
import Utildom

Utildom.unzip()
os.chdir('workfolder')  # 修改当前工作目录到workfolder


doc = minidom.parse('./word/document.xml')
#styles = minidom.parse('C:\\Users\\97498\\Desktop\\论文格式\\workfolder\\word\\styles.xml')
paragraphlist = []
document = doc.firstChild
body = document.firstChild


for subnode in body.childNodes:
    if subnode.tagName == 'w:p':
        paragraphlist.append(subnode)


for p in paragraphlist:
    #print(Utildom.getFullText(p))
    if Utildom.isTitle(p):
        typeOfTitle = Utildom.titleType(p)
        Utildom.modifyTitle(p, typeOfTitle, doc)
        print(Utildom.getFullText(p))



with open(file='./word/document.xml', mode='w', encoding='utf-8') as f:  #解码设置为utf-8
    doc.writexml(f,  encoding="utf-8")
    print('ok')



AddStyle.r()
styles = minidom.parse('C:\\Users\\97498\\Desktop\\论文格式\\workfolder\\word\\styles.xml')
doc = minidom.parse('./word/document.xml')
paragraphlist = []
document = doc.firstChild
body = document.firstChild


for subnode in body.childNodes:
    if subnode.tagName == 'w:p':
        paragraphlist.append(subnode)

with open('C:\\Users\\97498\\Desktop\\论文格式\\格式.txt', 'w', encoding='utf-8') as ftxt:
    for p in paragraphlist:
        #print(Utildom.getFullText(p))

        for run in p.getElementsByTagName('w:r'):
            if run.getElementsByTagName('w:t'):
                text = run.getElementsByTagName('w:t')[0].childNodes[0].data
            ftxt.write(str(Utildom.getProperties(run, styles)) + '    ' + text + '\n' )

Utildom.zipToDocx()
os.chdir('../')  # 切换工作目录至原来的路径