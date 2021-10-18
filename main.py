import zipfile
import os
import re
from lxml import etree
import sys

import AddStyle
import Util

dir = 'C:\\Users\\S\\Desktop\\论文格式'
#dir = 'C:\\Users\\97498\\Desktop\\论文格式'

os.chdir(dir)
f = zipfile.ZipFile("肖露露毕业论文.docx")  # 打开需要修改的docx文件
f.extractall('./workfolder')  # 提取要修改的docx文件里的所有文件到workfolder文件夹
f.close()
newf = zipfile.ZipFile('new肖露露毕业论文.docx', 'w')  # 创建一个新的docx文件，作为修改后的docx
os.chdir('workfolder')  # 修改当前工作目录到workfolder
#file = open(file='../标题.txt', mode='w')
tree = etree.ElementTree(file='./word/document.xml')
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
paralist = []
document = tree.getroot()
body = document[0]

for p in body:
    if p.tag == w + 'p':  # 避免添加表格中的文字
        paralist.append(p)

for p in paralist:      # 检测p段落是否为标题，如果是标题则修改其格式为标题所需格式
    if Util.isTitle(p):
        typeOfTitle = Util.titleType(p)
        Util.modifyTitle(p, typeOfTitle)
        print(Util.getFullText(p))

tree.write('./word/document.xml', encoding='utf-8')

AddStyle.r()


'''省略执行修改的程序'''

for root, dirs, files in os.walk('./'):  #将workfolder文件夹所有的文件压缩至new.docx
    print(root)
    for file in files:
        #root = os.path.relpath(root[,'workfolder'])
        dir = root + '/' + file
        newf.write(dir)


newf.close()
# os.chdir('../')   #切换工作目录至原来的路径
os.chdir('../')  # 切换工作目录至原来的路径