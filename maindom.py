import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom
import FormatMatch
import AddStyle
from Utildom import Util

# dir = 'C:\\Users\\S\\Desktop\\论文格式'
#dir = 'C:\\Users\\97498\\Desktop\\论文格式'
dir = '/home/s/桌面/paperformat/PaperFormat-Python'
Util.dir = dir
Util.unzip()
os.chdir('workfolder')  # 修改当前工作目录到workfolder
AddStyle.r()

doc = minidom.parse('./word/document.xml')
styles = minidom.parse(dir + '/workfolder/word/styles.xml')
themes = minidom.parse(dir + '/workfolder/word/theme/theme1.xml')
numbering = minidom.parse(dir + '/workfolder/word/numbering.xml')
Util.doc = doc
Util.styles = styles
Util.themes = themes
Util.numbering = numbering
Util.setUp()
paragraphlist = []
document = doc.firstChild
body = document.firstChild

# for subnode in body.childNodes:
#     if subnode.tagName == 'w:p':
#         paragraphlist.append(subnode)

# 输出所有目录之后，判断为标题的段落
para_after_content = Util.getChildNodesOfNormalText()
print('\n------Title------\n')
for p in para_after_content:
    if Util.isTitle(p):
        typeOfTitle = Util.titleType(p)
        """修改标题格式，之后再考虑"""
        # Util.modifyTitle(p, typeOfTitle, doc)
        print(Util.getFullText(p) + '  titletype = '+str(Util.titleType(p)+1))


# with open(file='./word/document.xml', mode='w', encoding='utf-8') as f:  # 解码设置为utf-8
#     doc.writexml(f, encoding="utf-8")
#     print('ok')
#
#
# with open(dir +'/段落格式.txt', 'w', encoding='utf-8') as pftxt:
#     for p in paragraphlist:
#        pftxt.write(str(Util.getParagraphProperties(p)) + '    ' + Util.getFullText(p) + '\n\n')
#        #  if p.getElementsByTagName('w:pPr'):
#        #      print('true')
#        #  else:
#        #      print('false')
#
# with open(dir + '\\格式.txt', 'w', encoding='utf-8') as ftxt:
#     for p in paragraphlist:
#         # print(Utildom.getFullText(p))
#
#         for run in p.getElementsByTagName('w:r'):
#             if run.getElementsByTagName('w:t'):
#                 text = run.getElementsByTagName('w:t')[0].childNodes[0].data
#                 ftxt.write(str(Util.getRunProperties(run)) + '    ' + text + '\n')
print('\n-------Content------\n')
for t in Util.content_text_list:
    print(t)
# for p in Util.getChildNodesOfNormalText():
#     print(Util.getFullText(p))
FormatMatch.matchNormal(doc)
Util.zipToDocx()
os.chdir('../')  # 切换工作目录至原来的路径
