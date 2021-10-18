import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom

dir = 'C:\\Users\\S\\Desktop\\论文格式'


# dir = 'C:\\Users\\97498\\Desktop\\论文格式'

def unzip():
    os.chdir(dir)
    f = zipfile.ZipFile("肖露露毕业论文.docx")  # 打开需要修改的docx文件
    f.extractall('./workfolder')  # 提取要修改的docx文件里的所有文件到workfolder文件夹
    f.close()


def zipToDocx():
    newf = zipfile.ZipFile('../new肖露露毕业论文.docx', 'w')  # 创建一个新的docx文件，作为修改后的docx
    for root, dirs, files in os.walk('./'):  # 将workfolder文件夹所有的文件压缩至new.docx

        for file in files:
            # root = os.path.relpath(root[,'workfolder'])
            dir = root + '/' + file
            newf.write(dir)

    newf.close()


def getFullText(p) -> str:
    text = ''
    for t in p.getElementsByTagName('w:t'):
        text += t.childNodes[0].data
    return text


def isTitle(p) -> bool:
    text = getFullText(p)
    if len(p.getElementsByTagName('w:hyperlink')) != 0:  # 如果段落中包含超链接则判断其不为标题（可能为目录）
        return False

    reg = ["[1-9][0-9]*[\s]+?", "[1-9][0-9]*[.．、\u4E00-\u9FA5]", "[1-9][0-9]*\.[1-9][0-9]*",
           "[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*"]
    for i in range(len(reg)):
        reg[i] = re.compile(reg[i])

    i = 0
    match = False
    if len(text) < 40:
        for i in reversed(range(len(reg))):
            if re.match(reg[i], text):
                match = True
                break

    return match


def titleType(p) -> int:
    text = getFullText(p)
    reg = ["[1-9]", "[1-9][0-9]*\.[1-9][0-9]*", "[1-9][0-9]*\.[1-9][0-9]*\..+?"]
    for i in range(len(reg)):
        reg[i] = re.compile(reg[i])

    type = 0
    for type in reversed(range(len(reg))):
        if re.match(reg[type], text):
            return type


def modifyTitle(p, type, doc):
    for t in p.getElementsByTagName('w:t'):
        r = t.parentNode
        rpr = doc.createElement('w:rPr')
        rStyle = doc.createElement('w:rStyle')
        rStyle.setAttribute('w:val', 'heading1')
        rpr.appendChild(rStyle)
        if len(r.getElementsByTagName('w:rPr')) != 0:
            r.removeChild(r.getElementsByTagName('w:rPr')[0])
        if type != 0:
            node = doc.createElement('w:sz')
            node.setAttribute('w:val', '28')
            rpr.appendChild(node)
            node = doc.createElement('w:szCs')
            node.setAttribute('w:val', '28')
            rpr.appendChild(node)
        r.insertBefore(rpr, t)


# 列表为字体和字号
def getProperties(run, styles) -> dict:
    propertydict = {'eastAsia': None, 'ascii': None, 'sz': None, 'szCs': None}
    getFonts(run, propertydict, styles)
    return propertydict


def getFonts(run, propertydict, styles):
    def addFontToPropertydict(propertydict, font):
        propertydict['eastAsia'] = font.getAttribute('w:eastAsia') \
            if font.getAttribute('w:eastAsia') != '' and propertydict['eastAsia'] is None else \
            propertydict['eastAsia']
        propertydict['ascii'] = font.getAttribute('w:ascii') \
            if font.getAttribute('w:ascii') != '' and propertydict['ascii'] is None else \
            propertydict['ascii']
        if font.getAttribute('w:eastAsiaTheme'):
            propertydict['eastAsia'] = '等线 Light' if font.getAttribute('w:eastAsiaTheme')[0:5] == 'major' else '等线'
        if font.getAttribute('w:asciiTheme'):
            propertydict['ascii'] = '等线 Light' if font.getAttribute('w:asciiTheme')[0:5] == 'major' else '等线'

    def searchInStyle(propertydict, styles, styleId):
        pre_style = None
        for style in styles.getElementsByTagName('w:style'):
            if style.getAttribute('w:styleId') == styleId:
                rfonts = style.getElementsByTagName('w:rFonts')
                if len(rfonts) != 0:
                    font = rfonts[0]
                    addFontToPropertydict(propertydict, font)
                return style

    # 在rPr中查找
    rfonts = run.getElementsByTagName('w:rFonts')
    if len(rfonts) != 0:
        font = rfonts[0]
        addFontToPropertydict(propertydict, font)

    # 在RunStyle中查找
    if propertydict['eastAsia'] is None or propertydict['ascii'] is None:

        rstyle = run.getElementsByTagName('w:rStyle')
        if len(rstyle) != 0:
            rsid = rstyle[0].getAttribute('w:val')

            # 循环查询，以查询指定rStyle的父Style
            while rsid != '':
                pre_style = searchInStyle(propertydict, styles, rsid)
                rsid = ''
                if propertydict['eastAsia'] is None or propertydict['ascii'] is None:
                    # 递归在basedOn style中查找
                    if len(pre_style.getElementsByTagName('w:basedOn')) != 0:
                        rsid = pre_style.getElementsByTagName('w:basedOn')[0].getAttribute('w:val')

    # 在pStyle中查找
    if propertydict['eastAsia'] is None or propertydict['ascii'] is None:
        p = run.parentNode
        if len(p.getElementsByTagName('w:pPr')) != 0:
            ppr = p.getElementsByTagName('w:pPr')[0]
            if len(ppr.getElementsByTagName('w:pStyle')) != 0:
                pstyle = ppr.getElementsByTagName('w:pStyle')[0]
                psid = pstyle.getAttribute('w:val')

                # 循环查询，以查询指定pStyle的父Style
                while psid != '':
                    pre_style = searchInStyle(propertydict, styles, psid)
                    psid = ''
                    if propertydict['eastAsia'] is None or propertydict['ascii'] is None:
                        # 递归在basedOn style中查找
                        if len(pre_style.getElementsByTagName('w:basedOn')) != 0:
                            psid = pre_style.getElementsByTagName('w:basedOn')[0].getAttribute('w:val')

    # 查询default pStyle
    if propertydict['eastAsia'] is None or propertydict['ascii'] is None:

        for style in styles.getElementsByTagName('w:style'):
            if style.getAttribute('w:type') == 'paragraph' and style.getAttribute('w:default') == '1':
                rfonts = style.getElementsByTagName('w:rFonts')
                if len(rfonts) != 0:
                    font = rfonts[0]
                    addFontToPropertydict(propertydict, font)
                break

    if propertydict['eastAsia'] is None or propertydict['ascii'] is None:

        rprdefault = styles.getElementsByTagName('w:rPrDefault')
        if len(rprdefault) != 0:
            rfonts = rprdefault[0].getElementsByTagName('w:rFonts')
            if len(rfonts) != 0:
                font = rfonts[0]
                addFontToPropertydict(propertydict, font)
