import zipfile
import os
import re
from lxml import etree
from xml.dom import minidom

w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def getFullText(para) -> str:
    text = ''
    for t in para.iter(tag=w + 't'):
        text += t.text
    return text


def isTitle(para) -> bool:
    text = getFullText(para)
    if len(list(para.iter(tag = w + 'hyperlink'))) != 0:    #如果段落中包含超链接则判断其不为标题（可能为目录）
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

def titleType(para) -> int:
    text = getFullText(para)
    reg = ["[1-9]", "[1-9][0-9]*\.[1-9][0-9]*", "[1-9][0-9]*\.[1-9][0-9]*\..+?"]
    for i in range(len(reg)):
        reg[i] = re.compile(reg[i])

    type = 0
    for type in reversed(range(len(reg))):
        if re.match(reg[type], text):
            return type


def modifyTitle(para, typeOfTitle):
    for t in para.iter(tag=w + 't'):
        run = t.getparent()
        if len(list(run.iter(tag=w + 'rPr'))) != 0:
            run.remove(next(run.iter(tag=w + 'rPr')))
        rPr = etree.Element(w + 'rPr')
        rPr.append(etree.Element(w + 'rStyle', attrib={w + 'val': 'heading1'}))
        if typeOfTitle != 0:
            rPr.append(etree.Element(w + 'sz', attrib={w + 'val': '28'}))
            rPr.append(etree.Element(w + 'szCs', attrib={w + 'val': '28'}))

        run.insert(0, rPr)      #必须将<w:rPr>加在<w:t>之前

        # if match:
        #     if i == 2:
        #         file.write(text + '  三级标题\n')
        #     elif i == 1:
        #         file.write(text + '  二级标题\n')
        #     else:
        #         file.write(text + '  一级标题\n')
