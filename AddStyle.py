import zipfile
import os
import re
from lxml import etree
import sys
from Utildom import Util
from xml.dom import minidom
def r():
    # w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    # #os.chdir('workfolder')
    #
    # headingStyleOne = etree.Element(w + 'style', attrib={w + 'type': 'character', w + 'styleId': 'heading1'})
    # headingStyleOne.append(etree.Element(w + 'name', attrib={w + 'val': 'myHeading1'}))
    # rPr1 = etree.Element(w + 'rPr')
    # rPr1.append(etree.Element(w + 'sz', attrib={w + 'val': '32'}))
    # rPr1.append(etree.Element(w + 'szCs', attrib={w + 'val': '32'}))
    # rPr1.append(etree.Element(w + 'rFonts', attrib={w + 'ascii': 'Times New Roman', w + 'eastAsia': '黑体'}))
    # headingStyleOne.append(rPr1)
    #
    #
    # treestyle = etree.ElementTree(file='./word/styles.xml')
    # styles = treestyle.getroot()
    # styles.append(headingStyleOne)
    #
    #
    #
    # treestyle.write('./word/styles.xml', encoding='utf-8')
    # Util.styles = minidom.parse(Util.dir + '/workfolder/word/styles.xml')

    styles = Util.styles


    style_tbl_title_error = styles.createElement('w:style')
    style_tbl_title_error.setAttribute('w:type', 'paragraph')
    style_tbl_title_error.setAttribute('w:styleId', 'error')
    name = styles.createElement('w:name')
    name.setAttribute('w:val', 'error')
    style_tbl_title_error.appendChild(name)
    ppr = styles.createElement('w:pPr')
    jc = styles.createElement('w:jc')
    jc.setAttribute('w:val', 'center')
    ppr.appendChild(jc)
    rpr = styles.createElement('w:rPr')
    sz = styles.createElement('w:sz')
    sz.setAttribute('w:val', '21')
    color = styles.createElement('w:color')
    color.setAttribute('w:val', 'red')
    rpr.appendChild(color)
    rpr.appendChild(sz)
    style_tbl_title_error.appendChild(ppr)
    style_tbl_title_error.appendChild(rpr)
    styles.childNodes[0].appendChild(style_tbl_title_error)





    with open(file='./word/styles.xml', mode='w', encoding='utf-8') as f:  # 解码设置为utf-8
        styles.writexml(f, encoding="utf-8")