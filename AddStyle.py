import zipfile
import os
import re
from lxml import etree
import sys
from Utildom import Util
from xml.dom import minidom
def r():
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    #os.chdir('workfolder')

    headingStyleOne = etree.Element(w + 'style', attrib={w + 'type': 'character', w + 'styleId': 'heading1'})
    headingStyleOne.append(etree.Element(w + 'name', attrib={w + 'val': 'myHeading1'}))
    rPr1 = etree.Element(w + 'rPr')
    rPr1.append(etree.Element(w + 'sz', attrib={w + 'val': '32'}))
    rPr1.append(etree.Element(w + 'szCs', attrib={w + 'val': '32'}))
    rPr1.append(etree.Element(w + 'rFonts', attrib={w + 'ascii': 'Times New Roman', w + 'eastAsia': '黑体'}))
    headingStyleOne.append(rPr1)


    treestyle = etree.ElementTree(file='./word/styles.xml')
    styles = treestyle.getroot()
    styles.append(headingStyleOne)



    treestyle.write('./word/styles.xml', encoding='utf-8')
    Util.styles = minidom.parse(Util.dir + '/workfolder/word/styles.xml')
