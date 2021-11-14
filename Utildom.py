import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom


class Util:
    dir = 'C:\\Users\\97498\\Desktop\\论文格式'
    styles = None
    themes = None
    doc = None
    normal_run_property = {'eastAsia': "宋体", 'ascii': "Times New Roman", 'sz': "24", 'szCs': None, 'kern': None}
    # 目录字符串列表
    content_text_list = []

    # dir = 'C:\\Users\\S\\Desktop\\论文格式'


    # 一些准备性的工作，相当于初始化
    @classmethod
    def setUp(cls):
        # 将所有目录段落加入一个列表
        cls.getContent()

    @classmethod
    def unzip(cls):
        os.chdir(cls.dir)
        f = zipfile.ZipFile("张军毕业论文.docx")  # 打开需要修改的docx文件
        f.extractall('./workfolder')  # 提取要修改的docx文件里的所有文件到workfolder文件夹
        f.close()

    @classmethod
    def zipToDocx(cls):
        # 将修改后的document.xml写入
        with open(file='./word/document.xml', mode='w', encoding='utf-8') as f:  # 解码设置为utf-8
            cls.doc.writexml(f, encoding="utf-8")

        newf = zipfile.ZipFile(cls.dir + '/new张军毕业论文.docx', 'w')  # 创建一个新的docx文件，作为修改后的docx
        for root, dirs, files in os.walk('./'):  # 将workfolder文件夹所有的文件压缩至new.docx
            for file in files:
                # root = os.path.relpath(root[,'workfolder'])
                direction = root + '/' + file
                newf.write(direction)

        newf.close()

    @classmethod
    def getFullText(cls, p) -> str:
        text = ''
        for t in p.getElementsByTagName('w:t'):
            text += t.childNodes[0].data
        return text

    @classmethod
    def isTitle(cls, p) -> bool:
        text = cls.getFullText(p)
        # if len(p.getElementsByTagName('w:hyperlink')) != 0:  # 如果段落中包含超链接则判断其不为标题（可能为目录）
        #     return False

        reg = ["[1-9][0-9]*[\s]+?", "[1-9][0-9]*[.．、\u4E00-\u9FA5]", "[1-9][0-9]*\.[1-9][0-9]*",
               "[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*"]
        for i in range(len(reg)):
            reg[i] = re.compile(reg[i])

        match = False
        if len(text) < 30:
            for i in reversed(range(len(reg))):
                if re.match(reg[i], text):
                    # 若匹配的样式是1. xxxx的样式，则查找目录是否有相同的内容，如果有，则判断为标题。因为有些字数少的正文的编号项目也有可能匹配成功。
                    if i == 1:
                        for t in cls.content_text_list:
                            if re.sub('[0-9\.．\s]', '', text) in re.sub('[0-9\.．\s]', '', t):
                                match = True
                                break
                    else:
                        match = True
                    break
        return match

    @classmethod
    def titleType(cls, p) -> int:
        text = cls.getFullText(p)
        reg = ["[1-9]", "[1-9][0-9]*\.[1-9][0-9]*", "[1-9][0-9]*\.[1-9][0-9]*\..+?"]
        for i in range(len(reg)):
            reg[i] = re.compile(reg[i])

        type = 0
        for type in reversed(range(len(reg))):
            if re.match(reg[type], text):
                return type

    @classmethod
    def modifyTitle(cls, p, type, doc):

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

    # 获得目录的字符串列表，目前用于加强判断文章中的一级标题
    @classmethod
    def getContent(cls):
        begin = False
        lst = []
        for p in cls.doc.getElementsByTagName('w:p'):

            text = cls.getFullText(p)
            if text.replace(' ', '') == '目录':
                begin = True
                continue
            if begin:
                for t in lst:
                    if cls.getFullText(p).replace(' ', '') in t.replace(' ', ''):
                        begin = False
            if begin:
                lst.append(cls.getFullText(p))
            if begin is False and len(lst) > 0:
                break
        cls.content_text_list = lst

    # 判断某一段是否为目录
    @classmethod
    def isContent(cls, p) -> bool:
        # 当段落内含有hyperlink和fldChar时，该段落一定为目录的一部分（该方法只能判断目录的子集）
        if p.getElementsByTagName('w:hyperlink') and p.getElementsByTagName('w:fldChar'):
            return True
        text = cls.getFullText(p)
        # 如果某一段文字数小于30且开头有数字，或者摘要，附录，参考文献等字样，最后出现(xx) 或者xx （xx为数字）  等字样，则判断为目录
        if len(text) < 30 and (re.search(re.compile('^[1-9]+'), text) or re.search(re.compile('^摘(\s+)?要'), text) or
                               re.search(re.compile('^致(\s+)?谢'), text) or re.search(re.compile('^参考文献'), text) or
                               re.search(re.compile('^附(\s+)?录'), text) or re.search(
                             re.compile('^abstract', re.IGNORECASE), text)):
            if re.search(re.compile('[)）]$'), text):
                if re.search(re.compile('[(（][0-9]+[)）]$'), text):
                    return True
            else:
                if re.search(re.compile('[0-9]+$'), text):
                    return True

        return False

    # 返回一个body的子元素列表，该列表中的子元素都是位于目录之后的元素
    @classmethod
    def getChildNodesAfterContent(cls) -> list:
        childnode_list = cls.doc.childNodes[0].childNodes[0].childNodes
        location = 0
        begin = False
        for i in range(len(childnode_list)):
            if cls.isContent(childnode_list[i]):
                location = i
                begin = True

            elif begin:
                break
        return childnode_list[location + 1:]

    # @classmethod
    # 列表为字体和字号
    # def getProperties(cls, run, styles) -> dict:
    #     propertydict = {'eastAsia': None, 'ascii': None, 'sz': None, 'szCs': None, 'kern': None}
    #     cls.getRunProperties(run, propertydict)
    #     return propertydict

    # 获取一个Run里的所有与字有关(rPr)的里的格式，以字典返回。
    @classmethod
    def getRunProperties(cls, run) -> dict:
        propertydict = {'eastAsia': None, 'ascii': None, 'sz': None, 'szCs': None, 'kern': None}
        styles = cls.styles

        def searchFontInTheme(themes) -> dict:
            fontDict = {'majorEastAsia': None, 'minorEastAsia': None, 'majorHAnsi': None, 'minorHAnsi': None}
            majorFont = themes.getElementsByTagName('a:majorFont')[0]
            minorFont = themes.getElementsByTagName('a:minorFont')[0]
            for node in majorFont.childNodes:
                if node.tagName == 'a:latin':
                    fontDict['majorHAnsi'] = node.getAttribute('typeface')
                elif node.getAttribute('script') == 'Hans':
                    fontDict['majorEastAsia'] = node.getAttribute('typeface')
            for node in minorFont.childNodes:
                if node.tagName == 'a:latin':
                    fontDict['minorHAnsi'] = node.getAttribute('typeface')
                elif node.getAttribute('script') == 'Hans':
                    fontDict['minorEastAsia'] = node.getAttribute('typeface')
            return fontDict

        def addFontToPropertydict(propertydict, font, fontTheme):
            propertydict['eastAsia'] = font.getAttribute('w:eastAsia') \
                if font.getAttribute('w:eastAsia') != '' and propertydict['eastAsia'] is None else \
                propertydict['eastAsia']
            propertydict['ascii'] = font.getAttribute('w:ascii') \
                if font.getAttribute('w:ascii') != '' and propertydict['ascii'] is None else \
                propertydict['ascii']
            if font.getAttribute('w:eastAsiaTheme'):
                propertydict['eastAsia'] = fontTheme[font.getAttribute('w:eastAsiaTheme')]
            if font.getAttribute('w:asciiTheme'):
                propertydict['ascii'] = fontTheme[font.getAttribute('w:asciiTheme')]

        def addPropertyToPropertydict(propertydict, property, property_name):
            propertydict[property_name] = property.getAttribute('w:val') \
                if property.getAttribute('w:val') != '' and propertydict[property_name] is None else propertydict[
                property_name]

        # 在style中查找，并修改相关属性，返回查找到的style，以便于查找basedOn Style
        def searchInStyle(propertydict, styles, styleId):
            pre_style = None
            for style in styles.getElementsByTagName('w:style'):
                if style.getAttribute('w:styleId') == styleId:
                    rfonts = style.getElementsByTagName('w:rFonts')
                    if len(rfonts) != 0:
                        font = rfonts[0]
                        addFontToPropertydict(propertydict, font, fontTheme)

                    sz_list = style.getElementsByTagName('w:sz')
                    if sz_list:
                        sz = sz_list[0]
                        addPropertyToPropertydict(propertydict, sz, 'sz')

                    szCs_list = style.getElementsByTagName('w:szCs')
                    if szCs_list:
                        szCs = szCs_list[0]
                        addPropertyToPropertydict(propertydict, szCs, 'szCs')

                    kern_list = style.getElementsByTagName('w:kern')
                    if kern_list:
                        kern = kern_list[0]
                        addPropertyToPropertydict(propertydict, kern, 'kern')
                    return style

        fontTheme = searchFontInTheme(cls.themes)
        # 在rPr中查找
        rfonts = run.getElementsByTagName('w:rFonts')
        if len(rfonts) != 0:
            font = rfonts[0]
            addFontToPropertydict(propertydict, font, fontTheme)
        sz_list = run.getElementsByTagName('w:sz')
        if sz_list:
            sz = sz_list[0]
            addPropertyToPropertydict(propertydict, sz, 'sz')

        szCs_list = run.getElementsByTagName('w:szCs')
        if szCs_list:
            szCs = szCs_list[0]
            addPropertyToPropertydict(propertydict, szCs, 'szCs')

        kern_list = run.getElementsByTagName('w:kern')
        if kern_list:
            kern = kern_list[0]
            addPropertyToPropertydict(propertydict, kern, 'kern')

        # 在RunStyle中查找
        if None in propertydict.values():

            rstyle = run.getElementsByTagName('w:rStyle')
            if len(rstyle) != 0:
                rsid = rstyle[0].getAttribute('w:val')

                # 循环查询，以查询指定rStyle的父Style
                while rsid != '':
                    pre_style = searchInStyle(propertydict, styles, rsid)
                    rsid = ''
                    if None in propertydict.values():
                        # 递归在basedOn style中查找
                        if len(pre_style.getElementsByTagName('w:basedOn')) != 0:
                            rsid = pre_style.getElementsByTagName('w:basedOn')[0].getAttribute('w:val')

        # 在pStyle中查找
        if None in propertydict.values():
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
                        if None in propertydict.values():

                            if len(pre_style.getElementsByTagName('w:basedOn')) != 0:
                                psid = pre_style.getElementsByTagName('w:basedOn')[0].getAttribute('w:val')

        # 查询default pStyle
        if None in propertydict.values():
            for style in styles.getElementsByTagName('w:style'):
                if style.getAttribute('w:type') == 'paragraph' and style.getAttribute('w:default') == '1':
                    rfonts = style.getElementsByTagName('w:rFonts')
                    if len(rfonts) != 0:
                        font = rfonts[0]
                        addFontToPropertydict(propertydict, font, fontTheme)
                    sz_list = style.getElementsByTagName('w:sz')
                    if sz_list:
                        sz = sz_list[0]
                        addPropertyToPropertydict(propertydict, sz, 'sz')

                    szCs_list = style.getElementsByTagName('w:szCs')
                    if szCs_list:
                        szCs = szCs_list[0]
                        addPropertyToPropertydict(propertydict, szCs, 'szCs')

                    kern_list = style.getElementsByTagName('w:kern')
                    if kern_list:
                        kern = kern_list[0]
                        addPropertyToPropertydict(propertydict, kern, 'kern')
                    break

        if None in propertydict.values():
            rprdefault = styles.getElementsByTagName('w:rPrDefault')

            if len(rprdefault) != 0:
                rpr = rprdefault[0]
                rfonts = rpr.getElementsByTagName('w:rFonts')
                if len(rfonts) != 0:
                    font = rfonts[0]
                    addFontToPropertydict(propertydict, font, fontTheme)
                sz_list = rpr.getElementsByTagName('w:sz')
                if sz_list:
                    sz = sz_list[0]
                    addPropertyToPropertydict(propertydict, sz, 'sz')

                szCs_list = rpr.getElementsByTagName('w:szCs')
                if szCs_list:
                    szCs = szCs_list[0]
                    addPropertyToPropertydict(propertydict, szCs, 'szCs')

                kern_list = rpr.getElementsByTagName('w:kern')
                if kern_list:
                    kern = kern_list[0]
                    addPropertyToPropertydict(propertydict, kern, 'kern')

        return propertydict

    @classmethod
    def getParagraphProperties(cls, p) -> dict:
        propertydict = {'jc': None, 'line': None, 'lineRule': None, 'before': None, 'beforeLines': None, 'after': None,
                        'afterLines': None, 'firstLine': None, 'firstLineChars': None}
        styles = cls.styles

        def addIndentationToPropertydict(propertydict, ind):
            if propertydict['firstLine'] is None:
                propertydict['firstLine'] = ind.getAttribute('w:firstLine') if ind.getAttribute(
                    'w:firstLine') != '' else None
            if propertydict['firstLineChars'] is None:
                propertydict['firstLineChars'] = ind.getAttribute('w:firstLineChars') if ind.getAttribute(
                    'w:firstLineChars') != '' else None

        def addSpacingToPropertydict(propertydict, spacing):
            if propertydict['line'] is None:
                propertydict['line'] = spacing.getAttribute('w:line') if spacing.getAttribute('w:line') != '' else None
            if propertydict['lineRule'] is None:
                propertydict['lineRule'] = spacing.getAttribute('w:lineRule') if spacing.getAttribute(
                    'w:lineRule') != '' else None
            if propertydict['before'] is None:
                propertydict['before'] = spacing.getAttribute('w:before') if spacing.getAttribute(
                    'w:before') != '' else None
            if propertydict['beforeLines'] is None:
                propertydict['beforeLines'] = spacing.getAttribute('w:beforeLines') if spacing.getAttribute(
                    'w:beforeLines') != '' else None
            if propertydict['after'] is None:
                propertydict['after'] = spacing.getAttribute('w:after') if spacing.getAttribute(
                    'w:after') != '' else None
            if propertydict['afterLines'] is None:
                propertydict['afterLines'] = spacing.getAttribute('w:afterLines') if spacing.getAttribute(
                    'w:afterLines') != '' else None

        def searchInStyle(propertydict, styles, styleId):
            for style in styles.getElementsByTagName('w:style'):
                if style.getAttribute('w:styleId') == styleId:
                    ppr_list = style.getElementsByTagName('w:pPr')
                    if ppr_list:
                        ppr = ppr_list[0]
                        addProperties(propertydict, ppr)
                    return style

        def addProperties(propertydict, ppr):
            jc_list = ppr.getElementsByTagName('w:jc')
            ind_list = ppr.getElementsByTagName('w:ind')
            spacing_list = ppr.getElementsByTagName('w:spacing')
            if jc_list:
                propertydict['jc'] = jc_list[0].getAttribute('w:val') if propertydict['jc'] is None and \
                                                                         jc_list[0].getAttribute(
                                                                             'w:val') != '' else propertydict['jc']
            if ind_list:
                addIndentationToPropertydict(propertydict, ind_list[0])
            if spacing_list:
                addSpacingToPropertydict(propertydict, spacing_list[0])

        # 在pPr中查找

        ppr_list = p.getElementsByTagName('w:pPr')
        if ppr_list:
            ppr = ppr_list[0]
            addProperties(propertydict, ppr)

        # 在pStyle中查找
        if None in propertydict.values():
            pstyle_list = p.getElementsByTagName('w:pStyle')
            if pstyle_list:
                psid = pstyle_list[0].getAttribute('w:val')

                # 循环查询，以查询指定pStyle的父Style
                while psid != '':
                    pre_style = searchInStyle(propertydict, styles, psid)
                    psid = ''
                    if None in propertydict.values():

                        if len(pre_style.getElementsByTagName('w:basedOn')) != 0:
                            psid = pre_style.getElementsByTagName('w:basedOn')[0].getAttribute('w:val')

        # 在default pStyle中查找
        if None in propertydict.values():
            styles = cls.styles
            for style in styles.getElementsByTagName('w:style'):
                if style.getAttribute('w:type') == 'paragraph' and style.getAttribute('w:default') == '1':
                    ppr = style.getElementsByTagName('w:pPr')[0]
                    addProperties(propertydict, ppr)

        # 在 default doc中查找
        if None in propertydict.values():
            styles = cls.styles
            ppr_default = styles.getElementsByTagName('w:pPrDefault')[0]
            ppr_list = ppr_default.getElementsByTagName('w:pPr')
            if ppr_list:
                ppr = ppr_list[0]
                addProperties(propertydict, ppr)

        return propertydict

    # 检测rPr属性是否符合给定的参考属性，不需要对比的属性在参考属性里设置为None
    @classmethod
    def correctRunProperty(cls, run, reference_run_property) -> bool:
        propertydict = cls.getRunProperties(run)
        text = run.getElementsByTagName('w:t')[0].childNodes[0].data
        if re.search(re.compile("[\u4E00-\u9FA5]"), text):
            if propertydict['eastAsia'] != reference_run_property['eastAsia'] and reference_run_property[
                'eastAsia'] is not None:
                return False
        if re.search(re.compile("[0-9a-zA-z]"), text):
            if propertydict['ascii'] != reference_run_property['ascii'] and reference_run_property['ascii'] is not None:
                return False
        for key in propertydict.keys():
            if key != 'eastAsia' and key != 'ascii':
                if propertydict[key] != reference_run_property[key] and reference_run_property[key] is not None:
                    return False
        return True

    # 检测段落格式是否符合给定的参考属性，不需要对比的属性在参考属性里设置为None, 返回错误信息字符串
    @classmethod
    def correctParagraphProperty(cls, p, reference_paragraph_property) -> str:
        propertydict = cls.getParagraphProperties(p)
        text = ''
        if propertydict['firstLineChars'] is not None:
            if propertydict['firstLineChars'] != reference_paragraph_property[
                'firstLineChars'] and reference_paragraph_property is not None:
                text += '缩进错误 '
            propertydict['firstLine'] = reference_paragraph_property['firstLine']  # 当firstLineChars存在时，忽略firstLine这个属性

        for key in propertydict.keys():
            if key != 'firstLineChars':
                if propertydict[key] != reference_paragraph_property[key] and reference_paragraph_property[
                    key] is not None:
                    if key == 'jc':
                        text += '对齐方式错误 '
                    if key == 'line' or key == 'lineRule':
                        text += '行间距错误 '
                    if key == 'firstLine':
                        text += '缩进错误 '
        return text

    # 将这个run的文字背景设置为红色，当一个run中的格式出错时使用该函数
    @classmethod
    def setRed(cls, run):
        rpr_list = run.getElementsByTagName('w:rPr')
        highlight = cls.doc.createElement('w:highlight')
        highlight.setAttribute('w:val', 'red')
        if rpr_list:
            rpr = rpr_list[0]
            rpr.appendChild(highlight)
        else:
            rpr = cls.doc.createElement('w:rPr')
            rpr.appendChild(highlight)
            t = run.getElementsByTagName('w:t')[0]
            run.insertBefore(rpr, t)

    # 在有格式错误的段落最后添加格式错误的标记
    @classmethod
    def addMark(cls, p, error_message):
        run = cls.doc.createElement('w:r')
        rpr = cls.doc.createElement('w:rPr')
        vertAlign = cls.doc.createElement('w:vertAlign')
        vertAlign.setAttribute('w:val', 'subscript')
        color = cls.doc.createElement('w:color')
        color.setAttribute('w:val', 'red')
        t = cls.doc.createElement('w:t')
        text_node = cls.doc.createTextNode(error_message)
        t.appendChild(text_node)
        rpr.appendChild(vertAlign)
        rpr.appendChild(color)
        run.appendChild(rpr)
        run.appendChild(t)
        p.appendChild(run)
