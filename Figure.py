import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom
from Utildom import Util


class Figure:
    figures = []

    @classmethod
    def setup(cls):
        cls.getFigures()

    @classmethod
    def isPic(cls, p):
        for r in p.getElementsByTagName('w:r'):
            if r.getElementsByTagName('w:drawing') or r.getElementsByTagName('pic:pic') or r.getElementsByTagName(
                    'w:object'):
                return True
            # elif r.getElementsByTagName('mc:AlternateContent'):
            #     ac = r.getElementsByTagName('mc:AlternateContent')[0]
            #     choice_list = ac.getElementsByName('mc:Choice')
            #     fallback_list = ac.getElementsByName('mc:Fallback')
            #     if choice_list and fallback_list:
            #         if choice_list[0].getElementsByTagName('w:drawing') or fallback_list[0].getElementsByTagName('pic:pic'):
            #             location = cls.getFigureTitleLocation(p)[3]
            #             if location is not None:
            #                 return True
        return False

    @classmethod
    def getFigures(cls):
        for node in Util.getChildNodesOfNormalText():
            if cls.isPic(node):
                cls.figures.append(node)

    # 查找图标题位置
    @classmethod
    def getFigureTitleLocation(cls, p):
        location = 0
        # 从图往前找
        while location > -3:
            location -= 1
            node = p.previousSibling
            if node.tagName == 'w:p':
                text = Util.getFullText(node)
                if re.search('^图[\s]?[1-9][0-9]*\.[1-9][0-9]*', text):
                    return node, text, location

        # 从图往后找
        location = 0
        while location < 3:
            location += 1
            node = p.nextSibling
            if node.tagName == 'w:p':
                text = Util.getFullText(node)
                if re.search('^图[\s]?[1-9][0-9]*\.[1-9][0-9]*', text):
                    return node, text, location

        return None, None, None

    @classmethod
    def containText(cls, p):
        if Util.getFullText(p) != '':
            return True
        return False

    @classmethod
    def exceedMargin(cls, r):
        if r.getElementsByTagName('w:drawing'):
            inline_list = r.getElementsByTagName('w:drawing')[0].getElementsByTagName('wp:inline')
            if inline_list:
                extent = inline_list[0].getElementsByTagName('wp:extent')[0]
                cx = int(extent.getAttribute('cx'))
                cx = cx / 12700 * 20
                if cx > Util.page_width - Util.left_margin - Util.right_margin:
                    return True

        elif r.getElementsByTagName('w:object'):
            shape = r.getElementsByTagName('w:object')[0].getElementsByTagName('v:shape')[0]
            attr_style = shape.getAttribute('style')
            width = float(re.findall(r'width:([0-9\.]+)pt', attr_style)[0])
            if width * 20 > Util.page_width - Util.left_margin - Util.right_margin:
                return True
        return False

    @classmethod
    def addTitleErrorMessage(cls, fig, p, error_message):
        if '图缺少标题' in error_message:
            p = Util.doc.createElement('w:p')
            ppr = Util.doc.createElement('w:pPr')
            pstyle = Util.doc.createElement('w:pStyle')
            pstyle.setAttribute('w:val', 'error')
            ppr.appendChild(pstyle)
            r = Util.doc.createElement('w:r')
            t = Util.doc.createElement('w:t')
            text_node = Util.doc.createTextNode(error_message)
            t.appendChild(text_node)
            r.appendChild(t)
            p.appendChild(ppr)
            p.appendChild(r)
            Util.doc.childNodes[0].childNodes[0].insertBefore(p, fig.nextSibling)

        else:
            Util.addMarkP(p, error_message)