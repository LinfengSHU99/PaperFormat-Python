import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom
from Utildom import Util



class Table:
    doc = Util.doc
    styles = Util.styles
    tables_with_borders = []

    @classmethod
    def setup(cls):
        cls.getTableWithBorders(cls.styles)

    # 从tbl和style中查找表格的总边框信息
    @classmethod
    def getTableBorders(cls, tbl, styles):
        tblBorders = None
        tblPr = tbl.getElementsByTagName('w:tblPr')[0] if len(tbl.getElementsByTagName('w:tblPr')) > 0 else None
        if tblPr is not None:
            tblBorders = tblPr.getElementsByTagName('w:tblBorders')[0] if len(
                tblPr.getElementsByTagName('w:tblBorders')) > 0 else None
            if tblBorders is None:
                tbl_style_id = tblPr.getElementsByTagName('w:tblStyle')[0].getAttribute('w:val') if len(
                    tblPr.getElementsByTagName('w:tblStyle')) > 0 else None
                # 循环查找
                while tbl_style_id is not None:
                    for style in styles.getElementsByTagName('w:style'):
                        if style.getAttribute('w:styleId') == tbl_style_id:
                            tbl_style_id = None
                            tblBorders = style.getElementsByTagName('w:tblBorders')[0] if len(
                                style.getElementsByTagName('w:tblBorders')) > 0 else None
                            if tblBorders is None:
                                if style.getElementsByTagName('w:basedOn'):
                                    tbl_style_id = style.getElementsByTagName('w:basedOn')[0].getAttribute('w:val')
                            break
            if tblBorders is None:
                # 如果没有找到，在default style中找
                for style in styles.getElementsByTagName('w:style'):
                    if style.getAttribute('w:type') == 'table' and style.getAttribute('w:default') == '1':
                        tblBorders = style.getElementsByTagName('w:tblBorders')[0] if style.getElementsByTagName(
                            'w:tblBorders') else None

        return tblBorders


    @classmethod
    def getTableTitleLocation(cls, tbl):
        location = 0
        # 从表格往前找标题
        while location > -3:
            location -= 1
            node = tbl.previousSibling
            if node.tagName == 'w:p':
                text = Util.getFullText(node)
                if re.search('^表[\s]?[1-9][0-9]*\.[1-9][0-9]*', text):
                    return node, text, location
        location = 0
        # 从表格往后找标题
        while location > 3:
            location += 1
            node = tbl.nextSibling
            if node.tagName == 'w:p':
                text = Util.getFullText(node)
                if re.search('^表[\s]?[1-9][0-9]*\.[1-9][0-9]*', text):
                    return node, text, location
        return None, None, None

    # 获取正文中有边框的表格（可以过滤一些使用表格格式，但作者本意不是表格的表格）
    @classmethod
    def getTableWithBorders(cls, styles):
        for node in Util.getChildNodesOfNormalText():
            if node.tagName == 'w:tbl':
                contain = 0
                if cls.getTableBorders(node, styles):
                    tblBorders = cls.getTableBorders(node, styles)
                    for border in tblBorders.childNodes :
                        if border.getAttribute('w:val') != 'none':
                            cls.tables_with_borders.append(node)
                            contain = 1
                            break
                if contain == 0:
                    for tc_borders in node.getElementsByTagName('w:tcBorders'):
                        for border in tc_borders.childNodes:
                            if border.getAttribute('w:val') != 'nil':
                                cls.tables_with_borders.append(node)
                                break
                        else:
                            continue
                        break

    @classmethod
    def isThreeLineTable(cls, tbl, styles) -> bool:

        # 收集表格的总边框信息
        tblBorders = cls.getTableBorders(tbl, styles)
        tbl_top = 'none'
        tbl_bottom = 'none'
        tbl_left = 'none'
        tbl_right = 'none'
        tbl_insideH = 'none'
        tbl_insideV = 'none'
        if tblBorders is not None:
            tbl_top = tblBorders.getElementsByTagName('w:top')[0].getAttribute(
                'w:val') if tblBorders.getElementsByTagName(
                'w:top') else 'none'
            tbl_bottom = tblBorders.getElementsByTagName('w:bottom')[0].getAttribute(
                'w:val') if tblBorders.getElementsByTagName(
                'w:bottom') else 'none'
            tbl_left = tblBorders.getElementsByTagName('w:left')[0].getAttribute(
                'w:val') if tblBorders.getElementsByTagName(
                'w:left') else 'none'
            tbl_right = tblBorders.getElementsByTagName('w:right')[0].getAttribute(
                'w:val') if tblBorders.getElementsByTagName(
                'w:right') else 'none'
            tbl_insideV = tblBorders.getElementsByTagName('w:insideV')[0].getAttribute(
                'w:val') if tblBorders.getElementsByTagName(
                'w:insideV') else 'none'
            tbl_insideH = tblBorders.getElementsByTagName('w:insideH')[0].getAttribute(
                'w:val') if tblBorders.getElementsByTagName(
                'w:insideH') else 'none'

        first_bottom = True
        tr_list = tbl.getElementsByTagName('w:tr')
        row_cnt = 0
        # tcBorders = None
        if len(tr_list) <= 2:
            return True
        if tbl_insideV != 'none':
            return False
        for tr in tr_list:

            row_cnt += 1
            tc_list = tr.getElementsByTagName('w:tc')
            for tc in tc_list:
                tcBorders = None
                tcPr = tc.getElementsByTagName('w:tcPr')[0] if len(tc.getElementsByTagName('w:tcPr')) > 0 else None
                if tcPr is not None:
                    tcBorders = tcPr.getElementsByTagName('w:tcBorders')[0] if len(
                        tcPr.getElementsByTagName('w:tcBorders')) > 0 else None
                    if tcBorders is not None:
                        # 判断每个单元格是否有左右边框，以及整个表是否有左右边框
                        if tcBorders.getElementsByTagName('w:left'):
                            left_border = tcBorders.getElementsByTagName('w:left')[0]
                            if left_border.getAttribute('w:val') != 'nil':
                                return False
                        elif tbl_left != 'none':
                            return False
                        if tcBorders.getElementsByTagName('w:right'):
                            right_border = tcBorders.getElementsByTagName('w:right')[0]
                            if right_border.getAttribute('w:val') != 'nil':
                                return False
                        elif tbl_right != 'none':
                            return False

                        if row_cnt == 1:  # 表格第一行
                            # 判断表格第一行的顶部是否有线
                            if tcBorders.getElementsByTagName('w:top'):  #
                                top_border = tcBorders.getElementsByTagName('w:top')[0]
                                if top_border.getAttribute('w:val') == 'nil':
                                    return False
                            elif tbl_top == 'none':
                                return False
                            if tcBorders.getElementsByTagName('w:bottom'):
                                bottom_boder = tcBorders.getElementsByTagName('w:bottom')[0]
                                if bottom_boder.getAttribute('w:val') == 'nil':
                                    first_bottom = False
                            elif tbl_insideH == 'none':
                                return False

                        if row_cnt == 2:  # 表格第二行
                            if tcBorders.getElementsByTagName('w:top'):
                                top_border = tcBorders.getElementsByTagName('w:top')[0]
                                if top_border.getAttribute('w:val') == 'nil' and first_bottom is False:
                                    return False

                        if 1 < row_cnt < len(tr_list):  # 除去第一行和最后一行
                            if tcBorders.getElementsByTagName('w:bottom'):
                                if tcBorders.getElementsByTagName('w:bottom')[0].getAttribute('w:val') != 'nil':
                                    return False
                            elif tbl_insideH != 'none':
                                return False
                        if 2 < row_cnt <= len(tr_list):  # 除去前两行
                            if tcBorders.getElementsByTagName('w:top'):
                                if tcBorders.getElementsByTagName('w:top')[0].getAttribute('w:val') != 'nil':
                                    return False
                            elif tbl_insideH != 'none':
                                return False
                        if row_cnt == len(tr_list):
                            if tcBorders.getElementsByTagName('w:top'):
                                if tcBorders.getElementsByTagName('w:top')[0].getAttribute('w:val') != 'nil':
                                    return False
                            elif tbl_insideH != 'none':
                                return False
                            if tcBorders.getElementsByTagName('w:bottom'):
                                if tcBorders.getElementsByTagName('w:bottom')[0].getAttribute('w:val') == 'nil':
                                    return False
                            elif tbl_bottom == 'none':
                                return False
                if tcBorders is None:  # 没有tcBorders的情况
                    if row_cnt == 1:
                        if tbl_top == 'none':
                            return False
                        if tbl_insideH == 'none':
                            return False
                    if len(tr_list) > row_cnt > 1:
                        if row_cnt != 2 and tbl_insideH != 'none':
                            return False
                    if row_cnt == len(tr_list):
                        if tbl_insideH != 'none':
                            return False
                        if tbl_bottom == 'none':
                            return False
        return True

    @classmethod
    def addTitleErrorMessage(cls, tbl, p, error_message):
        if '表格缺少标题' in error_message:
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
            Util.doc.childNodes[0].childNodes[0].insertBefore(p, tbl)

        else:
            Util.addMarkP(p, error_message)
