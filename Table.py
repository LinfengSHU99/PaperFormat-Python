import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom

# TODO: 有一些表格实际上不是表格而是公式，如何识别出这些作者本意不是表格的表格，忽略它们，是下一个解决的问题

class Table:
    @classmethod
    def isThreeLineTable(cls, tbl, styles) -> bool:

        # 从style中查找表格的总边框信息
        def getTableBorders(tbl, styles):
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

        # 收集表格的总边框信息
        tblBorders = getTableBorders(tbl, styles)
        tbl_top = 'none'
        tbl_bottom = 'none'
        tbl_left = 'none'
        tbl_right = 'none'
        tbl_insideH = 'none'
        tbl_insideV = 'none'
        if tblBorders is not None:
            tbl_top = tblBorders.getElementsByTagName('w:top')[0].getAttribute('w:val') if tblBorders.getElementsByTagName(
                'w:top') else 'none'
            tbl_bottom = tblBorders.getElementsByTagName('w:bottom')[0].getAttribute('w:val') if tblBorders.getElementsByTagName(
                'w:bottom') else 'none'
            tbl_left = tblBorders.getElementsByTagName('w:left')[0].getAttribute('w:val') if tblBorders.getElementsByTagName(
                'w:left') else 'none'
            tbl_right = tblBorders.getElementsByTagName('w:right')[0].getAttribute('w:val') if tblBorders.getElementsByTagName(
                'w:right') else 'none'
            tbl_insideV = tblBorders.getElementsByTagName('w:insideV')[0].getAttribute('w:val') if tblBorders.getElementsByTagName(
                'w:insideV') else 'none'
            tbl_insideH = tblBorders.getElementsByTagName('w:insideH')[0].getAttribute('w:val') if tblBorders.getElementsByTagName(
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