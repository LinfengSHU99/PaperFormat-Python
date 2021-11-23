import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom

import AddStyle
from Utildom import Util
from Table import Table

title0_run_property = {'eastAsia': "黑体", 'ascii': "Times New Roman", 'sz': "32", 'szCs': None, 'kern': None}
title1_run_property = {'eastAsia': "黑体", 'ascii': "Times New Roman", 'sz': "28", 'szCs': None, 'kern': None}
title2_run_property = {'eastAsia': "黑体", 'ascii': "Times New Roman", 'sz': "24", 'szCs': None, 'kern': None}
title_run_property_list = [title0_run_property, title1_run_property, title2_run_property]
normal_run_property = {'eastAsia': "宋体", 'ascii': "Times New Roman", 'sz': "24", 'szCs': None, 'kern': None}

normal_paragraph_property = {'jc': 'both', 'line': '360', 'lineRule': 'auto', 'before': None, 'beforeLines': None,
                             'after': None,
                             'afterLines': None, 'firstLine': '480', 'firstLineChars': '200'}

table_title_run_property = {'eastAsia': "宋体", 'ascii': "Times New Roman", 'sz': "21", 'szCs': None, 'kern': None}
table_title_paragraph_property = {'jc': 'center', 'line': '360', 'lineRule': 'auto', 'before': None,
                                  'beforeLines': None,
                                  'after': None,
                                  'afterLines': None, 'firstLine': None, 'firstLineChars': None}


def matchNormal(doc):
    reference_run_property = None
    reference_paragraph_property = None
    # a = Util.getChildNodesOfNormalText()
    for child_node in Util.getChildNodesOfNormalText():
        if child_node.tagName == 'w:p':
            text = Util.getFullText(child_node)
            if text.strip(' ') == '':  # 仅当段落有文字时检测
                continue
            if Util.isTabTitle(child_node):
                continue
            if Util.isPicTitle(child_node):
                continue
            if Util.isTitle(child_node):  # 如果该段落为标题，进行标题格式的匹配
                reference_run_property = title_run_property_list[Util.titleType(child_node)]
                for run in child_node.getElementsByTagName('w:r'):
                    if run.getElementsByTagName('w:t'):  # 仅当run中有文字时进行检测
                        if Util.correctRunProperty(run, reference_run_property) is False:
                            Util.setRed(run)

            else:  # 如果该段落不是标题，则视为正文（根据以后需求增添if-else语句视为其他内容）
                reference_run_property = normal_run_property
                reference_paragraph_property = normal_paragraph_property
                for run in child_node.getElementsByTagName('w:r'):  # 匹配run的格式
                    if run.getElementsByTagName('w:t'):
                        if Util.correctRunProperty(run, reference_run_property) is False:
                            Util.setRed(run)
                # 匹配段落格式
                error_message = Util.correctParagraphProperty(child_node, reference_paragraph_property)
                if error_message != '':
                    Util.addMark(child_node, error_message)


def matchTable():
    for tbl in Table.tables_with_borders:
        p, tbl_title, location = Table.getTableTitleLocation(tbl)
        error_message = ''
        if not Table.isThreeLineTable(tbl, Util.styles):
            error_message += '不是三线表 '
        if tbl_title is None:
            error_message += '表格缺少标题 '

        else:
            if location > 0:
                error_message += '表格标题应在表格上方 '
            error_message += Util.correctParagraphProperty(p, table_title_paragraph_property)
            for r in p.getElementsByTagName('w:r'):
                if not Util.correctRunProperty(r, table_title_run_property):
                    Util.setRed(r)
        Table.addTitleErrorMessage(tbl, p, error_message)
