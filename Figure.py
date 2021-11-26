import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom
from Utildom import Util


class Figure:
    #
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
    def exceedMargin(cls, p):
        # TODO: to be finished
        pass
