#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import xml.sax
from datetime import datetime

import openpyxl

from utils.utils.excel import write_sheet

node_map = {}
root_id = None


class EmmxHandler(xml.sax.ContentHandler):
    def __init__(self):
        self.node = None
        self.shape_level = 0
        self.validate = False
        self.in_text = False
        self.in_tp = False
        self.in_level_data = False
        self.in_note = False

    # 元素开始事件处理
    def startElement(self, tag, attributes):
        if tag == 'Shape':
            if self.shape_level > 0:  # Shape嵌套，只读取外层
                return
            else:
                self.node = {'ID': attributes.get('ID'), 'Type': attributes.get('Type')}
                if attributes.get('Type') == 'MainIdea':
                    global root_id
                    root_id = self.node['ID']
            self.shape_level += 1
        elif tag == 'Text':
            self.in_text = True
        elif tag == 'tp':
            if self.in_text:
                self.in_tp = True
        elif tag == 'LevelData':
            self.in_level_data = True
        elif tag == 'SubLevel':
            if self.shape_level > 0 and self.in_level_data:
                v = attributes.get('V')
                self.node['children'] = str(v).split(';')
        elif tag == 'Note':
            self.in_note = True

    # 元素结束事件处理
    def endElement(self, tag):
        if tag == 'Shape':
            if self.shape_level > 1:
                return
                self.shape_level -= 1
            else:
                if self.validate:
                    global node_map
                    node_map[self.node['ID']] = self.node
                    print(self.node)
                self.__init__()
        elif tag == 'tp':
            self.in_tp = False
        elif tag == 'Text':
            self.in_text = False
        elif tag == 'LevelData':
            self.in_level_data = False
        elif tag == 'Note':
            self.in_note = False

    # 内容事件处理
    def characters(self, content):
        if self.shape_level == 1 and (not self.in_note) and self.in_text and self.in_tp and content != '':
            self.node['content'] = content
            self.validate = True


def recursive_list_node(node, parents, rows):
    parents.append(node)
    rows.append(node_list_to_row(parents))
    if len(parents) < 4:  # 只取前四级
        if 'children' in node and len(node['children']) > 0:
            for node_id in node['children']:
                sub_node = node_map[node_id]
                recursive_list_node(sub_node, parents, rows)
    parents.pop(len(parents) - 1)


def node_list_to_row(node_list):
    return list(map(lambda i: i['content'], node_list))


def build_rows():
    global node_map, root_id
    root = node_map[root_id]
    rows = []
    recursive_list_node(root, [], rows)
    return rows


def write_excel(path, headers, rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    # sheet.title = sheet_names[0]
    write_sheet(sheet, headers, rows)

    workbook.save(path)


if (__name__ == "__main__"):
    # 创建一个 XMLReader
    parser = xml.sax.make_parser()
    # turn off namepsaces
    parser.setFeature(xml.sax.handler.feature_namespaces, 0)

    # 重写 ContextHandler
    Handler = EmmxHandler()
    parser.setContentHandler(Handler)

    parser.parse('page.xml')
    rows = build_rows()
    print(rows)
    header = ['系统', '一级模块', '二级模块', '三级模块']
    excel_path = os.path.abspath(
        os.path.join(os.getcwd(), 'GMP_' + datetime.now().strftime('%Y%m%d%H%M%S') + '.xlsx'))

    write_excel(excel_path, header, rows)
