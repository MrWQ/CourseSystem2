import time

from django.test import TestCase
from docx import *
import re
# Create your tests here.

# doc文档

# 获取文档对象
# doc = Document(path)
# 获取文档中的段落，列表类型
# paragraphs = doc.paragraphs
# try:
#     # 由于文档中含义标题“三、毕业要求”，为去除该标题所以选取文档段落对象从第二行开始
#     paragraphs = paragraphs[1:len(paragraphs)]
# except Exception as e:
#     print(e)
# for paragraph in paragraphs:
#     print(paragraph.text)

# 将doc转存为docx
from py2neo import Graph, Node

from course.util import getFileType, getFileInformation

docxpath = 'C:\\Users\\10615\Desktop\\4.docx'
docpath = 'C:\\Users\\10615\Desktop\\1.doc'

from win32com import client as wc
import os
def docToDocx(docpath):
    print(getFileType(docpath))
    if str(getFileType(docpath)) == 'doc':
        print('docccccccccccccc')
    try:
        if os.path.lexists(docpath):
            docxpath = docpath + 'x'
            word = wc.Dispatch('Word.Application')
            doc = word.Documents.Open(docpath)  # 打开目标路径下的doc文件
            doc.SaveAs(docxpath, 12, False, "", True, "", False, False, False, False)  # 转化后路径下的docx文件
            doc.Close()
            word.Quit()
            print('doc 成功转换为 docx')
            try:
                os.remove(docpath)
                print('成功删除原doc文件')
            except Exception as e:
                print(e)
        else:
            print('该路径文件不存在')
    except Exception as e:
        print(e)




if __name__ == '__main__':
    # docToDocx(docpath)
    graph = Graph("http://localhost:7474", username="neo4j", password='431879')
    a= '1.2.2.docx'
    b = a.split('.')
    a= a.replace('.','-',len(b)-2)
    print(a)
    # a = Course()
    # a.name = '面向对象技术（C++/Java）'
    # a.courseNumber = '132435'
    # graph.push(a)

    # indexpoint = IndexPoint()
    # indexpoint.name='testttttt'
    # pt = indexpoint.match(graph,'2.3')
    # print(type(pt))
    # print(len(pt))
    # for i in pt:
    #     print(type(i))
    #     print(i.name)
    #     print(i.describe)
    # print(getFileInformation('E:\pycharmProject\CourseSystem\\upload\course'))

