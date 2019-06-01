# -*- coding: utf-8 -*-
from py2neo import Graph
from course.models import IndexPoint
from course.util import *
from course.dao import saveListOfObjectToDB

# 通过文档段落列表获取文档第一级指标点列表
def getListOnFirstLevelOfIndexpoint(paragraphs):
    listOnFirstLevelOfIndexPoint = []
    for paragraph in paragraphs:
        if "：" in paragraph.text:
            listOnFirstLevelOfIndexPoint.append(paragraph.text)
    return listOnFirstLevelOfIndexPoint

# 通过文档段落列表获取文档第二级指标点列表
def getListOnSecondLevelOfIndexpoint(paragraphs):
    listOnSeconeLevelOfIndexPoint = []
    for paragraph in paragraphs:
        if ":" in paragraph.text or "：" in paragraph.text:
            pass
        elif paragraph.text =='':
            pass
        else:
            listOnSeconeLevelOfIndexPoint.append(paragraph.text)
    return listOnSeconeLevelOfIndexPoint

# 创建第一级指标点对象
# 参数：第一级指标点文本内容
def createFirstLevelOfIndexPointObject(paragraphText):
    indexPoint = IndexPoint()
    try:
        paragraphTextList = paragraphText.split("：")
        indexPoint.name = paragraphTextList[0]
        indexPoint.describe = paragraphTextList[1]
    except Exception as e:
        print(e)
    return indexPoint

# 创建第二级指标点对象
# 参数：第二级指标点文本内容
def createSecondLevelOfIndexPointObject(paragraphText):
    indexPoint = IndexPoint()
    try:
        paragraphTextList = paragraphText.split(" ")
        indexPoint.name = paragraphTextList[0]
        indexPoint.describe = paragraphTextList[1]
    except Exception as e:
        print(e)
    return indexPoint


# 创建一级指标点和二级指标点之间的关系，并持久化到数据库
def createAndSaveRelationToDB(filePath, graph):
    # 获取文档对象
    doc = Document(filePath)
    # 获取文档中的段落，列表类型
    paragraphs = doc.paragraphs
    try:
        # 由于文档中含义标题“三、毕业要求”，为去除该标题所以选取文档段落对象从第二行开始
        paragraphs = paragraphs[1:len(paragraphs)]
    except Exception as e:
        print(e)

    # print(getListOnFirstLevelOfIndexpoint(paragraphs))        #第一级指标点内容列表
    # print(getListOnSecondLevelOfIndexpoint(paragraphs))       #第二级指标点内容列表
    # 第一级指标点对象
    listOnFirstLevelOfIndexpoint = getListOnFirstLevelOfIndexpoint(paragraphs)
    firstLevelOfObjectlist = []
    for indexpointText in listOnFirstLevelOfIndexpoint:
        # 把第一级指标点内容传人 来 创建对象，在把对象加入到列表中
        firstLevelOfObjectlist.append(createFirstLevelOfIndexPointObject(indexpointText))
    print(firstLevelOfObjectlist[0].name + '==' + firstLevelOfObjectlist[1].describe)
    # 第二级指标点对象
    listOnSecondLevelOfIndexpoint = getListOnSecondLevelOfIndexpoint(paragraphs)
    secondLevelofObjectlist = []
    for indexpointText in listOnSecondLevelOfIndexpoint:
        secondLevelofObjectlist.append(createSecondLevelOfIndexPointObject(indexpointText))
    print(secondLevelofObjectlist[0].name + '==' + secondLevelofObjectlist[1].describe)

    # 创建关系
    if firstLevelOfObjectlist != [] and secondLevelofObjectlist != []:
        for fistLevelObject in firstLevelOfObjectlist:
            for secondLevelofObject in secondLevelofObjectlist:
                try:
                    # 通过分割字符串确定所属关系 例：1.sfe 子项有 1.1skfsof
                    fistLevelObjectNumber = fistLevelObject.name.split('．')[0]
                    secondLevelofObjectNumber = secondLevelofObject.name.split('.')[0]
                    if (fistLevelObjectNumber == secondLevelofObjectNumber):
                        fistLevelObject.subitem.add(secondLevelofObject)
                except Exception as e:
                    print(e)

    # 一键持久化到数据库
    saveListOfObjectToDB(graph, firstLevelOfObjectlist)


if __name__ == '__main__':
    # neo4j图形数据库
    graph = Graph("http://localhost:7474", username="neo4j", password='431879')
    # doc文档
    path = 'C:\\Users\\10615\Desktop\\2.docx'
    createAndSaveRelationToDB(path,graph)



