# -*- coding: utf-8 -*-
import os
import re
import time
from py2neo import Graph
from py2neo.ogm import GraphObject, Property, RelatedTo, RelatedFrom
from course.dao import *
# from course.models import Course,TeachingObjective


# 将doc转存为docx
def docSaveToDocx(docpath):
    from win32com import client as wc
    import pythoncom
    pythoncom.CoInitialize()
    try:
        if os.path.exists(docpath):
            docxpath = docpath + 'x'    #转换后docx文件路径
            # 如果已存在同名文件，修改现在文件文件名
            while (os.path.exists(docxpath)):
                temp = str(docxpath).split('.')
                docxpath = temp[len(temp) - 2] + '(1).' + temp[len(temp) - 1]
            word = wc.Dispatch('Word.Application')
            doc = word.Documents.Open(docpath)  # 打开目标路径下的doc文件
            # 如果存在同名文件，修改现在上传文件的文件名
            while (os.path.exists(docxpath)):
                temp = str(docxpath).split('.')
                fileName = temp[len(temp) - 2] + '(1).' + temp[len(temp) - 1]
            doc.SaveAs(docxpath, 12, False, "", True, "", False, False, False, False)  # 另存为docx 路径为docxpath
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
    return docxpath

# 获取文件类型
def getFileType(filePath):
    try:
        if os.path.exists(filePath):
            try:
                fileTypeList = str(filePath).split('.')
                fileType = fileTypeList[len(fileTypeList)-1]
                return str(fileType)
            except Exception as e:
                print(e)
        else:
            print('该路径文件不存在')
    except Exception as e:
        print(e)

# 获取文件夹的文件信息
# 参数：文件夹路径
def getFileInformation(dirPath):
    # uploadPath = os.path.abspath('..') + '\\upload\course\\'
    fileInformationList = []
    fileNameList = os.listdir(dirPath)
    for filename in fileNameList :
        fileInformation = {'name': '', 'size': '', 'date': ''}
        filePath = dirPath + '\\' + filename
        fileInformation['name'] = filename
        fileInformation['size'] = os.path.getsize(filePath)
        fileInformation['date'] = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getmtime(filePath)))
        fileInformationList.append(fileInformation)
    fileInformationList.sort(key=lambda file: file['date'],reverse=True)
    return fileInformationList

# 对课程名称处理，确保课程名称唯一
def coursenameMange(name):
    name = name.upper()
    name = name.replace(' ','')
    name = name.replace(' ','')
    name = name.replace('（','(')
    name = name.replace('）',')')
    return name

#获取课程类别 选修还是必修
#参数：从表格中读取到的字符串
#返回：课程类别字符串
# 获取课程类别
def getCourseCategory(courseCategory):
    courseCategory = courseCategory+' '
    pattren = re.compile(r'□(.*?) ')
    try:
        courseCategory = re.findall(pattren,courseCategory)[0]
        if courseCategory == '选修':
            courseCategory = '必修'
        elif courseCategory == '必修':
            courseCategory == '选修'
        else:
            courseCategory = None
    except Exception as e :
        print(e)
        courseCategory = ''
    return (courseCategory)


# 创建课程对象
# 参数：文档路径，课程表格索引=0
def createCourseObject(filePath, tableIndex):
    # 创建文档对象
    doc = Document(filePath)
    # 获取课程所在表格
    try:
        courseTable = doc.tables[tableIndex]     #文档中第index个表格
    except Exception as e:
        print(e)
    try:
        # 创建课程对象
        from course.models import Course
        course = Course()
    except Exception as e:
        print(e)
    course.name = coursenameMange(courseTable.cell(2, 1).text)              #课程名称
    course.courseNumber = courseTable.cell(0,1).text                        #课程编号
    course.totalHours = courseTable.cell(1,1).text                          #总学时
    course.courseCategory = getCourseCategory(courseTable.cell(3,1).text)   #课程类别
    course.writer = courseTable.cell(4,1).text                              #执笔人
    course.prerequisiteCourses = coursenameMange(courseTable.cell(5,1).text) #先修课程
    course.credit = courseTable.cell(0,3).text                              #学分
    course.experimentalHours = courseTable.cell(1,3).text                   #实验/上机学时
    course.englishName = courseTable.cell(2,3).text                         #英文名称
    course.appliedSpecialty = courseTable.cell(3,3).text                    #适用专业
    course.auditor = courseTable.cell(4,3).text                             #审核人
    return course

# 创建教学目标对象的对象列表
# 参数:文档路径，教学目标表格索引=1
# 注：第二个表格中并没有 贡献度 和 课程名称，所以对象的贡献度属性没有赋值
def createListOfTeachingObjectiveObject(filePath, tableIndex):
    objectList = []
    # 创建文档对象
    doc = Document(filePath)
    # 获取课程所在表格
    try:
        table = doc.tables[tableIndex]     #文档中第index个表格
    except Exception as e:
        print(e)
    for row in range(1,len(table.rows)):
        try:
            from course.models import TeachingObjective
            teachingObjective = TeachingObjective()
        except Exception as e:
            print(e)
        teachingObjective.id = table.cell(row, 0).text            #序号，name
        teachingObjective.describe = table.cell(row, 1).text        #描述
        teachingObjective.waysToAchieve = table.cell(row, 2).text   #达成途径
        teachingObjective.mainCriteria = table.cell(row, 3).text    #主要判据
        objectList.append(teachingObjective)
    return objectList

def updateNameForTeachingObjectiveObject(objectList, courseName):
    try:
        for object in objectList:
            object.name = courseName + object.id
    except Exception as e:
        print(e)
        print('教学目标name更新失败')

# 更新教学目标对象的 贡献度 属性
# 参数：课程对象列表，表格索引=2
def updateContributionDegreeForTeachingObjectiveObject(objectList,filePath,tableIndex):
    # 创建文档对象
    doc = Document(filePath)
    # 获取课程所在表格
    try:
        table = doc.tables[tableIndex]     #文档中第index个表格
    except Exception as e:
        print(e)
    if len(objectList)>0 :
        if len(table.rows) > 0:
            if len(table.rows) > 1:
                #教学目标对象 贡献值属性 赋值
                for row in range(1, len(table.rows)):
                    for object in objectList:
                        if object.name == table.cell(row, 2).text :                 #表格中 对应的本课程教学目标 列
                            object.contributionDegree = table.cell(row, 3).text     #表格中 贡献值 列
                        else:
                            pass
            else:
                print('表格只有表头')
        else:
            print('表格不存在')
    else:
        print('对象列表长度为0')
# 更新教学目标对象的 课程名称 属性
# 参数：课程对象列表，课程名称
def updateCourseNameForTeachingObjectiveObject(objectListOfTeachingObjective,courseName):
    try:
        for teachingObjective in objectListOfTeachingObjective:
            teachingObjective.courseName = courseName
    except Exception as e:
        print(e)


# 获取 教学目标对应的指标点的序号并完成教学目标对象贡献值属性的赋值
# 参数：教学目标对象列表,文档路径，表格索引=2
def getTheIndexOfTeachingObjectiveMapIndexPoint(teachingObjective,path,index):
    # 创建文档对象
    doc = Document(path)
    # 获取课程所在表格
    try:
        table = doc.tables[index]     #文档中第index个表格
    except Exception as e:
        print(e)
    try:
        for row in range(1, len(table.rows)):
            value = table.cell(row, 2).text     #表示映射关系的表格值
            # 对不同表格值做不同处理
            if '~' in value:
                valueList = []
                tempList = value.split('~')
                for temp in range(int(tempList[0]),int(tempList[1])+1):
                    valueList.append(str(temp))
                if teachingObjective.id in valueList:
                    # 教学目标贡献值属性赋值
                    teachingObjective.contributionDegree = table.cell(row, 3).text
                    # 表格中对应指标点的文本，用于提取序号
                    tableText = table.cell(row, 1).text
                    # 正则表达式 匹配指标点的序号
                    pattren = re.compile(r'^(\d*).(\d*)', re.S)
                    index = re.match(pattren, tableText).group(0)
                    return index
                else:
                    pass
            elif '、' in value:
                valueList = []
                tempList = value.split('、')
                for temp in tempList:
                    valueList.append(str(temp))
                if teachingObjective.id in valueList:
                    # 教学目标贡献值属性赋值
                    teachingObjective.contributionDegree = table.cell(row, 3).text
                    # 表格中对应指标点的文本，用于提取序号
                    tableText = table.cell(row, 1).text
                    # 正则表达式 匹配指标点的序号
                    pattren = re.compile(r'^(\d*).(\d*)', re.S)
                    index = re.match(pattren, tableText).group(0)
                    return index
                else:
                    pass
            elif '，' in value:
                valueList = []
                tempList = value.split('，')
                for temp in tempList:
                    valueList.append(str(temp))
                if teachingObjective.id in valueList:
                    # 教学目标贡献值属性赋值
                    teachingObjective.contributionDegree = table.cell(row, 3).text
                    # 表格中对应指标点的文本，用于提取序号
                    tableText = table.cell(row, 1).text
                    # 正则表达式 匹配指标点的序号
                    pattren = re.compile(r'^(\d*).(\d*)', re.S)
                    index = re.match(pattren, tableText).group(0)
                    return index
                else:
                    pass
            elif ',' in value:
                valueList = []
                tempList = value.split(',')
                for temp in tempList:
                    valueList.append(str(temp))
                if teachingObjective.id in valueList:
                    # 教学目标贡献值属性赋值
                    teachingObjective.contributionDegree = table.cell(row, 3).text
                    # 表格中对应指标点的文本，用于提取序号
                    tableText = table.cell(row, 1).text
                    # 正则表达式 匹配指标点的序号
                    pattren = re.compile(r'^(\d*).(\d*)', re.S)
                    index = re.match(pattren, tableText).group(0)
                    return index
                else:
                    pass
            else:
                if teachingObjective.id == value:
                    # 教学目标贡献值属性赋值
                    teachingObjective.contributionDegree = table.cell(row, 3).text
                    # 表格中对应指标点的文本，用于提取序号
                    tableText = table.cell(row, 1).text
                    # 正则表达式 匹配指标点的序号
                    pattren = re.compile(r'^(\d*).(\d*)', re.S)
                    index = re.match(pattren, tableText).group(0)
                    return index
                else:
                    pass
    except Exception as e:
        print(e)

# 构建 教学目标 和 指标点 之间的关系
# 参数：教学目标对象列表,文档路径，教学目标和指标点关系表格索引=2 ,图形数据库
def createRelationBetweenTeachingObjectiveAndIndexPoint(TO_ObjeceList, filePath, tableIndex,graph):
    try:
        for object in TO_ObjeceList:
            # 获取到 教学目标对象 对应的 指标点 的指标点序号
            indexOfIndexPoint = getTheIndexOfTeachingObjectiveMapIndexPoint(object, filePath, tableIndex)
            # print(indexOfIndexPoint)
            try:
                from course.models import IndexPoint
                # 创建 指标点 对象
                indexPoint = IndexPoint()
                try:
                    # 根据 指标点序号 获取到指标点对象的节点列表
                    indexPointList = indexPoint.match(graph,indexOfIndexPoint)
                    # print(type(indexPointList))
                    # 将列表中第一个节点实例化为 指标点对象
                    indexPoint = indexPointList.first()
                    # print(type(indexPoint))
                    object.contribution.add(indexPoint)#添加关系贡献
                except Exception as e:
                    print('从数据库获取指标点节点出错')
                    print(e)

            except Exception as e:
                print(e)
    except Exception as e:
        print(e)

# 构建 教学目标 和 课程 之间的关系
# 参数：教学目标对象列表,文档路径，教学目标和课程关系表格索引=0,图形数据库
def createRelationBetweenTeachingObjectiveAndCourse(TO_ObjeceList,courseObject):
    try:
        for object in TO_ObjeceList:
            # 构建课程 达到 教学目标关系
            object.reach.add(courseObject)
    except Exception as e:
        print(e)

# 构建 课程对象 与 课程 之间关系
def createRelationBetweenCourseAndCourse(courseObject,couseName,graph):
    # 从数据库中查找特点node
    prerequisiteCourseNode = graph.nodes.match('Course', name=couseName).first()
    if prerequisiteCourseNode:
        # 将node转换为对象
        prerequisiteCourse = courseObject.wrap(prerequisiteCourseNode)
    else:
        try:
            from course.models import Course
            prerequisiteCourse = Course()
        except Exception as e:
            print(e)
        prerequisiteCourse.name = couseName
    # 构建关系
    courseObject.Prerequisite.add(prerequisiteCourse)

# 构建 课程 和 先修课程 之间关系
# 参数：课程对象,图形数据库
def createRelationBetweenCourseAndPrerequisiteCourse(courseObject,graph):
    try:
        prerequisiteCourseName = courseObject.prerequisiteCourses
        if prerequisiteCourseName == '无' :  #如果没有先修课程
            pass
        elif '、' in prerequisiteCourseName:
            prerequisiteCourseList = prerequisiteCourseName.split('、')
            for name in prerequisiteCourseList:
                createRelationBetweenCourseAndCourse(courseObject,name,graph)
        elif '，' in prerequisiteCourseName:
            prerequisiteCourseList = prerequisiteCourseName.split('，')
            for name in prerequisiteCourseList:
                createRelationBetweenCourseAndCourse(courseObject,name,graph)
        else:
            createRelationBetweenCourseAndCourse(courseObject, prerequisiteCourseName, graph)
    except Exception as e:
        print(e)

def saveRelationToDB(filePath, graph):
    # 创建课程对象
    course = createCourseObject(filePath, 0)
    # 创建教学目标对象列表
    TO_ObjectList = createListOfTeachingObjectiveObject(filePath, 1)
    # 更新教学目标对象 贡献值
    updateContributionDegreeForTeachingObjectiveObject(TO_ObjectList, filePath, 2)
    # 更新教学目标对象 课程名称
    updateCourseNameForTeachingObjectiveObject(TO_ObjectList, course.name)
    # 构建关系
    createRelationBetweenTeachingObjectiveAndIndexPoint(TO_ObjectList, filePath, 2, graph)
    createRelationBetweenTeachingObjectiveAndCourse(TO_ObjectList, course)
    createRelationBetweenCourseAndPrerequisiteCourse(course, graph)
    # 更新教学目标对象 name  将name作为主键
    updateNameForTeachingObjectiveObject(TO_ObjectList, course.name)
    #     持久化到数据库
    saveListOfObjectToDB(graph, TO_ObjectList)
    graph.push(course)


if __name__ == '__main__':
    path = 'C:\\Users\\10615\Desktop\\1.docx'
    # path = 'C:\\Users\\10615\Desktop\\《专业实习》课程教学大纲（实践）-2016版.docx'
    # path = 'C:\\Users\\10615\Desktop\\《大学体育》课程教学大纲（理论）-2016版.docx'
    path = 'C:\\Users\\10615\Desktop\\list\\2016版-审核完成\专业类课程-理论\\《编译原理》课程教学大纲（理论）-2016版.docx'
    # path = 'E:\pycharmProject\CourseSystem\\upload\course\\《专业实习》课程教学大纲（实践）-2016版.docx'
    # graph = Graph("http://localhost:7474", username="neo4j", password='431879')
    # saveRelationToDB(path,graph)
    # 创建课程对象
    course = createCourseObject(path, 0)
    print('课程名称：',course.name)
    print('课程编号：',course.courseNumber)
    print("课程类别：",course.courseCategory)
    print("先修课程：",course.prerequisiteCourses)
    print("执笔人：",course.writer)
