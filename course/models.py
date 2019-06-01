from docx import *
from py2neo import Graph
from py2neo.ogm import GraphObject, Property, RelatedTo, RelatedFrom

from course.util import *
from docx import *

from course.util import *


# Create your models here.

#课程类
# 课程编号（Course number）
# 总 学 时（Total hours）
# 课程名称（Course title）
# 课程类别（Course category）
# 执 笔 人（Writer）
# 先修课程 （Prerequisite courses）
# 学    分（credit）
# 实验/上机学时（Experimental hours）
# 英文名称（English name）
# 适用专业（Applied specialty）
# 审 核 人（Auditor）
class Course(GraphObject):
    __primarylabel__ = "Course"
    __primarykey__ = "name"
    # 属性
    name = Property()
    courseNumber = Property()
    totalHours = Property()
    courseCategory = Property()
    writer = Property()
    prerequisiteCourses = Property()
    credit = Property()
    experimentalHours = Property()
    englishName = Property()
    appliedSpecialty = Property()
    auditor = Property()
    #关系 先修
    Prerequisite = RelatedTo('Course')  #先修

#教学目标类
# 描述（describe）
# 达成途径（Ways to achieve）
# 主要判据（Main criteria）
# 贡献度（Contribution degree）
# 课程名称
class TeachingObjective(GraphObject):
    __primarylabel__ = "TeachingObjective"
    __primarykey__ = "name"
    # 属性
    id = Property()
    name = Property()
    describe = Property()
    waysToAchieve = Property()
    mainCriteria = Property()
    contributionDegree = Property()
    courseName = Property()
#     关系 贡献，达成
    contribution = RelatedTo('IndexPoint')
    reach = RelatedFrom('Course')

#指标点类
#主标签：指标点
# 主键：name
# 属性：name，describe
# 关系：subitem
class IndexPoint(GraphObject):
    __primarylabel__ = "IndexPoint"
    __primarykey__ = "name"
    #属性
    name = Property()
    describe = Property()
    #关系 子项
    subitem = RelatedTo('IndexPoint')


if __name__ == '__main__':
    graph = Graph("http://localhost:7474", username="neo4j", password='431879')

    index = IndexPoint()
    index.name = 'indexPoint22'
    index.describe = ' jisjfi'

    teach = TeachingObjective()
    teach.name = '1'
    teach.describe = 'stestatt'
    teach.waysToAchieve = 'way'
    teach.mainCriteria = 'sfef'

    teach.contribution.add(index)
    graph.push(teach)


