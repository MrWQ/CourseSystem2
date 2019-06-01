# -*- coding: utf-8 -*-
from py2neo import Graph, Node, Relationship
import json
from docx import *
from course.util import *
# 全局变量
# neo4j中的标签
lables = ['Course', 'TeachingObjective', 'IndexPoint']
graph = Graph("http://localhost:7474", username="neo4j", password='431879')


# 将对象列表中对象持久化到图形数据库
def saveListOfObjectToDB(DBgraph,objectlist):
    try:
        for object in objectlist:
            # print(type(object))
            DBgraph.push(object)
    except Exception as e:
        print(e)


# 将直接查询的数据格式化处理
def nodes_to_dict(nodes_data):
    data_list = []
    for data_l in nodes_data:
        data_n = data_l['n']
        data_dict = dict(data_n)
        data_list.append(data_dict)
    return data_list

# 执行cql获取数据
# 返回数据列表
def getData(cql):
    try:
        # 执行查询并转换为列表
        data = graph.run(cql).data()
    except Exception as e:
        print('cql run error')
        print('Neo Error: ',e)
    return data


#获取所有同标签节点name
def getNodeByLabel(label):
    cql = 'MATCH (n:'+ label +') RETURN n'
    try:
        data = getData(cql)
        data_list = nodes_to_dict(data)

    except Exception as e:
        print('input lable not found')
        print(e)
    return data_list


def searchNodeByLable(lable,keyword):
    cql = "match(n:" + str(lable) + ") where n.name =~'.*" + str(keyword) + ".*' return n"
    try:
        data = getData(cql)
        data_list = nodes_to_dict(data)

    except Exception as e:
        print('input lable not found')
        print(e)
    return data_list
# 标签，对应前端，图例
def getCategories():
    categories = []
    for lable in lables:
        categories.append({'name':lable})
    return categories

# 从数据库中获取所有节点
def getNodes():
    # name：节点名称 取neo4j中的节点名称
    # category:对应图例的索引，必须为整数值。
    # 名称和值都可以展示
    # symbolSize设置节点大小
    # nodeinfo 存放详细信息
    nodes = []
    for lable in lables:
        data_list = getNodeByLabel(lable)
        for data in data_list:
            node = {'name': '', 'category': '', 'nodeinfo': {}}
            node['category'] = lables.index(lable)
            node['name'] = data['name']

            node['nodeinfo'] = data

            nodes.append(node)
    nodes.sort(key=lambda node: node['name'])
    return nodes


def searchNodes(keyword):
    nodes = []
    for lable in lables:
        data_list = searchNodeByLable(lable, keyword)
        for data in data_list:
            node = {'name': '', 'category': '', 'nodeinfo': {}}
            node['category'] = lables.index(lable)
            node['name'] = data['name']
            node['nodeinfo'] = data
            nodes.append(node)
    nodes.sort(key=lambda node: node['name'])
    return nodes


# 返回相邻节点
def get_adjacent_nodes(node_name):
    links = []
    cql = 'MATCH p=({name:"'+ str(node_name) + '"})-->() RETURN p'
    dataList = getData(cql)
    for data in dataList:
        recoad = data['p']
        for relation in recoad.relationships:
            link = {'source': '', 'target': '', 'value': ''}
            link['source'] = relation.start_node['name']
            link['target'] = relation.end_node['name']
            links.append(link)
    cql = 'MATCH p=()-->({name:"'+ str(node_name) + '"}) RETURN p'
    dataList = getData(cql)
    for data in dataList:
        recoad = data['p']
        for relation in recoad.relationships:
            link = {'source': '', 'target': '', 'value': ''}
            link['source'] = relation.start_node['name']
            link['target'] = relation.end_node['name']
            links.append(link)
    return links


# 从数据库中获取所有边
def getLinks():
    # source:源节点 的名称
    # target：目标节点 的名称
    # value：可有可无， 取neo4j中的边的值
    # 方向可以展示，单向，可覆盖方向。 例：2》3和3》2 ，最终展示结果是后面方向覆盖前面方向
    links = []
    cql = 'MATCH p=()-->() RETURN p'
    dataList = getData(cql)
    for data in dataList:
        recoad = data['p']
        for relation in recoad.relationships:
            link = {'source': '', 'target': '', 'value': ''}
            link['source'] = relation.start_node['name']
            link['target'] = relation.end_node['name']
            links.append(link)
    return links


def get_nodes_by_links(linke_list, nodes_list):
    name_list = []
    nodes = []
    for link in linke_list:
        source = link['source']
        target = link['target']
        name_list.append(source)
        name_list.append(target)
    name_list = list(set(name_list))
    for node in nodes_list:
        if node['name'] in name_list:
            nodes.append(node)
    return nodes

if __name__ == '__main__':

    name = '1'
    courseName = 'Java Web技术'
    cql = 'MATCH (n:TeachingObjective {name:"'+ name +'",courseName:"'+ courseName +'"})  RETURN n '
    cql = 'MATCH (n:TeachingObjective {name:"'+ name +'",courseName:"'+ courseName +'"})  delete n '
    cql = 'match (n) return n.name limit 10'
    cql = 'MATCH p=()-[r:REACH]->() RETURN p'
    cql = 'match (n:Course) return n LIMIT 10'
    # 模糊查询
    cql = "match(n:Course) where n.name =~'.*JAVA.*' return n"
    # cql = 'MATCH (n:course) RETURN n LIMIT 25'
    # cql = 'create (n:course { tit: "javatest",coursename:"标题"}) return n '
    # cql = 'MATCH (n { tit: "javatest",coursename:"标题"}) RETURN n LIMIT 25'
    # cql = 'create (n:course '+ jsonToCQLjson(courseTable) +') return n'


    # print(getNodeByLabel("Course"))
    # print(getCategories())
    # print(getNodes())
    # print(getLinks())
    # print(getData(cql))
    # print(searchNodes(1))
    links =[{'source': '交互式软件系统设计', 'target': '交互式软件系统设计1', 'value': ''}, {'source': '交互式软件系统设计1', 'target': '3.3', 'value': ''}]
    nodes = getNodes()
    print(get_nodes_by_links(links,nodes))