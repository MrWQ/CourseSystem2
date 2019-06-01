import os
from django.contrib import messages
from docx import *
from py2neo import Graph
from course.dao import *
from course.util import *
from django.shortcuts import render
from django.http import HttpResponse, HttpResponseRedirect
from course.models import Course
from course.models import *
import json
from course.utilForIndexpoint import createAndSaveRelationToDB

# 全局变量
graph = Graph("http://localhost:7474", username="neo4j", password='431879')
uploadPath = 'upload\course'
dirPath = os.path.abspath('.') + '\\' + uploadPath

# 测试返回json数据
def testjson(request):
    resp = {'errorcode': 100, 'detail': 'Get success'}
    return HttpResponse(json.dumps(resp), content_type="application/json")

# 中间跳转
def tip(request):
    message = request.GET.get('str')
    context ={}
    context['message'] = message
    return render(request,'base.html',context)

# 基础风格模板
def base(request):
    return render(request, 'base.html')

# 主页视图
# 根据设置的资源不同，动态调整主页
def index(request):
    resources = [
        {'a': '/course/index/', 'img': '/static/img/1.jpg', 'h1': '主页','p': '课程体系知识图谱系统' },
        {'a': '/course/upload/', 'img': '/static/img/2.jpg', 'h1': '上传文件','p': '上传课程大纲' },
        {'a': '/course/files/', 'img': '/static/img/3.jpg', 'h1': '文件管理','p': '管理上传的文件' },
        {'a': '/course/data/', 'img': '/static/img/4.jpg', 'h1': '信息展示','p': '课程体系知识图谱信息展示' },
    ]
    context = {}
    context['resources'] = resources
    return render(request, 'index.html',context)

# 批量上传
def upload(request):
    if request.method == 'GET':
        err = {}
        err['err'] = 0
        err['message'] = 'not upload'
        return render(request, 'upload.html', err)
    elif request.method == 'POST':
        uploadFiles = request.FILES  # 获取上传文件
        for name,uploadFile in uploadFiles.items():
            # uploadPath = 'upload\course\\'  # 文件上传路径
            fileName = uploadFile.name  # 文件名
            # 对文件名特殊字符进行处理
            fileName = str(fileName).replace('+','加')
            fileTemp = str(fileName).split('.')
            if len(fileTemp)>2:
                fileName = str(fileName).replace('.','点',len(fileTemp)-2)
            # dirPath = os.path.abspath('.') + '\\' + uploadPath  #上传目录绝对路径
            # 如果存在同名文件，修改现在上传文件的文件名
            while (os.path.exists(dirPath + '\\' + fileName)):
                temp = str(fileName).split('.')
                fileName = temp[len(temp) - 2] + '(1).' + temp[len(temp) - 1]
            #     保存文件
            with open(os.path.join(uploadPath, fileName), 'wb') as f:
                for line in uploadFile.chunks():
                    f.write(line)
        err = {}
        err['err'] = 1
        err['message'] = '上传成功'
        return render(request, 'upload.html', err)


# 上传文件信息展示
# 根据请求参数不同完成对上传的文件的操作
def files(request):
    context ={}
    context['message'] = '文件管理页面'
    context['data']= getFileInformation(dirPath)
    return render(request, 'files.html', context)
# 已上传文件的删除功能
def delete(request):
    deletefile = request.GET['filename']
    deletefile = deletefile.replace('../','')
    deletefile = deletefile.replace('..\\','')
    filePath = dirPath + '\\' + deletefile
    if os.path.exists(filePath):
        try:
            os.remove(filePath )
            message = '删除成功'
        except Exception as e:
            message = '删除成功'
            print(e)
    else:
        message = '文件未找到'
    return HttpResponse(message)
# 已上传文件doc转换为docx功能
def change(request):
    changefile = request.GET['filename']
    filePath = dirPath + '\\' + changefile
    if os.path.exists(filePath):
        if getFileType(filePath) == 'doc':
            filePath = docSaveToDocx(filePath)
            message = 'doc 成功转换为docx'
        else:
            message = '不是doc文件不能转换'
    else:
        message = '文件未找到'
    return HttpResponse(message)

def update(request):
    updatefile = request.GET['filename']

    filePath = dirPath + '\\' + updatefile
    if os.path.exists(filePath):
        try:
            saveRelationToDB(filePath,graph)
            message = '更新文件数据到数据库成功'
        except Exception as e:
            print(e)
            message = '更新文件数据到数据库失败'
    else:
        message = '文件未找到'
    return HttpResponse(message)

# 预先创建指标点/毕业要求
def start():
    file = os.path.abspath('.') + r'/static/毕业要求.docx'
    createAndSaveRelationToDB(file,graph)

def end():
    try:
        getData('MATCH p=()-->() delete p')
        getData('MATCH (n) delete n')
    except Exception as e:
        print(e)
# 数据页面
# 数据库的数据展示
def data(request):
    # 指标点初始化
    startkey  = request.GET.get('start')
    endkey = request.GET.get('end')
    keyword = request.GET.get('keyword')
    adjacent = request.GET.get('adjacent')
    if startkey:
        start()
        time.sleep(1)
    if endkey:
        end()
    context = {}
    context['links'] = getLinks()
    context['categories'] = getCategories()
    if keyword:
        context['datas'] = searchNodes(keyword)
    else:
        # 返回所有
        context['datas'] = getNodes()

    if adjacent and keyword:
        context['links'] = get_adjacent_nodes(keyword)
        context['datas'] = get_nodes_by_links(get_adjacent_nodes(keyword),getNodes())
    else:
        context['links'] = getLinks()

    return render(request, 'data.html', context)


