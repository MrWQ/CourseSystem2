# -*- coding: utf-8 -*-  
# from django.contrib import admin
from django.conf.urls import handler404
from django.urls import path
from course import views

urlpatterns = [
    path('json/',views.testjson),
    path('',views.index),
    path('base/',views.base),
    path('index/', views.index),            #主页
    path('upload/', views.upload),          #上传
    path('files/', views.files),            #文件管理
    path('change/', views.change),              #文件管理-doc转docx
    path('update/', views.update),              #文件管理-更新文件内容到数据库
    path('delete/', views.delete),              #文件管理-删除文件
    path('data/', views.data),              #数据展示


]

