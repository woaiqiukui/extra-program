# -*- coding: utf-8 -*-
"""
Created on Fri Mar  1 22:20:53 2019

@author: hxh
"""

import warnings
import itertools
import random
import xlrd
from xlutils.copy import copy
import re
from testtt import test1
from testtt import test2
from testtt import test3
from testtt import upgrade
from testtt import repeat

def author(x):#生成全部可能的操作，x为拥有的权限
    a=[]
    for i in range (len(x)):
        a.append(list(itertools.combinations(x,i+1)))
    return a#返回一个二维列表，存放全部可能的操作

def action(list):#随机选取一个操作,list为存放所有操作的列表
    a=random.randint(0,len(list[0])-1)#随机选择第一维
    act=list[a][random.randint(0,len(list[a])-1)]#在第二维中随机选择一个操作
    print("该用户随机执行的操作是：",act)#输出执行的操作
    return act#返回执行的操作

def num(aut,act,row,j,name,last):#统计操作出现的次数,act是操作，row是行数,j是列数,name是表名,last是之前的频率
    temp=0#用于计算权限被使用到的次数
    if (row==1):#如果是在第一行
        for i in act:
            if aut in i:#表示该权限被使用到了
                temp+=1
        return (temp/row)#返回的是一个频率
    else:
        for i in range (len(act)):
            if (aut==act[i]):#表示该权限被使用到了
                temp+=1+last*(row-1)
                return (temp/row)
        temp=last*(row-1)
        return (temp/row)


def compare(act,aut,name,i):#比较操作和权限的相似度,act是操作，aut是权限
    #比较的方法是比较操作和权限的长度
    #因为不考虑出现超出权限之外的操作，所以当长度相等时即权限等于操作
    #操作权限较小则是有未使用的权限
    if (i%5==0 and i!=0):#每五次进行一次权限更新
        aut=upgrade(aut,name)
    act=list(action(author(aut)))
    act=str(act)
    aut=str(aut)
    act=re.findall('\d+', act)#提取字符串中的数字
    aut=re.findall('\d+', aut)#同上
    l=len(aut)#权限的数量
    temp=0#设置一个计数器，查看用户表是否已经存在
    wb=xlrd.open_workbook("D:/xianyupro/author.xls")
    newb=copy(wb)
    sheetnames=wb.sheet_names()#获取已经在excel中的全部表名，用于比较是否插入新表
    for s in sheetnames:
        if (s==name):
            wbsheet=newb.get_sheet(name)                                             
            nrows=len(wbsheet.rows)
            wbsheet.write(nrows,0,aut)
            wbsheet.write(nrows,1,act)
            j=0
            for i in aut:
                n=num(i,act,nrows,j,name,test1(int(i)))
                test2(int(i),n)
                i=int(i)
                wbsheet.write(nrows,i+1,n)
                j+=1
            temp+=1
            break
    if (temp==0):
        wbsheet=newb.add_sheet(name)
        wbsheet.write(0,0,"权限")
        wbsheet.write(0,1,"操作")
        for i in range(4):#有多少权限就设置多少频率
            j=str(i+1)
            wbsheet.write(0,i+2,"权限"+j)
        nrows=len(wbsheet.rows)
        wbsheet.write(nrows,0,aut)
        wbsheet.write(nrows,1,act)
        j=0
        for i in aut:
            n=num(i,act,nrows,j,name,test1(int(i)))
            test2(int(i),n)
            i=int(i)
            wbsheet.write(nrows,i+1,n)
            j+=1
    newb.save("D:/xianyupro/author.xls")
    print(aut)
    return aut#可能权限会发生改变，需要返回新的权限
    
if __name__=='__main__':
    warnings.filterwarnings("ignore")
    temp=[]
    while True:
        name=input('输入用户名')
        if (repeat(name)==1):
            break
        else:
            print("该用户已存在，请输入其他用户名")
    print("请输入你需要他拥有的权限，可选1,2,3,4")
    a=sorted(list(set(input())))#去除输入中的重复数字并且完成排序
    for i in range (len(a)):
        temp.append(a[i])
    print("该用户拥有的权限是：",temp)
    for i in range(300):
        act=list(action(author(temp)))
        temp=compare(act,temp,name,i)
    test3()#清零