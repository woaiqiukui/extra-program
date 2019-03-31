# -*- coding: utf-8 -*-
"""
Created on Fri Feb 15 16:36:27 2019

@author: hxh
"""

import xlrd
import random
from xlutils.copy import copy
import string

class influence:
    def __init__(self,id):
        workbook = xlrd.open_workbook('D:/xianyupro/excel1.xls')
        sh=workbook.sheet_by_index(0)
        nrows=sh.nrows
        temp=0#计数器
        for i in range(nrows):
            if (id==sh.cell(i,0).value):
                temp+=1
                print(id,"已存在")
                break
        if (temp==0):
            print(id,"不在现有表单中,已将其加入")
            if(id[0]=='d'):
                print("该用户的身份是医生")
                #将影响力初始化为70
                #打开excel,并添加
                w=xlrd.open_workbook('D:/xianyupro/excel1.xls')
                new=copy(w)
                sheet=new.get_sheet(0)
                sheet.write(nrows,0,id)#ID
                sheet.write(nrows,1,70)#影响力
                sheet.write(nrows,2,'r,w')#权限
                new.save('D:/xianyupro/excel1.xls')

            elif (id[0]=='p'):
                print("该用户的身份是病人")
                w=xlrd.open_workbook('D:/xianyupro/excel1.xls')
                new=copy(w)
                sheet=new.get_sheet(0)
                sheet.write(nrows,0,id)#ID
                #方便实验决定使用随机数来模拟标签相似性
                #初始影响力由相似性直接决定，初始最高40
                rand=random.randint(0,100)
                if (rand>40):
                    rand=40#超过40就设为40，后期通过读操作增加影响力
                if (rand>29):#病友间的影响力划分：0-30；30-50
                    sheet.write(nrows,2,'r')#30-50的权限为r,只读
                else:
                    sheet.write(nrows,2,'n')#0-29的权限为n,没有读写权限
                sheet.write(nrows,1,rand)
                new.save('D:/xianyupro/excel1.xls')
                
                
    def calculate(self,id):#从表单中返回权限并选择操作，读或写对影响力增减产生作用
        if (id[0]=='p'):#id是病友
            workbook = xlrd.open_workbook('D:/xianyupro/excel1.xls')
            sh=workbook.sheet_by_index(0)
            nrow=sh.nrows
            position,influence=0,0
            for i in range(nrow):
                if (sh.cell(i,0).value==id):
                    position=i#找到该id的行数
                    break
            priority=sh.cell(position,2).value
            print("该用户的权限是：",priority)
            #为模拟实验随机选择权限中的一个操作进行
            if(priority=='r'):#进行读操作
                print("当前进行的操作是：",priority)
                influence=sh.cell(position,1).value#获取当前的影响力
                influence+=1#读操作影响力加1
                if (influence>50):
                    influence=50#病友影响力上限为50
                w=xlrd.open_workbook('D:/xianyupro/excel1.xls')
                new=copy(w)
                sheet=new.get_sheet(0)
                if (influence>29):#影响力超过30即可获得r权限
                    sheet.write(position,2,"r")
                sheet.write(position,1,influence)
                new.save('D:/xianyupro/excel1.xls')
            else:#没有读写权限
                print("当前没有权限进行操作")
                
        elif(id[0]=='d'):#当前id为医生
            workbook = xlrd.open_workbook('D:/xianyupro/excel1.xls')
            sh=workbook.sheet_by_index(0)
            nrow=sh.nrows
            position,influence=0,0
            for i in range(nrow):
                if (sh.cell(i,0).value==id):
                    position=i#找到该id的行数
                    break
            priority=sh.cell(position,2).value
            print("该用户的权限是：",priority)
            #为模拟实验随机选择权限中的一个操作进行
            if (priority=='r'):
                print("当前进行的操作是：",priority)
                influence=sh.cell(position,1).value#获取当前的影响力
                influence+=1#读操作影响力加1
                if (influence>100):
                    influence=100#医生影响力上限为100
                w=xlrd.open_workbook('D:/xianyupro/excel1.xls')
                new=copy(w)
                sheet=new.get_sheet(0)
                if (influence>49):#影响力超过50即可获得r,w权限
                    sheet.write(position,2,"r,w")
                sheet.write(position,1,influence)
                new.save('D:/xianyupro/excel1.xls')
            else:#选择r或者w操作
                i=random.randint(1,2)#得到一个随机数来决定进行r还是w操作
                if (i==1):#进行r操作
                    print("当前进行的操作是：r")
                    influence=sh.cell(position,1).value#获取当前的影响力
                    influence+=1#读操作影响力加1
                    if (influence>100):
                        influence=100#医生影响力上限为100
                    w=xlrd.open_workbook('D:/xianyupro/excel1.xls')
                    new=copy(w)
                    sheet=new.get_sheet(0)
                    if (influence>49):#影响力超过50即可获得r,w权限
                        sheet.write(position,2,"r,w")
                    sheet.write(position,1,influence)
                    new.save('D:/xianyupro/excel1.xls')
                else:#进行w操作
                    print("当前进行的操作是：w")
                    print("您是否满意该医生的操作？满意请选1，不满意请选2，您的选择将决定了他对您的影响力")
                    a=0
                    while True:
                        a=input()#从键盘接受一个数字
                        a=int(a)
                        if (a==1 or 2):
                            break
                        print("您的输入有误，请重新输入")
                    if (a==1):#满意医生的操作，则医生影响力加2
                        print("您满意医生的操作，医生的影响力加2")
                        influence=sh.cell(position,1).value#获取当前的影响力
                        influence+=2#影响力加2
                        if (influence>100):
                            influence=100#医生影响力上限为100
                        w=xlrd.open_workbook('D:/xianyupro/excel1.xls')
                        new=copy(w)
                        sheet=new.get_sheet(0)
                        if (influence>49):#影响力超过50即可获得r,w权限
                            sheet.write(position,2,"r,w")
                        sheet.write(position,1,influence)
                        new.save('D:/xianyupro/excel1.xls')
                    else:#不满意医生的操作，使得医生的影响力减2
                        print("您不满意医生的影响力，医生的影响力减2")
                        influence=sh.cell(position,1).value#获取当前的影响力
                        influence-=2#影响力减2
                        if (influence<30):
                            influence=30#医生影响力下限为30
                        w=xlrd.open_workbook('D:/xianyupro/excel1.xls')
                        new=copy(w)
                        sheet=new.get_sheet(0)
                        if (influence<51):#影响力低于50即失去获得w权限
                            sheet.write(position,2,"r")
                        else:
                            sheet.write(position,2,"r,w")
                        sheet.write(position,1,influence)
                        new.save('D:/xianyupro/excel1.xls')


if __name__=='__main__':
    while True:
        print ("开始模拟实验")
        str=(lambda x:'p' if x==0 else 'd')(random.randint(0,1)) 
        str=str+string.digits[random.randint(0,9)]#随机生成d/p+（0-9）的字符串
        in1=influence(str)
        in1.calculate(str)
        print("是否需要暂停？1 暂停，2 继续")
        x=0
        while True:
            a=int(input())
            if (a==1):
                x=1
                break
            if (a==2):
                x=2
                break
            else:
                print("您的输入有误，请重新输入")
        if(x==1):
            break