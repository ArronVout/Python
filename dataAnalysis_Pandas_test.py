#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
import pandas as pd
import numpy as np
##import matplotlib.pyplot as plt

##-------config start------
TestStr = "酒店名称"
HOTEL_LIST=["集团本部",
            "长江半岛酒店",
            "湄潭酒店",
            "土城圣地客栈",
            "遵义宾馆",
            "凤凰酒店",
            "侏罗纪酒店",
            "大瀑布酒店",
            "机场酒店",
            "鑫达鑫酒店",
            "星火酒店",
            "余庆健康酒店"]
COL = ["酒店",
       "卡号",
       "客户姓名",
       "充值代码",
       "充值金额",
       "充值时间",
       "操作员",
       "备注"]

CONFIG = "config.xls"
IN_FILE = "InputInfo_12-18.xls"
AllList  = "./getedinfo/充值清单_all.xlsx"
DaliFile = "./getedinfo/每日充值清单_"+time.strftime("%Y-%m-%d",time.localtime())+".xlsx"  #生成当天的文件
##-------config end------

print("----start----"+DaliFile)

##writer = pd.ExcelWriter(AllList)
##io_all = pd.io.excel.ExcelFile(AllList)
##io_new = pd.io.excel.ExcelFile(IN_FILE)
##for hotel in HOTEL_LIST: #遍历所有酒店
##    try:
##        old = pd.read_excel(io_all,sheet_name=hotel)
##    except: #出错，说明没有这一个sheet,则新创建
##        old = pd.DataFrame([],columns = COL)
####        df.to_excel(writer,sheet_name=hotel,index=False)
##        
##    if hotel in IN_INFO: #如果有信数据，则合并
##        new = pd.read_excel(io_new,sheet_name=hotel)
##        old = old.append(new)
##    old.to_excel(writer,sheet_name=hotel,index=False)
##writer.save()

try:
    io_all = pd.io.excel.ExcelFile(AllList)
except:    #读取所有信息出错，则文件不存在，先创建一个文件
    io_all = pd.DataFrame([],columns=COL)
    io_all.to_excel(AllList,index=False)


#导出酒类卡的清单信息
write_config = pd.io.excel.ExcelFile(CONFIG)
##cardlist = pd.read_excel(CONFIG, sheet_name="cardlist",index=False)  #家里
cardlist = pd.read_excel(write_config, sheet_name="cardlist")   #办公室用，只能xls
##print(cardlist)
CardInfo = cardlist["卡号"]
##print("--------酒卡清单-------------》")
##print(CardInfo)

# 第一步：将需要解析的充值文件信息读出
df = pd.read_excel(pd.io.excel.ExcelFile(IN_FILE))
print(df)

##第二步： 处理读出的数据“逐行读取数据，进行处理，得到有充值的酒店的起始行，及对应酒店名称；
##数据格式，第一列的数据格式：酒店名称：湄潭酒店； 第二列：铂金卡/66601030
postion = []  # 放置有充值的酒店起始行所在文件中的位置；
hotel = []    # 放置对应起始位置处的酒店名称
for i in range(0,df.shape[0]):   #df.shape[0]: 行的数量，1：列的数量
    tpstr = str(df.iloc[i,0]) 
    if TestStr in str(tpstr):    #包含有"酒店名称"，则说明是某家的统计起始行
        postion.append(i)
        hotel.append(tpstr[5:])        
postion.append(df.shape[0])     # 将最后一行的位置，加入位置清单，供备用
hotel.append("this is the end")  # 最后一行对应的名字

##print("----------------打印：各酒店开头行的所在位置及对应的酒店-----------------------")

postionCount = len(postion)     #确认当前列表数据个数，即：有充值的酒店数量+1
##print("列表个数：:" + str(postionCount))

#从第一个酒店开始，分别将各自的相关信息找出
rowdata = []                    #用于放置提取的数据
for j in range(0, postionCount-1): 
##    low = postion[j]              #第j家有充值酒店的信息起始行，不含实际数据
##    high = postion[j+1]           #j+1家的，最后1个酒店时，是文档的最后一行
    for k in range(postion[j]+1,postion[j+1]-1): #第j家酒店的数据中，从第一行实际数据开始，循环提取所有数据
        newline = [hotel[j]]+list(df.iloc[k,[0,1,4,5,6,7,8]])
        print(newline)
        try:
            newline[1] = newline[1].split("/")[1]
        except:
            newline[1] = "Nan"
            print("-----")
        print(newline)
        rowdata.append(newline)  #提取需要的数据进行打包
    rowdata.append([]) #第j家数据取完后，增加一个空行

print(rowdata)
##增加本次提取的数据
GetedInfo = pd.DataFrame(rowdata,columns=COL)
print(GetedInfo)

if postionCount > 1:    ## 有数据，则创建对应数据的充值报表文件
##    DaliFile = "./getedinfo/充值日期_"+str(df.iloc[3,6])[0:10]+".xlsx"  #生成当天的文件
    DaliFile = "./getedinfo/充值日期_"+str(GetedInfo.iloc[0,5])[0:10]+".xlsx"  #生成当天的文件
    print("充值日期文件："+DaliFile)
    print(df.iloc[3,])
    writer_Dali = pd.ExcelWriter(DaliFile)
else:
##    print("无充值数据")
    exit()    
##print("----------------打印：各酒店开头行的所在位置-----------------------")
#建立输出引擎，将所有数据，归总到各酒店，并将统计日期的数据，单独形成报表
writer_all = pd.ExcelWriter(AllList)

#按当前已有酒店名称进行遍历，建立输出引擎，
##如果酒店的名称存在于最新得到的充值酒店列表hotel中，则需要读取输出表中已有数据，在后增加数据；然后导出；
##读取酒店的前期数据出错，则说明前期无该酒店的数据，则增加相应酒店的sheet页面，以酒店名字命名
try:
    AllInfo = pd.read_excel(io_all,sheet_name = "ALL")  #保存所有信息
except:
    AllInfo = pd.DataFrame([],columns = COL)
    AllInfo.to_excel(writer_all,sheet_name="ALL")
    
for hotelname in HOTEL_LIST:
    try:
        PreInfo = pd.read_excel(io_all,sheet_name = hotelname)  #获取所有酒店的放在一起的前期数据
    except:   #出错，说明没有这个酒店的前期数据，新建一个hotelname为名字的sheet
        PreInfo = pd.DataFrame([],columns = COL)
        PreInfo.to_excel(writer_all,sheet_name=hotelname)
        
##    print("----------------打印：当前酒店："+hotelname+"-----------------------")   
    if hotelname in hotel: #如果当前便利的酒店名字在充值酒店清单中，则有新数据需要合并，没有则保持原数据
        NewInfo = GetedInfo.loc[GetedInfo["酒店"] == hotelname,]  #每次保存按酒店名称提取的充值数据；
        PreInfo = PreInfo.append(NewInfo)   #数据帧对象的append 不是本地修改，而是创建副本，因此重新赋值
        AllInfo = AllInfo.append(NewInfo)
        try:
            Input = pd.read_excel(writer_DataFile,sheet_name = hotelname)  #获取hotelname酒店对应的前期数据
        except:
            Input = pd.DataFrame([],columns = COL)
        
        Input = Input.append(NewInfo)  #将当前遍历的酒店数据，增加到输入清单中，
        NewInfo.to_excel(writer_Dali,sheet_name = hotelname,index=False) #将统计日期的充值数据单独列出报表
        Input.to_excel(writer_all,sheet_name = hotelname,index=False)  #将统计日期的充值数据增加到报表
        
    #hotelname 酒店的数据合并处理完成，导出
    PreInfo = PreInfo.drop_duplicates("充值时间")   #去除重复项，防止对某天进行多次执行的情况
##    print("----------------去重复后--------------------------")
##    print(PreInfo)    
    PreInfo.to_excel(writer_all,sheet_name = hotelname,index=False)

AllInfo = AllInfo.drop_duplicates("充值时间")
AllInfo.to_excel(writer_all,sheet_name = "ALL",index=False)

writer_all.save()
writer_Dali.save()

print("----end----")
