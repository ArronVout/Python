#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
import pandas as pd
import numpy as np
##import matplotlib.pyplot as plt

##-------config info start------
TestStr = "酒店名称"
IN_FILE = "InputInfo.xls"
IN_FILE1 = "InputInfo.xls"
##IN_INFO = ["湄潭酒店","遵义宾馆","凤凰酒店","机场酒店"]  #有测试信息的酒店
OUT_FILE_NAME = "会员卡充值清单_"+time.strftime("%Y-%m-%d",time.localtime())+".xlsx"  #生成当天的文件
HOTEL_LIST=["集团本部","长江半岛酒店","湄潭酒店","土城圣地客栈","遵义宾馆","凤凰酒店","侏罗纪酒店","大瀑布酒店","机场酒店","鑫达鑫酒店","星火酒店","余庆健康酒店"]
COL = ["酒店","卡号","客户姓名","充值代码","充值金额","充值时间","操作员","备注"]
TODAY_FILE = ""
##-------config info  end------

print("----start----")
##writer = pd.ExcelWriter(OUT_FILE_NAME)
##io_old = pd.io.excel.ExcelFile(OUT_FILE_NAME)
##io_new = pd.io.excel.ExcelFile(IN_FILE)
##for hotel in HOTEL_LIST: #遍历所有酒店
##    try:
##        old = pd.read_excel(io_old,sheet_name=hotel)
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
    io_old = pd.io.excel.ExcelFile(OUT_FILE_NAME)
except:  #出错，则文件没有，先创建一个文件
    io_old = pd.DataFrame([],columns=COL)
    io_old.to_excel(OUT_FILE_NAME,index=False)

io_new = pd.io.excel.ExcelFile(IN_FILE1)
df = pd.read_excel(io_new)  # 读取最新导出的充值信息

postion = []  # 放置有充值的酒店起始行所在文件中的位置；
hotel = []    # 放置对应起始位置处的酒店名称

##逐行读取数据，进行处理，得到有充值的酒店的起始行，及对应酒店名称；
##数据格式，第一列的数据格式
##酒店名称：湄潭酒店
##铂金卡/66601030
for i in range(0,df.shape[0]):   #df.shape[0]: 行的数量，1：列的数量
    tpstr = str(df.iloc[i,0]) 
    if TestStr in str(tpstr):    #包含有"酒店名称"，则说明是某家的统计起始行
        postion.append(i)
        hotel.append(tpstr[5:])
        
postion.append(df.shape[0])     # 将最后一行的位置，加入位置清单，供备用
hotel.append("this is the end")  # 最后一行对应的名字

print("----------------打印：各酒店开头行的所在位置及对应的酒店-----------------------")
print(postion)
print(hotel)

postionCount = len(postion)     #确认当前列表数据个数，即：有充值的酒店数量+1
##print("列表个数：:" + str(postionCount))

#从第一个酒店开始，分别将各自的相关信息找出
rowdata = []                    #用于放置提取的数据
for j in range(0, postionCount-1): 
##    low = postion[j]              #第j家有充值酒店的信息起始行，不含实际数据
##    high = postion[j+1]           #j+1家的，最后1个酒店时，是文档的最后一行
    for k in range(postion[j]+1,postion[j+1]-1): #第j家酒店的数据中，从第一行实际数据开始，循环提取所有数据
        rowdata.append([hotel[j]]+list(df.iloc[k,[0,1,4,5,6,7,8]]))  #提取需要的数据进行打包
    rowdata.append([]) #第j家数据取完后，增加一个空行

##以已提取的数据为基础，按标准格式创建空表帧 
GetedInfo = pd.DataFrame(rowdata,columns=COL)
##print(GetedInfo)
if postionCount > 1:  ## 有数据，则创建对应数据的充asdfasdfasdf值报表文件
##    DataFileName = "会员卡充值清单_"+time.strftime("%Y-%y-%d",time.localtime())+".xlsx"  #生成当天的文件
    DataFileName = "会员卡充值清单_"+str(df.iloc[1,6])[0:10]+".xlsx"  #生成当天的文件
    print(DataFileName)
    DataFile = pd.DataFrame([],columns=COL)
    DataFile.to_excel(DataFileName,index=False)
    print("创建当天的文件：" + DataFileName)
    writer_DataFile = pd.ExcelWriter(DataFileName)
else:
    print("无充值数据")
    exit()
    
print("----------------打印：各酒店开头行的所在位置-----------------------")

#建立输出引擎，将所有数据，归总到各酒店，并将统计日期的数据，单独形成报表
writer = pd.ExcelWriter(OUT_FILE_NAME)
##writer_DataFile = pd.ExcelWriter(DataFileName)

#按当前已有酒店名称进行遍历，建立输出引擎，
##如果酒店的名称存在于最新得到的充值酒店列表hotel中，则需要读取输出表中已有数据，在后增加数据；然后导出；
##读取酒店的前期数据出错，则说明前期无该酒店的数据，则增加相应酒店的sheet页面，以酒店名字命名
for hotelname in HOTEL_LIST:    
    try:
        PreInfo = pd.read_excel(io_old,sheet_name = hotelname)  #获取hotelname酒店对应的前期数据
    except:   #出错，说明没有这个酒店的前期数据，新建一个hotelname为名字的sheet
        print("----------------打印：输出文件不存在，新建一个-----------------------")
        PreInfo = pd.DataFrame([],columns = COL)
        
    print("----------------打印：当前酒店："+hotelname+"-----------------------")   
    if hotelname in hotel: #如果当前便利的酒店名字在充值酒店清单中，则有新数据需要合并，没有则保持原数据
        NewInfo = GetedInfo.loc[GetedInfo["酒店"] == hotelname,]  #按酒店名称，将所有充值数据提取；
        PreInfo = PreInfo.append(NewInfo)   #数据帧对象的append 不是本地修改，而是创建副本，因此重新赋值
        print("新充值信息----->")
        print(NewInfo)
        try:
            Input = pd.read_excel(writer_DataFile,sheet_name = hotelname)  #获取hotelname酒店对应的前期数据
        except:
##            print("meiyou sheet")
            Input = pd.DataFrame([],columns = COL)

##        NewInfo = np.abs(NewInfo) 当前是Series类型
##        sum1 = pd.Series(["","","","",NewInfo["充值金额"].cumsum(),"",""])
##        NewInfo.add(sum1)
##        NewInfo = NewInfo.add(pd.Series("","","",NewInfo["充值金额"].cumsum(),"",""))
        
        Input = Input.append(NewInfo)
                    
        print("添加后,增加求和的----->Input")
        print(Input)
        Input.to_excel(writer_DataFile,sheet_name = hotelname,index=False)  #将统计日期的充值数据写入当天的报表
        
    print("----------------整理后的数据信息添加后的:去重复前----->PreInfo-----------------------")
    print(PreInfo)    
    #hotelname 酒店的数据合并处理完成，导出
    PreInfo = PreInfo.drop_duplicates("充值时间")   #去除重复项，防止对某天进行多次执行的情况
    print("----------------去重复后--------------------------")
    print(PreInfo)    
    PreInfo.to_excel(writer,sheet_name = hotelname,index=False)
    
GetedInfo.to_excel(writer,OUT_FILE_NAME)
writer_DataFile.save()
writer.save()

print("----end----")
