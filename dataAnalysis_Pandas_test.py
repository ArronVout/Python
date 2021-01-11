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

CONFIG = "config.xlsx"
##IN_FILE = "./datas/InputInfo_20210109.xls"
IN_FILE = "./datas/InputInfo_12-21.xls"
AllList  = "./getedinfo/1.充值清单_all.xlsx"
WineList = "./getedinfo/1.酒卡充值清单_all.xlsx"
DaliFile = "./getedinfo/1.每日充值清单_"+time.strftime("%Y-%m-%d",time.localtime())+".xlsx"  #生成当天的文件
##-------config end------

print("----start----"+DaliFile)

try:  #读取所有充值历史信息出错，则文件不存在，先创建一个文件
    io_all = pd.read_excel(AllList,sheet_name="ALL",index=False)
except:
    print("--------打开ALL信息出错-------："+AllList)
    io_all = pd.DataFrame([],columns=COL)
    io_all.to_excel(AllList,sheet_name="ALL",index=False)
##print("--------当前的已有ALL信息--------:"+AllList)
##print(io_all)
    
try:  #读取酒卡充值历史信息出错，则文件不存在，先创建一个文件
    io_wine = pd.read_excel(WineList,index=False)
except:    
    print("--------打开酒类卡信息出错-------" + WineList)
    io_wine = pd.DataFrame([],columns=COL)
    io_wine.to_excel(WineList,index=False)

#导出酒类卡的清单信息
write_config = pd.io.excel.ExcelFile(CONFIG)
cardlist = pd.read_excel(CONFIG, sheet_name="cardlist",index=False)  #家里
##cardlist = pd.read_excel(write_config, sheet_name="cardlist")   #办公室，只能xls
WineCard = cardlist["卡号"].astype(str)
print("--------已办理的酒卡清单-------------》")

# 第一步：将需要解析的充值文件信息读出
df = pd.read_excel(pd.io.excel.ExcelFile(IN_FILE))

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

#从第一个酒店开始，分别将各自的相关信息找出
rowdata = []   #用于放置提取的数据
winedata= []   #保存酒卡相关信息
for j in range(0, postionCount-1):
    print("------正在处理酒店：------"+hotel[j])
    for k in range(postion[j]+1,postion[j+1]-1): #第j家酒店的数据中，从第一行实际数据开始，循环提取所有数据
        newline = [hotel[j]]+list(df.iloc[k,[0,1,4,5,6,7,8]])        
        try:  #提取会员卡号
            newline[1] = str(newline[1]).split("/")[-1]
        except:
            print("---提取会员卡号出错--")
            newline[1] = "Nan"
##        print(newline)
        if newline[1] in WineCard.tolist():
            winedata.append(newline) #酒卡消费单独保存
        else:
            rowdata.append(newline)  #提取需要的数据进行打包
    rowdata.append([]) #第j家数据取完后，增加一个空行

##本次提取到的普通卡充值数据，格式化
GetedInfo = pd.DataFrame(rowdata,columns=COL)

 #将新获取的酒卡充值信息，添加到当前已有信息后保存
writer_wine= pd.ExcelWriter(WineList)
io_wine = io_wine.append(pd.DataFrame(winedata,columns=COL))
io_wine = io_wine.drop_duplicates("充值时间")
io_wine.to_excel(writer_wine,index=False)
writer_wine.save() 

if postionCount > 1:    ## 有数据，则创建对应数据的充值报表文件
##    DaliFile = "./getedinfo/充值日期_"+str(df.iloc[3,6])[0:10]+".xlsx"  #生成当天的文件
    DaliFile = "./getedinfo/1.充值日期_"+str(GetedInfo.iloc[0,5])[0:10]+".xlsx"  #生成当天的文件
##    print("充值日期文件："+DaliFile)
##    print(df.iloc[3,])
    writer_Dali = pd.ExcelWriter(DaliFile)
else:
    print("无充值数据")    
    exit()    #没有充值信息，直接退出程序
##print("----------------打印：各酒店开头行的所在位置-----------------------")
#建立输出引擎，将所有数据，归总到各酒店，并将统计日期的数据，单独形成报表
writer_all = pd.ExcelWriter(AllList)

#按当前已有酒店名称进行遍历，建立输出引擎，
##如果酒店的名称存在于最新得到的充值酒店列表hotel中，则需要读取输出表中已有数据，在后增加数据；然后导出；
##读取酒店的前期数据出错，则说明前期无该酒店的数据，则增加相应酒店的sheet页面，以酒店名字命名    
for hotelname in HOTEL_LIST:
    try: 
        PreInfo = pd.read_excel(AllList,sheet_name = hotelname,index=False)  #获取所有酒店的放在一起的前期数据
    except:   #出错，说明没有这个酒店的前期数据，新建一个hotelname为名字的sheet
        print("--------读取酒店当前信息出错-------:"+hotelname)
        PreInfo = pd.DataFrame([],columns = COL)
        PreInfo.to_excel(writer_all,sheet_name=hotelname)
        
    if hotelname in hotel:  #如果当前便利的酒店名字在充值酒店清单中，则有新数据需要合并，没有则保持原数据
        NewInfo = GetedInfo.loc[GetedInfo["酒店"] == hotelname,]  #每次保存按酒店名称提取的充值数据；
        PreInfo = PreInfo.append(NewInfo)   #数据帧对象的append 不是本地修改，而是创建副本，因此重新赋值
        io_all = io_all.append(NewInfo)
        NewInfo.to_excel(writer_Dali,sheet_name = hotelname,index=False) #将统计日期的充值数据单独列出报表
##        try:
##            Input = pd.read_excel(writer_DataFile,sheet_name = hotelname)  #获取hotelname酒店对应的前期数据
##        except:
##            print("------读取单个酒店历史数据出错------酒店:")
##            Input = pd.DataFrame([],columns = COL)
##        
##        Input = Input.append(NewInfo)  #将当前遍历的酒店数据，增加到输入清单中，
##        Input.to_excel(writer_all,sheet_name = hotelname,index=False)  #将统计日期的充值数据增加到报表

##    PreInfo：对当期有消费的，则增加了当期数据，没有当期数据的则保留了历史数据
    PreInfo = PreInfo.drop_duplicates("充值时间")   #去除重复项后，保存到历史统计数据库中
    PreInfo.to_excel(writer_all,sheet_name = hotelname,index=False)
    #hotelname酒店完成，进入下一家

io_all = io_all.drop_duplicates("充值时间")
io_all.to_excel(writer_all,sheet_name = "ALL",index=False)

writer_all.save()
writer_Dali.save()

print("----end----")
