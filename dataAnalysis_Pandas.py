import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

##-------config info start------
TestStr = "酒店名称"
OUT_FILE_NAME = "newfile.xlsx"
IN_FILE = "充值.xls"
IN_FILE1 = "D__运行中集团版程序_圣地集团_CSGroup3_Report_export.xls"
HOTEL_LIST=["集团本部","长江半岛酒店","湄潭酒店","土城圣地客栈","遵义宾馆","凤凰酒店","侏罗纪酒店","大瀑布酒店","机场酒店","鑫达鑫酒店","星火酒店","余庆健康酒店"]
newcolum11 = ["酒店","卡号","客户姓名","充值代码","充值金额","充值时间","操作员","备注"]

##-------config info  end------

##io = pd.io.excel.ExcelFile(fpath)  #
##data =pd.read_excel(io, sheet_name='持仓明细') # 速度较快，
##data =pd.read_excel(fpath, sheet_name='持仓明细') #慢

##path = os.getcwd()
##print("当前目录：" + path)
##
df1 = pd.read_excel(IN_FILE)
##
##print("----------------------------------------------------------------------------")
##print("-----头部查看前10行，默认5行----")
##print(df1.head())
##print("----------------------------------------------------------------------------")
##print("-----tail:从尾部查看，默认5条----")
##print(df1.tail())
##print(df1.values)
##print("----------------------------------------------------------------------------")
##print("-----shape():元组形式展示，查看行数及列数----")
##print("总行数："+str(df1.shape[0])+";————总列数："+str(df1.shape[1]))
####print("----------------------------------------------------------------------------")
####print("总行数："+str(len(df1))+";————总列数：错，不正确显示"+str(len(df1.head())))
##print("----------------------------------------------------------------------------")
##print("-----info:查看索引, 数据类型和内存信息----")
####test = list(df1.info())
##print(df1.info())
##print("----------------------------------------------------------------------------")
##print("-----info:打印索引----")
##print(df1.columns)
print("----------------------------------------------------------------------------")
## 确定字符串A包含字符串B: if A.find(b) !=-1  ;  if b in A
postion = []  # 放置酒店名称的位置信息；
hotel = [] #放置酒店名称
for i in range(0,df1.shape[0]):
    tpstr = str(df1.iloc[i,0])
    if len(tpstr)>8:
        print(tpstr+",提取得到会员卡号："+tpstr[-8:])
    else:
        print(tpstr)
    if TestStr in str(tpstr):
        postion.append(i)
        hotel.append(tpstr[5:])
postion.append(df1.shape[0])  # 将最后一行的位置，加入位置清单，备用
hotel.append("this is the end")
print("----------------各酒店开头行的所在位置-----------------------")
print(postion)
print(hotel)

rowdata = [] #用于放置提取的数据
postionCount = len(postion) #确认当前列表数据个数
print("列表个数：:" + str(postionCount))
##for j in (0, postionCount-1):
##    rowdata = rowdata.append(df1.iloc[[postion[j]+1,postion[j+1]-1],[0,1,4,5,6,7,8]])
for j in range(0, postionCount-1): #从第一个酒店开始，分别将各自的相关信息找出
    low = postion[j]
    high = postion[j+1]
    print("低："+str(low)+"; 高："+str(high))
##    rowdata.append(list(df1.iloc[low,[0,1,4,5,6,7,8]]))
##    rowdata.append([])
    for k in range(low+1,high-1): #在某个酒店的数据中，从第一行开始，循环提取所有数据
##        if df1.iloc[k,0] !="":
##            print("当前的行位置：" + print(k)+"; 本行数据：" + df1.iloc[k,0])
        rowdata.append([hotel[j]]+list(df1.iloc[k,[0,1,4,5,6,7,8]]))
    rowdata.append([])
## 以上，打印出按酒店分组，获取的数据；
##print(rowdata)


##newindex = pd.Series([df1.iloc[0,1],df1.iloc[0,2],df1.iloc[0,4],df1.iloc[0,5],df1.iloc[0,6],df1.iloc[0,7],df1.iloc[0,8]])
newcolum = pd.Series(["酒店","卡号","客户姓名","充值代码","充值金额","充值时间","操作员","备注"])
##创建空表帧,列为需要的数据，或后期根据表数组创建
newfile = pd.DataFrame(rowdata,columns=newcolum)

writer = pd.ExcelWriter(OUT_FILE_NAME)
io = pd.io.excel.ExcelFile(OUT_FILE_NAME)
##for i in range(0, postionCount-1):  #根据新数据保存，可能会将当前已有，但本次没有的数据清理掉
##    tem1 = newfile.loc[newfile["酒店"]==hotel[i],]  #把新数据整理出；
##    print("------整理出来的新数据-----------")
##    print(tem1)
##    tem2 = pd.read_excel(io, sheet_name = hotel[i]) #把已有的数据读出来；
##    print(hotel[i]+":------已有的数据-----------")
##    print(tem2)
##    print(hotel[i])
##    tem2 = tem2.append(tem1)  #append 不是本地修改，而是创建副本，因此重新赋值
##    print(hotel[i]+":------合并后的的数据-----------")
##    print(tem2)
##    tem2.to_excel(writer,sheet_name=hotel[i],index=False)

for i in range(0, postionCount-1):  #每次均将文件重写一次
    tem1 = newfile.loc[newfile["酒店"]==hotel[i],]  #把新数据整理出；
    print("------整理出来的新数据-----------")
    print(tem1)
    tem2 = pd.read_excel(io, sheet_name = hotel[i]) #把已有的数据读出来；
    print(hotel[i]+":------已有的数据-----------")
    print(tem2)
    print(hotel[i])
    tem2 = tem2.append(tem1)  #append 不是本地修改，而是创建副本，因此重新赋值
    print(hotel[i]+":------合并后的的数据-----------")
    print(tem2)
    tem2.to_excel(writer,sheet_name=hotel[i],index=False)
    
    
newfile.to_excel(writer,OUT_FILE_NAME)
writer.save()

##
##计算各列数据总和并作为新列添加到末尾
##df['Col_sum'] = df.apply(lambda x: x.sum(), axis=1)
##
##计算各行数据总和并作为新行添加到末尾
##df.loc['Row_sum'] = df.apply(lambda x: x.sum())


##for i in range(8,12):
##    print("----------------------------------------------------------------------------")
##    print("-----打印"+str(i)+"行数据, 是series类型---")
##    print(df1.loc[i])

##df1.to_excel('遵义宾馆',sheet_name='遵义宾馆')


##df["newtes"] = range(1,len(df)+1)
##df.info()
##print(df["酒店"])
##print(df['酒店'],df['充值'])

##pd.series(data, index, dtype, copy)
##data: 数据； index: 索引； 

##
##data = np.array(['a','b','c','d'])
####s = pd.Series(data, index=[1,2,3,4])
####print(s)
##
####数据帧 pd.DataFrale(Data,index,columns,dtype)
##df = pd.DataFrame(data)
##print(df)
##
#### 字典创建帧
##dict1 = {
##    'name':["Arron","bb","cc","dd"],
##    "age":[12,3,4,5]
##    }
##print(---字典创建的数据帧---)
##df1 = pd.DataFrame(dict1, index=[1,2,3,4])
##print(df1)
##
####插入列，位置：2，列名zf, 内容
##df1.insert(2, 'dz', ['z','f','d',1])
##
##print(df1)

##print("---求和----")
##df = dfe.sum()
##print(df)
##print("----求标准差---")
##print(dfe.std())

##print("---选择第2行，se-rial格式---")
##print(dfe.iloc[1])
##print("-------")
##print("----选择第4行，dataFrame格式取数据---")
##print(dfe.iloc[3])
####print("----取消费一列的数据---")
####print(dfe.loc[:,"消费"])
####print("----取实收一列的数据---")
####print(dfe["实收"])
##
##
##print("----取第2、第10行，第2列的数据---")
##print(dfe.iloc[[2,10], 1]) # 选择2和3行, 1列所有数据
##
##
##print("----取指定 第 2行，第2列的数据---")
##print(dfe.iloc[2,1])

##ts = pd.Series(np.random.randn(100),index=pd.date_range("20201209",periods=100))
##print(ts)
##ts=ts.cumsum()
##plt.plot(ts)
##plt.show()




