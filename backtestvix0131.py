import pandas as pd
import time
import copy
import numpy as np
import datetime as dt
import math


uvxydf = pd.read_excel(r'C:/Users/Tao/Desktop/投资/uvxy/0131/日数据.xlsx',sheet_name = 1)
uvxydf = uvxydf.dropna(axis= 0)
uvxydf =uvxydf.reset_index(drop = True)
uvxydf.set_index(['Date'],inplace =True)
uvxydf = uvxydf.resample('W', label='right').last()
uvxydf.reset_index(drop =False,inplace = True)
uvxydf = uvxydf.iloc[4:-1,:]
uvxydf =uvxydf.reset_index(drop = True)

expdf = pd.read_excel(r'C:/Users/Tao/Desktop/投资/uvxy/0131/数据整理（以前一日收盘价为底）.xlsx',sheet_name = 2).iloc[3:,:]
expdf = expdf.dropna(axis= 0)
expdf =expdf.reset_index(drop = True)

opendf = pd.read_csv(r'C:/Users/Tao/Desktop/投资/uvxy/0131/开盘系数(分母全部换为下单前一日收盘价）.csv',
                     engine = 'python',encoding='utf_8_sig').iloc[4:-1,:]
opendf =opendf.dropna(axis= 0)
opendf =opendf.reset_index(drop = True)

#起始点
startt = uvxydf[pd.Series(str(pd.to_datetime(i)) for i in uvxydf.iloc[:,0]).str.contains('2018-12')].index.values[-1]
endt =  uvxydf[pd.Series(str(pd.to_datetime(i)) for i in uvxydf.iloc[:,0]).str.contains('2019-12')].index.values[-1]

totaldf = pd.DataFrame(index = uvxydf.iloc[startt:endt+1,0],columns =['uvxy','预测系数',
                                                                      '开盘系数','仓位','收益净值','总资产'])
totaldf.fillna(0,inplace = True)

#初始总资产为1
totaldf.iloc[0,-1] = 1
totaldf.iloc[:,0] = uvxydf.iloc[startt:endt+1,1].values
totaldf.iloc[:,1] = expdf.iloc[startt:endt+1,1].values
totaldf.iloc[:,2] = opendf.iloc[startt:endt+1,1].values

# 仓位计算，仓位恒为0.05，预测系数小于等于开盘价*0.96为put，预测系数大于开盘价*1.04为call
for i in range(0, len(totaldf.index.values), 1):
    if totaldf.iloc[i , 1] <= totaldf.iloc[i , 2]*0.96:
        totaldf.iloc[i, 3] = -0.5
    elif totaldf.iloc[i , 1] >= totaldf.iloc[i , 2]*1.04:
        totaldf.iloc[i, 3] = 0.5

# 收益净值计算，-100% 与 （（1.1-0.12）*上周uvxy价格 - 本周uvxy价格）/（0.12*上周uvxy价格） 的最大值
for i in range(0, len(totaldf.index.values)-1, 1):
    if totaldf.iloc[i, 3] == -0.5:
        totaldf.iloc[i, 4] = -max(-1, (0.96 * totaldf.iloc[i, 0] - totaldf.iloc[i + 1, 0]) / (
                0.14 * totaldf.iloc[i, 0]))
    elif totaldf.iloc[i, 3] == 0.5:
        totaldf.iloc[i, 4] = max(-1, (totaldf.iloc[i + 1, 0] - 1.04 * totaldf.iloc[i, 0] ) / (
                0.14 * totaldf.iloc[i, 0]))

# 总资产计算，本期总资产=上期总资产*(1+ 本期收益率*上期仓位)
for i in range(1, len(totaldf.index.values), 1):
    totaldf.iloc[i, 5] = totaldf.iloc[i - 1, 5] * (1 + totaldf.iloc[i-1, 3] * totaldf.iloc[i - 1, 4])

totaldf.to_excel(r'C:/Users/Tao/Desktop/投资/uvxy/pool/2019年开始(p_vol_both_0131）.xlsx' , index=True, encoding='utf_8_sig')
# print(totaldf)
