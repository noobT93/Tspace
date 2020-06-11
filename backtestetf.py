import pandas as pd
import time
import copy
import numpy as np
import datetime as dt
import math

#预测系数为预测t+1周开盘/t周收盘
expdf = pd.read_excel(r'C:/Users/Tao/Desktop/投资/HS300/w1231/数据整理.xlsx',sheet_name = 3).iloc[2:,:]
expdf = expdf.dropna(axis= 0)
expdf =expdf.reset_index(drop = True)

#开盘系数为t+1周开盘/t周收盘，etf系数为t+2周开盘/t周收盘
etfdf = pd.read_excel(r'C:/Users/Tao/Desktop/投资/HS300/w1231/日转周数据.xlsx',sheet_name = 0).iloc[3:-2,-2:]
etfdf =etfdf.dropna(axis= 0)
etfdf =etfdf.reset_index(drop = True)
etfdf.iloc[:,0] = etfdf.iloc[:,0].map(lambda y:math.exp(y))
etfdf.iloc[:,1] = etfdf.iloc[:,1].map(lambda y:math.exp(y))
etfdf['return'] = 0
etfdf.iloc[:,-1] = etfdf.iloc[:,-2].values/etfdf.iloc[:,-3].values

#起始点
'''startt = etfdf[etfdf.iloc[:,0].str.contains('2016-01')].index.values[0]
endt = etfdf[etfdf.iloc[:,0].str.contains('2017-01')].index.values[0]'''

totaldf = {}
totaldf['预测系数'] = expdf['自然对数'].values
totaldf['开盘系数'] = etfdf['开盘系数'].values
totaldf['仓位'] = 0
totaldf['收益净值'] = etfdf['return'].values
totaldf['总资产'] = 0
totaldf = pd.DataFrame(totaldf)
totaldf.iloc[0,-1] = 1

# 仓位计算，最低系数与预测系数比较
for i in range(0, len(totaldf.index.values), 1):
    if totaldf.iloc[i, 0]  >= max(totaldf.iloc[i, 1] ,1):
        totaldf.iloc[i, 2] = 1

# 总资产计算，本期总资产=上期总资产*(1+ 本期收益率*上期仓位)
for i in range(1, len(totaldf.index.values), 1):
    totaldf.iloc[i, -1] = totaldf.iloc[i - 1, -1] * (1 + (totaldf.iloc[i - 1, -2] - 1) * totaldf.iloc[i - 1, -3])

totaldf.to_excel(r'C:/Users/Tao/Desktop/投资/HS300/pool/2016-2019_p_vol_new1.xlsx', index=True, encoding='utf_8_sig')





# print(totaldf)
