import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import datetime as dt
from chinese_calendar import is_workday
import copy
import numpy as np
import math
import scipy.stats as st
import csv
import random
import codecs
import threading
import time


daily_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/uvxy/0229/日数据.xlsx',sheet_name=1)
trans_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/uvxy/0229/日转数据.xlsx',sheet_name=1)
daily_df.set_index('Date', inplace=True)
daily_df['workday'] = daily_df.index.map(lambda y: y.isoweekday())
daily_df['bool'] = 0
for num in range(0, len(daily_df.index.values) - 1, 1):
    if daily_df.iloc[num, -2] >= daily_df.iloc[num + 1, -2]:
        daily_df.iloc[num, -1] = 1
tradelist = daily_df[daily_df['bool'] == 1].index.values
tradedf =  daily_df.resample('W', label='right').last()

final_df = pd.DataFrame(index = trans_df.iloc[:-1,0].values,columns = ['coe'])
final_df.iloc[:,0] =[ math.log(i) for i in list(tradedf.iloc[1:,0].values/trans_df.iloc[:-1,4].values)]
final_df.to_csv(r'C:/Users/Tao/Desktop/投资/uvxy/0229/uvxy系数(分母全部换为下单前一日收盘价）.csv', index=True, encoding='utf_8_sig')

test_df = pd.DataFrame(index = trans_df.iloc[:,0].values,columns = ['coe'])
test_df.iloc[:,0] = tradedf.iloc[:,0].values/trans_df.iloc[:,4].values
test_df.to_csv(r'C:/Users/Tao/Desktop/投资/uvxy/0229/开盘系数(分母全部换为下单前一日收盘价）.csv', index=True, encoding='utf_8_sig')