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


daily_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/0331/日数据.xlsx',sheet_name=0)
trans_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/0331/日转数据.xlsx',sheet_name=0)
daily_df.set_index('Date', inplace=True)
daily_df['workday'] = daily_df.index.map(lambda y: y.isoweekday())
daily_df['bool'] = 0
for num in range(0, len(daily_df.index.values) - 1, 1):
    if daily_df.iloc[num, -2] >= daily_df.iloc[num + 1, -2]:
        daily_df.iloc[num, -1] = 1

tradedf =  daily_df.resample('W', label='right').last()
tradedf.High = daily_df.High.resample('W', label='right').max().values
tradedf.Low = daily_df.Low.resample('W', label='right').min().values

final_df = pd.DataFrame(index = trans_df.iloc[:-1,0].values,columns = ['coe'])
final_df.coe = [ math.log(i) for i in list(tradedf.iloc[1:,0].values/trans_df.iloc[:-1,4].values)]
final_df['volcoe'] = [math.log(abs(i-1)) if i != 1 else 0 for i in
                   list(tradedf.iloc[1:,0].values/trans_df.iloc[:-1,4].values)]
final_df['hvol'] =[abs(i-1) for i in list(trans_df.iloc[1:,1].values/trans_df.iloc[:-1,4].values)]
final_df['lvol'] =[abs(i-1) for i in list(trans_df.iloc[1:,2].values/trans_df.iloc[:-1,4].values)]
final_df['slvol'] = 0
for i in range(0,len(final_df.index),1):
    final_df.iloc[i,-1] = math.log(final_df.iloc[i,-3]) if \
        final_df.iloc[i,-3]>final_df.iloc[i,-2] else math.log(final_df.iloc[i,-2])
final_df[['coe','volcoe','slvol']].to_csv(
    r'C:/Users/Tao/Desktop/投资/spy/0331/spy系数(分母全部换为下单前一日收盘价）.csv', index=True, encoding='utf_8_sig')
