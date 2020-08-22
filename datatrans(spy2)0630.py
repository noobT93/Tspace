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


daily_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/下次用这个/日数据.xlsx',sheet_name=0)
daily_df.set_index('Date', inplace=True)

#周五
tradedf =  daily_df.resample('W', label='right').last()

tradedf.High = daily_df.High.resample('W', label='right').max().values
tradedf.Low = daily_df.Low.resample('W', label='right').min().values

final_df = pd.DataFrame(index = tradedf.index[:-1].values,columns = ['slvol'])
final_df['slvol'] = 0

for i in range(0,len(final_df.index)):
    final_df.iloc[i,0] = math.log(tradedf.iloc[i+1,1]/tradedf.iloc[i,3]) if tradedf.iloc[i+1,1]+ tradedf.iloc[
    i+1,2] >= 2*tradedf.iloc[i,3]  else math.log(tradedf.iloc[i+1,2]/tradedf.iloc[i,3])
#volcoe 下周开盘价波动，svol周内最大波动
final_df.to_csv(
    r'C:/Users/Tao/Desktop/投资/spy/下次用这个/spy系数(分母全部换为下单前一日收盘价）.csv', index=True, encoding='utf_8_sig')
