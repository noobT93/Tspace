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
import pandas_datareader.data as web
import statsmodels.api as sm
from stepwise_regression import step_reg

#数据总跨度为2015.4.1-2020.7.1
#测试集2018.7.1-2020.7.1
#训练集2015.4.1-2020.4.1

start = dt.datetime(2015,4,1)
end = dt.datetime(2020,7,1)

writer = pd.ExcelWriter(r'C:/Users/Tao/Desktop/投资/final/基本数据.xlsx')
for i in ['spy','^VIX','^DJI','^IXIC','^GSPC','^HSI','sh','sz','cy']:
    if i not in ['sh','sz','cy']:
        locals()[i + 'df'] = web.DataReader(i, 'yahoo', start, end)
        locals()[i + 'df1'] = locals()[i + 'df'].resample('W', label='right').last()
        locals()[i + 'df1'] = locals()[i + 'df1'].drop(['Open', 'Adj Close'], axis=1)
        locals()[i + 'df1']['High'] = locals()[i + 'df'].High.resample('W', label='right').max()
        locals()[i + 'df1']['Low'] = locals()[i + 'df'].Low.resample('W', label='right').min()
        locals()[i + 'df1']['highvol'] = 0
        locals()[i + 'df1']['lowvol'] = 0
        for wnum in range(0, len(locals()[i + 'df1'].index), 1):
            aaa = locals()[i + 'df1'].index[wnum - 1] if wnum != 0 else pd.to_datetime(
                locals()[i + 'df'].index.values[0])
            bbb = locals()[i + 'df1'].index[wnum]
            tlist = [datei for datei in locals()[i + 'df'].index.values if aaa <= pd.to_datetime(datei) < bbb]
            try:
                highd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :].High == locals()[i + 'df1'].iloc[
                    wnum, 0]].index.values[0]
                locals()[i + 'df1'].iloc[wnum, -2] = locals()[i + 'df'].loc[highd, 'Volume']
            except:
                locals()[i + 'df1'].iloc[wnum, 3:] =np.nan
            try:
                lowd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :].Low == locals()[i + 'df1'].iloc[
                    wnum, 1]].index.values[0]
                locals()[i + 'df1'].iloc[wnum, -1] = locals()[i + 'df'].loc[lowd, 'Volume']
            except:
                locals()[i + 'df1'].iloc[wnum, 3:] = np.nan
        locals()[i + 'df1'].to_excel(writer, sheet_name=i)
    else:
        locals()[i + 'df'] = pd.read_excel(r'C:/Users/Tao/Desktop/投资/final/%s.xls' % i )
        locals()[i + 'df'].时间 = locals()[i + 'df'].时间.apply(lambda y: dt.datetime.strptime(y[:-2], '%Y-%m-%d'))
        locals()[i + 'df'].set_index('时间', inplace=True)
        start_a = [a for a in locals()[i + 'df'].index if a >= start ][0]
        end_a = [a for a in locals()[i + 'df'].index if a <= end ][-1]
        locals()[i + 'df'] = locals()[i + 'df'].loc[start_a:end_a,:]
        locals()[i + 'df1'] = locals()[i + 'df'].resample('W', label='right').last()
        locals()[i + 'df1']['High'] = locals()[i + 'df'].最高.resample('W', label='right').max()
        locals()[i + 'df1']['Low'] = locals()[i + 'df'].最低.resample('W', label='right').min()
        locals()[i + 'df1']['Close'] = locals()[i + 'df'].收盘.resample('W', label='right').last()
        locals()[i + 'df1']['Volume'] = locals()[i + 'df']['总手(万)'].resample('W', label='right').last()
        locals()[i + 'df1']['highvol'] = 0
        locals()[i + 'df1']['lowvol'] = 0
        locals()[i + 'df1'] = locals()[i + 'df1'][['High','Low','Close','Volume','highvol','lowvol']]
        for wnum in range(0, len(locals()[i + 'df1'].index), 1):
            aaa = locals()[i + 'df1'].index[wnum - 1] if wnum != 0 else pd.to_datetime(
                locals()[i + 'df'].index.values[0])
            bbb = locals()[i + 'df1'].index[wnum]
            tlist = [datei for datei in locals()[i + 'df'].index.values if aaa <= pd.to_datetime(datei) < bbb]
            try:
                highd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :].最高 == locals()[i + 'df1'].iloc[
                    wnum, 0]].index.values[0]
                locals()[i + 'df1'].iloc[wnum, -2] = locals()[i + 'df'].loc[highd, '总手(万)']
            except:
                locals()[i + 'df1'].iloc[wnum, :3] = locals()[i + 'df1'].iloc[wnum - 1, 2] if wnum != 0 else np.nan
                locals()[i + 'df1'].iloc[wnum, 3:] = locals()[i + 'df1'].iloc[wnum - 1, 3]
            try:
                lowd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :].最低 == locals()[i + 'df1'].iloc[
                    wnum, 1]].index.values[0]
                locals()[i + 'df1'].iloc[wnum, -1] = locals()[i + 'df'].loc[lowd, '总手(万)']
            except:
                locals()[i + 'df1'].iloc[wnum, :3] = locals()[i + 'df1'].iloc[wnum - 1, 2] if wnum != 0 else np.nan
                locals()[i + 'df1'].iloc[wnum, 3:] = locals()[i + 'df1'].iloc[wnum - 1, 3]
        locals()[i + 'df1'].to_excel(writer, sheet_name=i)
writer.save()

