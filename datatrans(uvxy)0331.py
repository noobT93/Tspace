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
from datetime import datetime

writer = pd.ExcelWriter(r'C:/Users/Tao/Desktop/投资/uvxy/0331/日转数据.xlsx')
for i in ['vix','uvxy','dji','ixic','spx','hsi','sh']:
    locals()[i+'df'] = pd.read_excel(r'C:/Users/Tao/Desktop/投资/uvxy/0331/日数据.xlsx',sheet_name=[
        'vix','uvxy','dji','ixic','spx','hsi','sh'].index(i))
    locals()[i + 'df']['Date'] = locals()[i + 'df']['Date'].map(lambda y:pd.to_datetime(y) )
    locals()[i + 'df'].set_index('Date',inplace = True)
    #获取每周最高价、最低价对应的交易量
    locals()[i + 'df2'] = locals()[i + 'df'].resample('W', label='right').last()
    locals()[i + 'df2']['Open'] = locals()[i + 'df']['Open'].resample('W', label='right').first()
    locals()[i + 'df2']['Close'] = locals()[i + 'df']['Close'].resample('W', label='right').last()
    locals()[i + 'df2']['Adj Close'] = locals()[i + 'df']['Adj Close'].resample('W', label='right').last()
    locals()[i + 'df2']['High'] = locals()[i + 'df']['High'].resample('W', label='right').max()
    locals()[i + 'df2']['Low'] = locals()[i + 'df']['Low'].resample('W', label='right').min()
    if i not in ['vix','hsi','sh']:
        locals()[i + 'df2']['highvol'] = 0
        locals()[i + 'df2']['lowvol'] = 0
        for wnum in range(0,len(locals()[i + 'df2'].index),1):
            aaa= locals()[i + 'df2'].index[wnum-1] if wnum != 0 else pd.to_datetime(locals()[i + 'df'].index.values[0])
            bbb = locals()[i + 'df2'].index[wnum]
            tlist = [datei for datei in locals()[i + 'df'].index.values if aaa<= pd.to_datetime(datei)<bbb]
            highd = locals()[i + 'df'].loc[tlist,:][locals()[i + 'df'].loc[tlist,:]['High']==locals()[i + 'df2'].iloc[
                wnum,1]].index.values[0]
            locals()[i + 'df2'].iloc[wnum,-2] = locals()[i + 'df'].loc[highd,'Volume']
            lowd = locals()[i + 'df'].loc[tlist,:][locals()[i + 'df'].loc[tlist,:]['Low']==locals()[i + 'df2'].iloc[
                wnum,2]].index.values[0]
            locals()[i + 'df2'].iloc[wnum,-1] = locals()[i + 'df'].loc[lowd,'Volume']
    if i not in ['hsi','sh']:
        # 将每周最后一个交易日的数据删除再取周数据
        locals()[i + 'df']['workday'] = locals()[i + 'df'].index.map(lambda y: y.isoweekday())
        locals()[i + 'df']['bool'] = 0
        for num in range(0, len(locals()[i + 'df'].index.values) - 1, 1):
            if locals()[i + 'df'].iloc[num, -2] >= locals()[i + 'df'].iloc[num + 1, -2]:
                locals()[i + 'df'].iloc[num, -1] = 1
        wvol = locals()[i + 'df']['Volume'].resample('W', label='right').last().values
        locals()[i + 'df'] = locals()[i + 'df'].drop(locals()[i + 'df'][locals()[i + 'df']['bool'] == 1].index.values,
                                                     axis=0)
        locals()[i + 'df'] = locals()[i + 'df'].iloc[:-1, :-2]
        locals()[i + 'df1'] = locals()[i + 'df'].resample('W', label='right').last()
        locals()[i + 'df1']['Open'] = locals()[i + 'df']['Open'].resample('W', label='right').first()
        locals()[i + 'df1']['Close'] = locals()[i + 'df']['Close'].resample('W', label='right').last()
        locals()[i + 'df1']['Adj Close'] = locals()[i + 'df']['Adj Close'].resample('W', label='right').last()
        locals()[i + 'df1']['High'] = locals()[i + 'df']['High'].resample('W', label='right').max()
        locals()[i + 'df1']['Low'] = locals()[i + 'df']['Low'].resample('W', label='right').min()
        locals()[i + 'df1']['weekvol'] = wvol[:len(locals()[i + 'df1'].index)]
        if i != 'vix':
            # 赋值周内交易量
            locals()[i + 'df1']['highvolw'] = locals()[i + 'df2'].iloc[:len(locals()[i + 'df1'].index.values),
                                              -2].values
            locals()[i + 'df1']['lowvolw'] = locals()[i + 'df2'].iloc[:len(locals()[i + 'df1'].index.values), -1].values
            # 根据日转前四天的最高价、最低价重定位当天交易量
            locals()[i + 'df1']['highvol'] = 0
            locals()[i + 'df1']['lowvol'] = 0
            for dnum in range(0, len(locals()[i + 'df1'].index), 1):
                aaa = locals()[i + 'df1'].index[dnum - 1] if dnum != 0 else pd.to_datetime(
                    locals()[i + 'df'].index.values[0])
                bbb = locals()[i + 'df1'].index[dnum]
                tlist = [datei for datei in locals()[i + 'df'].index.values if aaa <= pd.to_datetime(datei) < bbb]
                highd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :]['High'] == locals()[i + 'df1'].iloc[
                    dnum, 1]].index.values[0]
                locals()[i + 'df1'].iloc[dnum, -2] = locals()[i + 'df'].loc[highd, 'Volume']
                lowd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :]['Low'] == locals()[i + 'df1'].iloc[
                    dnum, 2]].index.values[0]
                locals()[i + 'df1'].iloc[dnum, -1] = locals()[i + 'df'].loc[lowd, 'Volume']
            locals()[i + 'df1'] = locals()[i + 'df1'].reindex(columns = ['Open','High','Low','Close','Adj Close'
                ,'Volume','highvol','lowvol','weekvol','highvolw','lowvolw'])
    else:
        locals()[i + 'df1'] = locals()[i + 'df'].resample('W', label='right').last()
        locals()[i + 'df1']['Open'] = locals()[i + 'df']['Open'].resample('W', label='right').first()
        locals()[i + 'df1']['Close'] = locals()[i + 'df']['Close'].resample('W', label='right').last()
        locals()[i + 'df1']['Adj Close'] = locals()[i + 'df']['Adj Close'].resample('W', label='right').last()
        locals()[i + 'df1']['High'] = locals()[i + 'df']['High'].resample('W', label='right').max()
        locals()[i + 'df1']['Low'] = locals()[i + 'df']['Low'].resample('W', label='right').min()
        locals()[i + 'df1']['highvolw'] = 0
        locals()[i + 'df1']['lowvolw'] = 0
        for wnum in range(0,len(locals()[i + 'df1'].index),1):
            aaa= locals()[i + 'df1'].index[wnum-1] if wnum != 0 else pd.to_datetime(locals()[i + 'df'].index.values[0])
            bbb = locals()[i + 'df1'].index[wnum]
            tlist = [datei for datei in locals()[i + 'df'].index.values if aaa<= pd.to_datetime(datei)<bbb]
            try:
                highd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :]['High'] == locals()[i + 'df1'].iloc[
                    wnum, 1]].index.values[0]
                locals()[i + 'df1'].iloc[wnum, -2] = locals()[i + 'df'].loc[highd, 'Volume']
            except:
                locals()[i + 'df1'].iloc[wnum, -2] = ""
            try:
                lowd = \
                locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :]['Low'] == locals()[i + 'df1'].iloc[
                    wnum, 2]].index.values[0]
                locals()[i + 'df1'].iloc[wnum,-1] = locals()[i + 'df'].loc[lowd,'Volume']
            except:
                locals()[i + 'df1'].iloc[wnum, -1] = ''
    locals()[i + 'df1'].to_excel(writer, sheet_name=i)
writer.save()

writer1 = pd.ExcelWriter(r'C:/Users/Tao/Desktop/投资/uvxy/0331/日转周数据.xlsx')
for i in ['vix','uvxy','dji','ixic','spx','hsi','sh']:
    locals()[i+'df'] = pd.read_excel(r'C:/Users/Tao/Desktop/投资/uvxy/0331/日数据.xlsx',sheet_name=[
        'vix','uvxy','dji','ixic','spx','hsi','sh'].index(i))
    locals()[i + 'df']['Date'] = locals()[i + 'df']['Date'].map(lambda y:pd.to_datetime(y) )
    locals()[i + 'df'].set_index('Date',inplace = True)
    locals()[i + 'df1'] = locals()[i + 'df'].resample('W', label='right').last()
    locals()[i + 'df1']['Open'] = locals()[i + 'df']['Open'].resample('W', label='right').first()
    locals()[i + 'df1']['High'] = locals()[i + 'df']['High'].resample('W', label='right').max()
    locals()[i + 'df1']['Low'] = locals()[i + 'df']['Low'].resample('W', label='right').min()
    locals()[i + 'df1'].to_excel(writer1, sheet_name=i)
writer1.save()


