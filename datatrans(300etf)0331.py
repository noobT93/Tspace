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

writer = pd.ExcelWriter(r'C:/Users/Tao/Desktop/投资/HS300/0331/日转周数据.xlsx')
for i in ['etf','sh','sz','cy','hsi','dji','ixic','spx']:
    locals()[i+'df'] = pd.read_excel(r'C:/Users/Tao/Desktop/投资/HS300/0331/日数据.xlsx',sheet_name=[
        'etf','sh','sz','cy','hsi','dji','ixic','spx'].index(i))
    locals()[i + 'df']['Date'] = locals()[i + 'df']['Date'].map(lambda y:pd.to_datetime(y) )
    locals()[i + 'df'].set_index('Date',inplace = True)

    locals()[i + 'df1'] = locals()[i + 'df'].resample('W', label='right').last()
    locals()[i + 'df1']['Open'] = locals()[i + 'df']['Open'].resample('W', label='right').last()
    locals()[i + 'df1']['High'] = locals()[i + 'df']['High'].resample('W', label='right').max()
    locals()[i + 'df1']['Low'] = locals()[i + 'df']['Low'].resample('W', label='right').min()
    locals()[i + 'df1']['highvol'] = 0
    locals()[i + 'df1']['lowvol'] = 0
    for wnum in range(0, len(locals()[i + 'df1'].index), 1):
        aaa = locals()[i + 'df1'].index[wnum - 1] if wnum != 0 else pd.to_datetime(locals()[i + 'df'].index.values[0])
        bbb = locals()[i + 'df1'].index[wnum]
        tlist = [datei for datei in locals()[i + 'df'].index.values if aaa <= pd.to_datetime(datei) < bbb]
        try:
            highd = locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :]['High'] == locals()[i + 'df1'].iloc[
            wnum, 1]].index.values[0]
            locals()[i + 'df1'].iloc[wnum, -2] = locals()[i + 'df'].loc[highd, 'Volume']
        except:
            locals()[i + 'df1'].iloc[wnum, -2] = np.nan
        try:
            lowd = locals()[i + 'df'].loc[tlist, :][locals()[i + 'df'].loc[tlist, :]['Low'] == locals()[i + 'df1'].iloc[
            wnum, 2]].index.values[0]
            locals()[i + 'df1'].iloc[wnum, -1] = locals()[i + 'df'].loc[lowd, 'Volume']
        except:
            locals()[i + 'df1'].iloc[wnum, -1] =np.nan
    locals()[i + 'df1'].to_excel(writer, sheet_name=i)
writer.save()



