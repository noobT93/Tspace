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

#基本数据汇总到一张sheet中
final_df = pd.DataFrame(index = pd.read_excel(r'C:/Users/Tao/Desktop/投资/final/基本数据.xlsx',sheet_name=0).Date,
                        columns = [])

for i in ['spy','^VIX','^DJI','^IXIC','^GSPC','^HSI','sh','sz','cy']:
    recent_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/final/基本数据.xlsx',sheet_name=[
        'spy','^VIX','^DJI','^IXIC','^GSPC','^HSI','sh','sz','cy'].index(i))
    final_df[i+'high'] = list(recent_df.High)
    final_df[i+'low'] = list(recent_df.Low)
    final_df[i+'close'] = list(recent_df.Close)
    final_df[i+'volume'] = list(recent_df.Volume)
    final_df[i+'highvol'] = list(recent_df.highvol)
    final_df[i+'lowvol'] = list(recent_df.lowvol)

#计算spy本周最大波动率,vix本周最大波动率
spymaxcoe = []
vixmaxcoe = []
for numi in range(1,len(final_df),1):
    indext = final_df.index[numi]
    indext_1 = final_df.index[numi-1]
    spymaxcoe.append(math.log(final_df.loc[indext,'spyhigh']/final_df.loc[
        indext_1,'spyclose']) if final_df.loc[indext,'spyhigh'] - final_df.loc[
        indext_1,'spyclose']>= final_df.loc[indext_1,'spyclose']- final_df.loc[
        indext,'spylow'] else math.log(final_df.loc[indext,'spylow']/final_df.loc[
        indext_1,'spyclose']))

    vixmaxcoe.append(math.log(final_df.loc[indext,'^VIXhigh']/final_df.loc[
        indext_1,'^VIXclose']) if final_df.loc[indext,'^VIXhigh'] - final_df.loc[
        indext_1,'^VIXclose']>= final_df.loc[indext_1,'^VIXclose']- final_df.loc[
        indext,'^VIXlow'] else math.log(final_df.loc[indext,'^VIXlow']/final_df.loc[
        indext_1,'^VIXclose']))

for vari in ['spy','^VIX','^DJI','^IXIC','^GSPC','^HSI','sh','sz','cy']:
    locals()[vari  + 'closep0'] = []
    locals()[vari  + 'maxp0'] = []
    if vari != '^VIX':
        locals()[vari  + 'closev0'] = []
        locals()[vari  + 'maxv0'] = []

    for numi in range(2, len(final_df), 1):
        indext = final_df.index[numi-1]
        indext_1 = final_df.index[numi - 2]
        locals()[vari  + 'closep0'].append(math.log(final_df.loc[indext,vari+'close']/final_df.loc[
        indext_1,vari+'close']))
        locals()[vari  + 'maxp0'].append(math.log(final_df.loc[indext,vari+'high']/final_df.loc[
        indext_1,vari+'close']) if final_df.loc[indext,vari+'high'] - final_df.loc[
        indext_1,vari+'close']>= final_df.loc[indext_1,vari+'close']- final_df.loc[
        indext,vari+'low'] else math.log(final_df.loc[indext,vari+'low']/final_df.loc[
        indext_1,vari+'close']))
        if vari != '^VIX':
            locals()[vari + 'closev0'].append(math.log(final_df.loc[indext,vari+'volume']/final_df.loc[
        indext_1,vari+'volume']))
            locals()[vari + 'maxv0'].append(math.log(final_df.loc[indext,vari+'highvol']/final_df.loc[
        indext_1,vari+'volume']) if final_df.loc[indext,vari+'high'] - final_df.loc[
        indext_1,vari+'close']>= final_df.loc[indext_1,vari+'close']- final_df.loc[
        indext,vari+'low'] else math.log(final_df.loc[indext,vari+'lowvol']/final_df.loc[
        indext_1,vari+'volume']))

    locals()[vari + 'closep1'] = locals()[vari  + 'closep0'][:-1]
    locals()[vari + 'closep2'] = locals()[vari  + 'closep0'][:-2]
    locals()[vari + 'maxp1'] = locals()[vari  + 'maxp0'][:-1]
    locals()[vari + 'maxp2'] = locals()[vari  + 'maxp0'][:-2]
    if vari != '^VIX':
        locals()[vari + 'closev1'] = locals()[vari + 'closev0'][:-1]
        locals()[vari + 'closev2'] = locals()[vari + 'closev0'][:-2]
        locals()[vari + 'maxv1'] = locals()[vari + 'maxv0'][:-1]
        locals()[vari + 'maxv2'] = locals()[vari + 'maxv0'][:-2]

spylast_df = pd.DataFrame(index = final_df.index,columns = ['spymaxcoe'])
spylast_df.iloc[1:,0] = spymaxcoe
for spyi in ['spy','^VIX','^DJI','^IXIC','^GSPC','^HSI','sh','sz','cy']:
    for spya in ['closep','maxp']:
        for spyb in ['0','1','2']:
            spylast_df[spyi + spya + spyb] = 0
            if spyb == '0':
                spylast_df.iloc[2:,-1] = locals()[spyi + spya + spyb]
            elif spyb == '1':
                spylast_df.iloc[3:, -1] = locals()[spyi + spya + spyb]
            else:
                spylast_df.iloc[4:, -1] = locals()[spyi + spya + spyb]

    if spyi != '^VIX':
        for spyc in ['closev', 'maxv']:
            for spyd in ['0', '1', '2']:
                spylast_df[spyi + spyc + spyd] = 0
                if spyd == '0':
                    spylast_df.iloc[2:, -1] = locals()[spyi + spyc + spyd]
                elif spyd == '1':
                    spylast_df.iloc[3:, -1] = locals()[spyi + spyc + spyd]
                else:
                    spylast_df.iloc[4:, -1] = locals()[spyi + spyc + spyd]
spylast_df.to_excel(r'C:/Users/Tao/Desktop/投资/final/spy基本数据.xlsx')


vixlast_df = pd.DataFrame(index = final_df.index,columns = ['^VIXmaxcoe'])
vixlast_df.iloc[1:,0] = vixmaxcoe
for vixi in ['^VIX','^DJI','^IXIC','^GSPC','^HSI','sh','sz','cy']:
    for vixa in ['closep','maxp']:
        for vixb in ['0','1','2']:
            vixlast_df[vixi + vixa + vixb] = 0
            if vixb == '0':
                vixlast_df.iloc[2:,-1] = locals()[vixi + vixa + vixb]
            elif vixb == '1':
                vixlast_df.iloc[3:, -1] = locals()[vixi + vixa + vixb]
            else:
                vixlast_df.iloc[4:, -1] = locals()[vixi + vixa + vixb]

    if vixi != '^VIX':
        for vixc in ['closev', 'maxv']:
            for vixd in ['0', '1', '2']:
                vixlast_df[vixi + vixc + vixd] = 0
                if vixd == '0':
                    vixlast_df.iloc[2:, -1] = locals()[vixi + vixc + vixd]
                elif vixd == '1':
                    vixlast_df.iloc[3:, -1] = locals()[vixi + vixc + vixd]
                else:
                    vixlast_df.iloc[4:, -1] = locals()[vixi + vixc + vixd]
vixlast_df.to_excel(r'C:/Users/Tao/Desktop/投资/final/vix基本数据.xlsx')