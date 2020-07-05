import pandas as pd
import pandas_datareader.data as web
import datetime as dt
import numpy as np


#截止时间（当周下单日）
end = dt.datetime(2020,7,4)
#检查时把取数中A股隐藏避免取数错误

#主表
final_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/stoploss下单模板0630.xlsx',sheet_name = 1)

final_df.iloc[:,[2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]] = final_df.iloc[:,[
              2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]].fillna(0)

#上证
sh_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/SH.xls')
sh_df.时间 = sh_df.时间.apply(lambda y:dt.datetime.strptime(y[:-2],'%Y-%m-%d'))
sh_df.set_index('时间',inplace = True)

sh_df1 = sh_df.resample('W', label='right').last()
sh_df1.开盘 = sh_df.开盘 .resample('W', label='right').last().values
sh_df1.最高 = sh_df.最高.resample('W', label='right').max().values
sh_df1.最低 = sh_df.最低 .resample('W', label='right').min().values

#上证T周
try:
    final_df.iloc[2,13] = sh_df1[sh_df1.index == (end + dt.timedelta(days=1))].最高.values
    final_df.iloc[3,13] = sh_df1[sh_df1.index == (end + dt.timedelta(days=1))].最低.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-2] < datei < sh_df1.index[-1]], :]
except:
    final_df.iloc[2,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=6))].最高.values
    final_df.iloc[3,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=6))].最低.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-3] < datei < sh_df1.index[-1]], :]

final_df.iloc[5, 13] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values[0]
final_df.iloc[6, 13] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values[0]

#上证T-1周
try:
    final_df.iloc[4,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=6))].收盘.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-3] < datei < sh_df1.index[-2]], :]
except:
    final_df.iloc[4,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=13))].收盘.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-4] < datei < sh_df1.index[-2]], :]

final_df.iloc[0, 13] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]

#上证T-2周
try:
    '''final_df.iloc[0,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=13))].开盘.values
    final_df.iloc[6, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=13))].收盘.values'''
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-4] < datei < sh_df1.index[-3]], :]
except:
    '''final_df.iloc[0,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=20))].开盘.values
    final_df.iloc[6, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=20))].收盘.values'''
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-5] < datei < sh_df1.index[-3]], :]

final_df.iloc[1, 13] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]

#上证T-3周
'''try:
    final_df.iloc[1, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=20))].收盘.values
    #final_df.iloc[5, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=20))].最低.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-5] < datei < sh_df1.index[-4]], :]
except:
    final_df.iloc[1, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=27))].收盘.values
    #final_df.iloc[5, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=27))].最低.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-6] < datei < sh_df1.index[-4]], :]

final_df.iloc[8, 13] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values
final_df.iloc[10, 13] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values
final_df.iloc[11, 13] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values'''

#深证
sz_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/sz.xls')
sz_df.时间 = sz_df.时间.apply(lambda y:dt.datetime.strptime(y[:-2],'%Y-%m-%d'))
sz_df.set_index('时间',inplace = True)

sz_df1 = sz_df.resample('W', label='right').last()
sz_df1.开盘 = sz_df.开盘 .resample('W', label='right').last().values
sz_df1.最高 = sz_df.最高.resample('W', label='right').max().values
sz_df1.最低 = sz_df.最低 .resample('W', label='right').min().values

#深证T周
try:
    final_df.iloc[5,15] = sz_df1[sz_df1.index == (end + dt.timedelta(days=1))].最高.values
    final_df.iloc[6,15] = sz_df1[sz_df1.index == (end + dt.timedelta(days=1))].最低.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-2] < datei < sz_df1.index[-1]], :]
except:
    final_df.iloc[5,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=6))].最高.values
    final_df.iloc[6,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=6))].最低.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-3] < datei < sz_df1.index[-1]], :]

final_df.iloc[8, 15] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values[0]
final_df.iloc[9, 15] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values[0]

#深证T-1周
try:
    final_df.iloc[7,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=6))].收盘.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-3] < datei < sz_df1.index[-2]], :]
except:
    final_df.iloc[7,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=13))].收盘.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-4] < datei < sz_df1.index[-2]], :]

final_df.iloc[3, 15] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]

#深证T-2周
try:
    final_df.iloc[0,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=13))].最高.values
    final_df.iloc[1,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=13))].最低.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-4] < datei < sz_df1.index[-3]], :]
except:
    final_df.iloc[0,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=20))].最高.values
    final_df.iloc[1,15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=20))].最低.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-5] < datei < sz_df1.index[-3]], :]

final_df.iloc[4,15] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]
final_df.iloc[10,15] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values[0]
final_df.iloc[11,15] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values[0]

#深证T-3周
try:
    final_df.iloc[2, 15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=20))].收盘.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-5] < datei < sz_df1.index[-4]], :]
except:
    final_df.iloc[2, 15] = sz_df1[sz_df1.index == (end - dt.timedelta(days=27))].收盘.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-6] < datei < sz_df1.index[-4]], :]

final_df.iloc[12, 15] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]

#创业
cy_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/cy.xls')
cy_df.时间 = cy_df.时间.apply(lambda y:dt.datetime.strptime(y[:-2],'%Y-%m-%d'))
cy_df.set_index('时间',inplace = True)

cy_df1 = cy_df.resample('W', label='right').last()
cy_df1.开盘 = cy_df.开盘 .resample('W', label='right').last().values
cy_df1.最高 = cy_df.最高.resample('W', label='right').max().values
cy_df1.最低 = cy_df.最低 .resample('W', label='right').min().values

#创业T周
try:
    final_df.iloc[5,17] = cy_df1[cy_df1.index == (end + dt.timedelta(days=1))].最高.values
    final_df.iloc[6,17] = cy_df1[cy_df1.index == (end + dt.timedelta(days=1))].最低.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-2] < datei < cy_df1.index[-1]], :]
except:
    final_df.iloc[5,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=6))].最高.values
    final_df.iloc[6,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=6))].最低.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-3] < datei < cy_df1.index[-1]], :]

final_df.iloc[8, 17] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values[0]
final_df.iloc[9, 17] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values[0]

#创业T-1周
try:
    final_df.iloc[7,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=6))].收盘.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-3] < datei < cy_df1.index[-2]], :]
except:
    final_df.iloc[7,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=13))].收盘.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-4] < datei < cy_df1.index[-2]], :]

final_df.iloc[3, 17] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]

#创业T-2周
try:
    final_df.iloc[0,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=13))].最高.values
    final_df.iloc[1,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=13))].最低.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-4] < datei < cy_df1.index[-3]], :]
except:
    final_df.iloc[0,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=20))].最高.values
    final_df.iloc[1,17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=20))].最低.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-5] < datei < cy_df1.index[-3]], :]

final_df.iloc[4, 17] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]
final_df.iloc[10,17] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values[0]
final_df.iloc[11,17] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values[0]

#创业T-3周
try:
    final_df.iloc[2, 17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=20))].收盘.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-5] < datei < cy_df1.index[-4]], :]
except:
    final_df.iloc[2, 17] = cy_df1[cy_df1.index == (end - dt.timedelta(days=27))].收盘.values
    range_df = cy_df.loc[[datei for datei in cy_df.index if cy_df1.index[-6] < datei < cy_df1.index[-4]], :]

final_df.iloc[12, 17] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values[-1]

#T周
'''spy = web.DataReader('SPY', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,1] =spy.High.resample('W', label='right').max().values
final_df.iloc[1,1] =spy.Low.resample('W', label='right').min().values'''

vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,3] =vix.Open.resample('W', label='right').last().values

'''dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,5] =dji.Close.resample('W', label='right').last().values
#final_df.iloc[3,5] =dji.High.resample('W', label='right').max().values
final_df.iloc[3,5] =dji.Volume.resample('W', label='right').last().values'''

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,7] =ixic.High.resample('W', label='right').max().values
final_df.iloc[1,7] =ixic.Low.resample('W', label='right').min().values
high_value = ixic.High.resample('W', label='right').max().values[0]
final_df.iloc[5,7] =ixic[ixic.High == high_value].Volume.values[0]
low_value = ixic.Low.resample('W', label='right').min().values[0]
final_df.iloc[6,7] =ixic[ixic.Low == low_value].Volume.values[0]

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,9] =spx.Open.resample('W', label='right').last().values
final_df.iloc[5,9] =spx.Volume.resample('W', label='right').last().values
final_df.iloc[9,9] =spx.High.resample('W', label='right').max().values
final_df.iloc[10,9] =spx.Low.resample('W', label='right').min().values
high_value = spx.High.resample('W', label='right').max().values[0]
final_df.iloc[11,9] =spx[spx.High == high_value].Volume.values[0]
low_value = spx.Low.resample('W', label='right').min().values[0]
final_df.iloc[12,9] =spx[spx.Low == low_value].Volume.values[0]

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[3,11] =hsi.High.resample('W', label='right').max().values
final_df.iloc[4,11] =hsi.Low.resample('W', label='right').min().values
high_value = hsi.High.resample('W', label='right').max().values[0]
final_df.iloc[6,11] =hsi[hsi.High == high_value].Volume.values[0]
low_value = hsi.Low.resample('W', label='right').min().values[0]
final_df.iloc[7,11] =hsi[hsi.Low == low_value].Volume.values[0]

#T-1周
spy = web.DataReader('spy', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[0,1] =spy.High.resample('W', label='right').max().values
final_df.iloc[1,1] =spy.Low.resample('W', label='right').min().values

vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[1,3] =vix.Close.resample('W', label='right').last().values

'''dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[1,5] =dji.High.resample('W', label='right').max().values
#final_df.iloc[5,5] =dji.Close.resample('W', label='right').last().values
high_value = dji.High.resample('W', label='right').max().values[0]
final_df.iloc[4,5] =dji[dji.High == high_value].Volume.values'''

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[2,7] =ixic.Close.resample('W', label='right').last().values
final_df.iloc[7,7] =ixic.Volume.resample('W', label='right').last().values
final_df.iloc[8,7] =ixic.High.resample('W', label='right').max().values
final_df.iloc[9,7] =ixic.Low.resample('W', label='right').min().values
high_value = ixic.High.resample('W', label='right').max().values[0]
final_df.iloc[11,7] =ixic[ixic.High == high_value].Volume.values[0]
low_value = ixic.Low.resample('W', label='right').min().values[0]
final_df.iloc[12,7] =ixic[ixic.Low == low_value].Volume.values[0]

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[1,9] =spx.Close.resample('W', label='right').last().values
final_df.iloc[6,9] =spx.Volume.resample('W', label='right').last().values
final_df.iloc[13,9] =spx.High.resample('W', label='right').max().values
final_df.iloc[14,9] =spx.Low.resample('W', label='right').min().values
high_value = spx.High.resample('W', label='right').max().values[0]
final_df.iloc[16,9] =spx[spx.High == high_value].Volume.values[0]
low_value = spx.Low.resample('W', label='right').min().values[0]
final_df.iloc[17,9] =spx[spx.Low == low_value].Volume.values[0]

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[5,11] =hsi.Close.resample('W', label='right').last().values
final_df.iloc[8,11] =hsi.Volume.resample('W', label='right').last().values

#T-2周
spy = web.DataReader('spy', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[2,1] =spy.Close.resample('W', label='right').last().values

'''vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=(timeend0+14)),end - dt.timedelta(days=(timeend0+8)))
final_df.iloc[2,3] =vix.Low.resample('W', label='right').min().values'''

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[0,5] =dji.Volume.resample('W', label='right').last().values

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[3,7] =ixic.Volume.resample('W', label='right').last().values
final_df.iloc[10,7] =ixic.Close.resample('W', label='right').last().values

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[2,9] =spx.High.resample('W', label='right').max().values
final_df.iloc[3,9] =spx.Low.resample('W', label='right').min().values
final_df.iloc[7,9] =spx.Volume.resample('W', label='right').last().values
final_df.iloc[15,9] =spx.Close.resample('W', label='right').last().values

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[0,11] =hsi.High.resample('W', label='right').max().values
final_df.iloc[1,11] =hsi.Low.resample('W', label='right').min().values

#T-3周
'''spy = web.DataReader('spy', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[2,1] =spy.Low.resample('W', label='right').min().values
final_df.iloc[3,1] =spy.Close.resample('W', label='right').last().values
low_value = spy.Low.resample('W', label='right').min().values[0]
final_df.iloc[7,1] =spy[spy.Low == low_value].Volume.values
#final_df.iloc[12,1] =spy.Volume.resample('W', label='right').last().values'''

'''vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[5,3] =vix.Close.resample('W', label='right').last().values
#final_df.iloc[5,3] =vix.Close.resample('W', label='right').last().values'''

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[1,5] =dji.Volume.resample('W', label='right').last().values

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[4,7] =ixic.Volume.resample('W', label='right').last().values

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[4,9] =spx.Close.resample('W', label='right').last().values
final_df.iloc[8,9] =spx.Volume.resample('W', label='right').last().values

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[2,11] =hsi.Close.resample('W', label='right').last().values


final_df = final_df.applymap(lambda y: np.nan if y == 0 else y )

final_df.to_excel(r'C:/Users/Tao/Desktop/投资/spy/stoploss取数.xlsx' , index=True, encoding='utf_8_sig')