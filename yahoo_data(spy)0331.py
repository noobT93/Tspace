import pandas as pd
import pandas_datareader.data as web
import datetime as dt
import numpy as np


#截止时间（当周下单日）
end = dt.datetime(2020,5,15)
#检查时把取数中A股隐藏避免取数错误

#主表
final_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/下单模板0331.xlsx',sheet_name = 1)

final_df.iloc[:,[2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]] = final_df.iloc[:,[
              2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]].fillna(0)

#上证
sh_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/SH.xls')
sh_df.时间 = sh_df.时间.apply(lambda y:dt.datetime.strptime(y[:-2],'%Y-%m-%d'))
sh_df.set_index('时间',inplace = True)

sh_df1 = sh_df.resample('W', label='right').last()
sh_df1.开盘 = sh_df.开盘 .resample('W', label='right').first().values
sh_df1.最高 = sh_df.最高.resample('W', label='right').max().values
sh_df1.最低 = sh_df.最低 .resample('W', label='right').min().values

#上证T周
try:
    final_df.iloc[0,13] = sh_df1[sh_df1.index == (end + dt.timedelta(days=2))].收盘.values
    final_df.iloc[1,13] = sh_df1[sh_df1.index == (end + dt.timedelta(days=2))].最高.values
    #range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-2] < datei < sh_df1.index[-1]], :]
except:
    final_df.iloc[0,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=5))].收盘.values
    final_df.iloc[1,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=5))].最高.values
    #range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-3] < datei < sh_df1.index[-1]], :]

#final_df.iloc[5, 13] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values

#上证T-1周
#try:
#    final_df.iloc[1,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=5))].最高.values
#except:
#    final_df.iloc[1,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=12))].最高.values

#上证T-2周
try:
    final_df.iloc[3,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=12))].最低.values
    final_df.iloc[4, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=12))].收盘.values
except:
    final_df.iloc[4,13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=19))].最低.values
    final_df.iloc[4, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=19))].收盘.values

#上证T-3周
try:
    final_df.iloc[2, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=19))].最高.values
    #range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-5] < datei < sh_df1.index[-4]], :]
except:
    final_df.iloc[2, 13] = sh_df1[sh_df1.index == (end - dt.timedelta(days=26))].最高.values
    #range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-6] < datei < sh_df1.index[-4]], :]

#final_df.iloc[6, 13] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values

if end.weekday() == 4:
    timeend0 = 5
if end.weekday() == 3:
    timeend0 = 4

#T周
spy = web.DataReader('SPY', 'yahoo',end - dt.timedelta(days=timeend0),end - dt.timedelta(days=1))
final_df.iloc[0,1] =spy.Close.resample('W', label='right').last().values

vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=timeend0),end - dt.timedelta(days=1))
final_df.iloc[0,3] =vix.Close.resample('W', label='right').last().values
#final_df.iloc[1,1] =vix.High.resample('W', label='right').max().values

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=timeend0),end - dt.timedelta(days=1))
final_df.iloc[0,5] =dji.Close.resample('W', label='right').last().values
#final_df.iloc[1,5] =dji.High.resample('W', label='right').max().values
#final_df.iloc[2,5] =dji.Volume.resample('W', label='right').last().values

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=timeend0),end - dt.timedelta(days=1))
final_df.iloc[0,7] =ixic.Close.resample('W', label='right').last().values
#final_df.iloc[1,7] =ixic.Volume.resample('W', label='right').last().values


spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=timeend0),end - dt.timedelta(days=1))
final_df.iloc[0,9] =spx.Close.resample('W', label='right').last().values
#final_df.iloc[1,9] =spx.High.resample('W', label='right').max().values
#final_df.iloc[4,9] =spx.Low.resample('W', label='right').min().values
final_df.iloc[2,9] =spx.Volume.resample('W', label='right').last().values

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=timeend0),end)
final_df.iloc[0,11] =hsi.Close.resample('W', label='right').last().values
final_df.iloc[1,11] =hsi.Low.resample('W', label='right').min().values
#final_df.iloc[4,11] =hsi.Volume.resample('W', label='right').last().values

#T-1周
spy = web.DataReader('spy', 'yahoo',end - dt.timedelta(days=(timeend0+7)),end - dt.timedelta(days=(timeend0+1)))
final_df.iloc[1,1] =spy.High.resample('W', label='right').max().values
final_df.iloc[4,1] =spy.Close.resample('W', label='right').last().values
'''high_value = spy.High.resample('W', label='right').max().values[0]
final_df.iloc[3,1] =spy[spy.High == high_value].Volume.values'''

'''vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=(timeend0+7)),end - dt.timedelta(days=(timeend0+1))))
final_df.iloc[1,3] =vix.High.resample('W', label='right').max().values
final_df.iloc[3,3] =vix.Low.resample('W', label='right').min().values'''

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=(timeend0+7)),end - dt.timedelta(days=(timeend0+1)))
'''final_df.iloc[3,5] =dji.Low.resample('W', label='right').min().values
final_df.iloc[5,5] =dji.Close.resample('W', label='right').last().values
low_value = dji.Low.resample('W', label='right').min().values[0]
final_df.iloc[3,5] =dji[dji.Low == low_value].Volume.values'''

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=(timeend0+7)),end - dt.timedelta(days=(timeend0+1)))
final_df.iloc[2,7] =ixic.Close.resample('W', label='right').last().values
'''high_value = ixic.High.resample('W', label='right').max().values[0]
final_df.iloc[6,7] =ixic[ixic.High == high_value].Volume.values
low_value = ixic.Low.resample('W', label='right').min().values[0]
final_df.iloc[2,7] =ixic[ixic.Low == low_value].Volume.values'''
#final_df.iloc[11,7] =ixic.Volume.resample('W', label='right').last().values

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=(timeend0+7)),end - dt.timedelta(days=(timeend0+1)))
final_df.iloc[1,9] =spx.High.resample('W', label='right').max().values
'''high_value = spx.High.resample('W', label='right').max().values[0]
final_df.iloc[8,9] =spx[spx.High == high_value].Volume.values
low_value = spx.Low.resample('W', label='right').min().values[0]
final_df.iloc[5,9] =spx[spx.Low == low_value].Volume.values'''

'''hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=(timeend0+7)),end - dt.timedelta(days=(timeend0+1)))
final_df.iloc[1,11] =hsi.High.resample('W', label='right').max().values
final_df.iloc[6,11] =hsi.Volume.resample('W', label='right').last().values
high_value = hsi.High.resample('W', label='right').max().values[0]
final_df.iloc[7,11] =hsi[hsi.High == high_value].Volume.values'''

#T-2周
'''spy = web.DataReader('spy', 'yahoo',end - dt.timedelta(days=(timeend0+14)),end - dt.timedelta(days=(timeend0+8)))
final_df.iloc[3,1] =spy.High.resample('W', label='right').max().values
final_df.iloc[11,1] =spy.Volume.resample('W', label='right').last().values'''

vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=(timeend0+14)),end - dt.timedelta(days=(timeend0+8)))
final_df.iloc[2,3] =vix.Low.resample('W', label='right').min().values

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=(timeend0+14)),end - dt.timedelta(days=(timeend0+8)))
final_df.iloc[1,5] =dji.High.resample('W', label='right').max().values
'''final_df.iloc[4,5] =dji.Low.resample('W', label='right').min().values
high_value = dji.High.resample('W', label='right').max().values[0]
final_df.iloc[8,5] =dji[dji.High == high_value].Volume.values
low_value = dji.Low.resample('W', label='right').min().values[0]
final_df.iloc[11,5] =dji[dji.Low == low_value].Volume.values'''

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=(timeend0+14)),end - dt.timedelta(days=(timeend0+8)))
final_df.iloc[1,7] =ixic.Low.resample('W', label='right').min().values
'''final_df.iloc[4,7] =ixic.Close.resample('W', label='right').last().values
high_value = ixic.High.resample('W', label='right').max().values[0]
final_df.iloc[7,7] =ixic[ixic.High == high_value].Volume.values
low_value = ixic.Low.resample('W', label='right').min().values[0]
final_df.iloc[10,7] =ixic[ixic.Low == low_value].Volume.values
final_df.iloc[12,7] =ixic.Volume.resample('W', label='right').last().values'''

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=(timeend0+14)),end - dt.timedelta(days=(timeend0+8)))
'''final_df.iloc[2,9] =spx.High.resample('W', label='right').max().values
final_df.iloc[5,9] =spx.Close.resample('W', label='right').last().values'''
high_value = spx.High.resample('W', label='right').max().values[0]
final_df.iloc[3,9] =spx[spx.High == high_value].Volume.values
#final_df.iloc[13,9] =spx.Volume.resample('W', label='right').last().values

'''hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=(timeend0+14)),end - dt.timedelta(days=(timeend0+8)))
final_df.iloc[4,11] =hsi.Low.resample('W', label='right').min().values
high_value = hsi.High.resample('W', label='right').max().values[0]
final_df.iloc[8,11] =hsi[hsi.High == high_value].Volume.values
final_df.iloc[7,11] =hsi.Volume.resample('W', label='right').last().values'''

#T-3周
spy = web.DataReader('spy', 'yahoo',end - dt.timedelta(days=26),end - dt.timedelta(days=20))
final_df.iloc[2,1] =spy.High.resample('W', label='right').max().values
final_df.iloc[3,1] =spy.Low.resample('W', label='right').min().values
#final_df.iloc[12,1] =spy.Volume.resample('W', label='right').last().values

vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=(timeend0+21)),end - dt.timedelta(days=(timeend0+15)))
final_df.iloc[1,3] =vix.High.resample('W', label='right').max().values
#final_df.iloc[5,3] =vix.Close.resample('W', label='right').last().values

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=(timeend0+21)),end - dt.timedelta(days=(timeend0+15)))
final_df.iloc[2,5] =dji.Close.resample('W', label='right').last().values
'''high_value = dji.High.resample('W', label='right').max().values[0]
final_df.iloc[9,5] =dji[dji.High == high_value].Volume.values
final_df.iloc[4,5] =dji.Volume.resample('W', label='right').last().values'''

'''ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=(timeend0+21)),end - dt.timedelta(days=(timeend0+15)))
high_value = ixic.High.resample('W', label='right').max().values[0]
final_df.iloc[8,7] =ixic[ixic.High == high_value].Volume.values
final_df.iloc[3,7] =ixic.Volume.resample('W', label='right').last().values'''

'''spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=(timeend0+21)),end - dt.timedelta(days=(timeend0+15)))
final_df.iloc[6,9] =spx.Close.resample('W', label='right').last().values
high_value = spx.High.resample('W', label='right').max().values[0]
final_df.iloc[10,9] =spx[spx.High == high_value].Volume.values
low_value = spx.Low.resample('W', label='right').min().values[0]
final_df.iloc[6,9] =spx[spx.Low == low_value].Volume.values
final_df.iloc[7,9] =spx.Volume.resample('W', label='right').last().values'''

'''hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=(timeend0+21)),end - dt.timedelta(days=(timeend0+15)))
final_df.iloc[2,11] =hsi.High.resample('W', label='right').max().values
final_df.iloc[5,11] =hsi.Low.resample('W', label='right').min().values
high_value = hsi.High.resample('W', label='right').max().values[0]
final_df.iloc[9,11] =hsi[hsi.High == high_value].Volume.values
low_value = hsi.Low.resample('W', label='right').min().values[0]
final_df.iloc[5,11] =hsi[hsi.Low == low_value].Volume.values
#final_df.iloc[12,11] =hsi.Volume.resample('W', label='right').last().values'''

final_df = final_df.applymap(lambda y: np.nan if y == 0 else y )

final_df.to_excel(r'C:/Users/Tao/Desktop/投资/spy/取数.xlsx' , index=True, encoding='utf_8_sig')