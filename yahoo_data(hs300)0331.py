import pandas as pd
import pandas_datareader.data as web
import datetime as dt
import numpy as np

#截止时间（下单日前周六）
end = dt.datetime(2020,5,9)
#检查时把取数中A股隐藏避免取数错误

#主表
final_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/HS300/HS300下单模板0331.xlsx',sheet_name = 1)

final_df.iloc[:,[2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]] = final_df.iloc[:,[
              2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]].fillna(0)

#etf
etf_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/hs300/etf.xls')
etf_df.时间 = etf_df.时间.apply(lambda y:dt.datetime.strptime(y[:-2],'%Y-%m-%d'))
etf_df.set_index('时间',inplace = True)

etf_df1 = etf_df.resample('W', label='right').last()
etf_df1.开盘 = etf_df.开盘 .resample('W', label='right').first().values
etf_df1.最高 = etf_df.最高.resample('W', label='right').max().values
etf_df1.最低 = etf_df.最低 .resample('W', label='right').min().values

#etf T周
try:
    final_df.iloc[0,1] = etf_df1[etf_df1.index == (end + dt.timedelta(days=1))].收盘.values
except:
    final_df.iloc[0,1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=6))].收盘.values

#range_df = etf_df.loc[[datei for datei in etf_df.index if etf_df1.index[-2]<datei<etf_df1.index[-1]],:]
#final_df.iloc[5, 1] = range_df[range_df.收盘== range_df.收盘[-1]]['总手'].values

#etf T-1周
try:
    final_df.iloc[4,1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=6))].收盘.values
except:
    final_df.iloc[4,1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=13))].收盘.values

#range_df = etf_df.loc[[datei for datei in etf_df.index if etf_df1.index[-3]<datei<etf_df1.index[-2]],:]
#final_df.iloc[6, 1] = range_df[range_df.最低== min(range_df.最低)]['总手'].values

#etf T-2周
try:
    final_df.iloc[2,1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=13))].最低.values
except:
    final_df.iloc[2,1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=20))].最低.values

#range_df = etf_df.loc[[datei for datei in etf_df.index if etf_df1.index[-4]<datei<etf_df1.index[-3]],:]
#final_df.iloc[5, 1] = range_df[range_df.收盘== range_df.收盘[-1]]['总手'].values

#etf T-3周
try:
    final_df.iloc[1, 1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=20))].最高.values
    final_df.iloc[3, 1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=20))].最低.values
except:
    final_df.iloc[1, 1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=27))].最高.values
    final_df.iloc[3, 1] = etf_df1[etf_df1.index == (end - dt.timedelta(days=27))].最低.values
#range_df = etf_df.loc[[datei for datei in etf_df.index if etf_df1.index[-5]<datei<etf_df1.index[-4]],:]
#final_df.iloc[6, 1] = range_df[range_df.最高== max(range_df.最高)]['总手'].values

#上证
sh_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/hs300/sh.xls')
sh_df.时间 = sh_df.时间.apply(lambda y:dt.datetime.strptime(y[:-2],'%Y-%m-%d'))
sh_df.set_index('时间',inplace = True)

sh_df1 = sh_df.resample('W', label='right').last()
sh_df1.开盘 = sh_df.开盘 .resample('W', label='right').first().values
sh_df1.最高 = sh_df.最高.resample('W', label='right').max().values
sh_df1.最低 = sh_df.最低 .resample('W', label='right').min().values

#上证T周
try:
    final_df.iloc[0,3] = sh_df1[sh_df1.index == (end + dt.timedelta(days=1))].收盘.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-2] < datei < sh_df1.index[-1]], :]
except:
    final_df.iloc[0,3] = sh_df1[sh_df1.index == (end - dt.timedelta(days=6))].收盘.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-3] < datei < sh_df1.index[-1]], :]

final_df.iloc[4, 3] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values

#上证T-1周
try:
    final_df.iloc[1,3] = sh_df1[sh_df1.index == (end - dt.timedelta(days=6))].最低.values
except:
    final_df.iloc[1,3] = sh_df1[sh_df1.index == (end - dt.timedelta(days=13))].最低.values

#上证T-2周
try:
    final_df.iloc[2, 3] = sh_df1[sh_df1.index == (end - dt.timedelta(days=13))].最低.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-4] < datei < sh_df1.index[-3]], :]
except:
    final_df.iloc[2, 3] = sh_df1[sh_df1.index == (end - dt.timedelta(days=20))].最低.values
    range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-5] < datei < sh_df1.index[-3]], :]

final_df.iloc[5, 3] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values

#上证T-3周
try:
    final_df.iloc[3, 3] = sh_df1[sh_df1.index == (end - dt.timedelta(days=20))].最低.values
except:
    final_df.iloc[3, 3] = sh_df1[sh_df1.index == (end - dt.timedelta(days=27))].最低.values
#range_df = sh_df.loc[[datei for datei in sh_df.index if sh_df1.index[-5]<datei<sh_df1.index[-4]],:]
#final_df.iloc[6,3] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values

#深证
sz_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/hs300/sz.xls')
sz_df.时间 = sz_df.时间.apply(lambda y:dt.datetime.strptime(y[:-2],'%Y-%m-%d'))
sz_df.set_index('时间',inplace = True)

sz_df1 = sz_df.resample('W', label='right').last()
sz_df1.开盘 = sz_df.开盘 .resample('W', label='right').first().values
sz_df1.最高 = sz_df.最高.resample('W', label='right').max().values
sz_df1.最低 = sz_df.最低 .resample('W', label='right').min().values

#深证T周
try:
    #final_df.iloc[0,5] = sz_df1[sz_df1.index == (end + dt.timedelta(days=1))].收盘.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-2] < datei < sz_df1.index[-1]], :]
except:
    #final_df.iloc[0,5] = sz_df1[sz_df1.index == (end - dt.timedelta(days=6))].收盘.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-3] < datei < sz_df1.index[-1]], :]

final_df.iloc[1, 5] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values

#深证T-1周
try:
    #final_df.iloc[2,5] = sz_df1[sz_df1.index == (end - dt.timedelta(days=6))].最高.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-3] < datei < sz_df1.index[-2]], :]
except:
    #final_df.iloc[2,5] = sz_df1[sz_df1.index == (end - dt.timedelta(days=13))].最高.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-4] < datei < sz_df1.index[-2]], :]
final_df.iloc[2, 5] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values

#深证T-2周
try:
    #final_df.iloc[3, 3] = sz_df1[sz_df1.index == (end - dt.timedelta(days=13))].最低.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-4]<datei<sz_df1.index[-3]],:]
except:
    #final_df.iloc[3, 3] = sz_df1[sz_df1.index == (end - dt.timedelta(days=20))].最低.values
    range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-5]<datei<sz_df1.index[-3]],:]

final_df.iloc[3, 5] = range_df[range_df.最低== min(range_df.最低)]['总手(万)'].values
final_df.iloc[4, 5] = range_df[range_df.收盘== range_df.收盘[-1]]['总手(万)'].values

#深证T-3周
#try:
#    final_df.iloc[1, 5] = sz_df1[sz_df1.index == (end - dt.timedelta(days=20))].最低.values
#except:
#    final_df.iloc[1, 5] = sz_df1[sz_df1.index == (end - dt.timedelta(days=27))].最低.values
#range_df = sz_df.loc[[datei for datei in sz_df.index if sz_df1.index[-5]<datei<sz_df1.index[-4]],:]
#final_df.iloc[6,3] = range_df[range_df.最高== max(range_df.最高)]['总手(万)'].values

#T周
'''hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[1,7] =hsi.Volume.resample('W', label='right').last().values'''

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,9] =dji.Close.resample('W', label='right').last().values
final_df.iloc[2,9] =dji.Low.resample('W', label='right').min().values
final_df.iloc[6,9] =dji.Volume.resample('W', label='right').last().values

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,11] =ixic.Close.resample('W', label='right').last().values
final_df.iloc[3,11] =ixic.Volume.resample('W', label='right').last().values

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=6),end )
final_df.iloc[0,13] =spx.Close.resample('W', label='right').last().values
final_df.iloc[1,13] =spx.Low.resample('W', label='right').min().values
final_df.iloc[4,13] =spx.Volume.resample('W', label='right').last().values

#T-1周
'''hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
high_value = hsi.High.resample('W', label='right').max().values[0]
final_df.iloc[2,7] =hsi[hsi.High == high_value].Volume.values'''

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[3,9] =dji.Low.resample('W', label='right').min().values
'''high_value = dji.High.resample('W', label='right').max().values[0]
final_df.iloc[5,9] =dji[dji.High == high_value].Volume.values'''

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[1,11] =ixic.Low.resample('W', label='right').min().values
final_df.iloc[5,11] =ixic.Volume.resample('W', label='right').last().values

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=13),end - dt.timedelta(days=7))
final_df.iloc[2,13] =spx.Low.resample('W', label='right').min().values

#T-2周
'''hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
high_value = hsi.High.resample('W', label='right').max().values[0]
final_df.iloc[3,7] =hsi[hsi.High == high_value].Volume.values'''

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[1,9] =dji.High.resample('W', label='right').max().values
final_df.iloc[4,9] =dji.Low.resample('W', label='right').min().values
final_df.iloc[5,9] =dji.Close.resample('W', label='right').last().values
high_value = dji.High.resample('W', label='right').max().values[0]
final_df.iloc[7,9] =dji[dji.High == high_value].Volume.values
final_df.iloc[8,9] =dji.Volume.resample('W', label='right').last().values

'''ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[3,9] =ixic.Low.resample('W', label='right').min().values
high_value = ixic.High.resample('W', label='right').max().values[0]
final_df.iloc[7,9] =ixic[ixic.High == high_value].Volume.values
final_df.iloc[8,9] =ixic.Volume.resample('W', label='right').last().values'''

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=20),end - dt.timedelta(days=14))
final_df.iloc[5,13] =spx.Volume.resample('W', label='right').last().values

#T-3周
'''hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[4,7] =hsi.Volume.resample('W', label='right').last().values'''

'''dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
high_value = dji.High.resample('W', label='right').max().values[0]
final_df.iloc[7,9] =dji[dji.High == high_value].Volume.values
low_value = dji.Low.resample('W', label='right').min().values[0]
final_df.iloc[8,9] =dji[dji.Low == low_value].Volume.values
final_df.iloc[9,9] =dji.Volume.resample('W', label='right').last().values'''

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[2,11] =ixic.Close.resample('W', label='right').last().values
low_value = ixic.Low.resample('W', label='right').min().values[0]
final_df.iloc[4,11] =ixic[ixic.Low == low_value].Volume.values

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=27),end - dt.timedelta(days=21))
final_df.iloc[3,13] =spx.Close.resample('W', label='right').last().values
final_df.iloc[6,13] =spx.Volume.resample('W', label='right').last().values

final_df = final_df.applymap(lambda y: np.nan if y == 0 else y )

final_df.to_excel(r'C:/Users/Tao/Desktop/投资/HS300/取数.xlsx' , index=True, encoding='utf_8_sig')