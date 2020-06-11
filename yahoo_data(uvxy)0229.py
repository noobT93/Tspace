import pandas as pd
import pandas_datareader.data as web
import datetime as dt
import numpy as np


#截止时间（当周下单日）
end = dt.datetime(2020,3,27)

#主表
final_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/uvxy/下单模板(p_vol)0229.xlsx',sheet_name = 1)

final_df.iloc[:,[2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]] = final_df.iloc[:,[
              2*i+ 1 for i in range(0,len(final_df.columns)//2,1)]].fillna(0)

#T周
vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=5),end - dt.timedelta(days=1))
final_df.iloc[0,1] =vix['Close'].resample('W', label='right').last().values
final_df.iloc[1,1] =vix['High'].resample('W', label='right').max().values

uvxy = web.DataReader('UVXY', 'yahoo',end - dt.timedelta(days=5),end - dt.timedelta(days=1))
final_df.iloc[0,3] =uvxy['Close'].resample('W', label='right').last().values
final_df.iloc[1,3] =uvxy['High'].resample('W', label='right').max().values
final_df.iloc[2,3] =uvxy['Low'].resample('W', label='right').min().values
final_df.iloc[5,3] =uvxy['Volume'].resample('W', label='right').last().values

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=5),end - dt.timedelta(days=1))
final_df.iloc[0,5] =dji['Close'].resample('W', label='right').last().values
final_df.iloc[4,5] =dji['Volume'].resample('W', label='right').last().values

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=5),end - dt.timedelta(days=1))
final_df.iloc[0,7] =ixic['Close'].resample('W', label='right').last().values
final_df.iloc[2,7] =ixic['Volume'].resample('W', label='right').last().values

spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=5),end - dt.timedelta(days=1))
final_df.iloc[0,9] =spx['Close'].resample('W', label='right').last().values
final_df.iloc[1,9] =spx['High'].resample('W', label='right').max().values

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=5),end)
final_df.iloc[0,11] =hsi['Close'].resample('W', label='right').last().values
final_df.iloc[2,11] =hsi['Low'].resample('W', label='right').min().values
final_df.iloc[5,11] =hsi['Volume'].resample('W', label='right').last().values

#T-1周
vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=12),end - dt.timedelta(days=6))
final_df.iloc[2,1] =vix['Close'].resample('W', label='right').last().values

uvxy = web.DataReader('UVXY', 'yahoo',end - dt.timedelta(days=12),end - dt.timedelta(days=6))
final_df.iloc[3,3] =uvxy['Low'].resample('W', label='right').min().values
final_df.iloc[4,3] =uvxy['Close'].resample('W', label='right').last().values
high_value = uvxy['High'].resample('W', label='right').max().values[0]
final_df.iloc[6,3] =uvxy[uvxy.High == high_value].Volume.values
final_df.iloc[7,3] =uvxy['Volume'].resample('W', label='right').last().values

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=12),end - dt.timedelta(days=6))
final_df.iloc[1,5] =dji['High'].resample('W', label='right').max().values
low_value = dji['Low'].resample('W', label='right').min().values[0]
final_df.iloc[5,5] =dji[dji.Low == low_value].Volume.values

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=12),end - dt.timedelta(days=6))
low_value = ixic['Low'].resample('W', label='right').min().values[0]
final_df.iloc[3,7] =ixic[ixic.Low == low_value].Volume.values

'''
spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=12),end - dt.timedelta(days=6))
final_df.iloc[2,9] =spx['Low'].resample('W', label='right').min().values
final_df.iloc[4,9] =spx['Close'].resample('W', label='right').last().values
'''

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=12),end - dt.timedelta(days=6))
final_df.iloc[1,11] =hsi['High'].resample('W', label='right').max().values
high_value = hsi['High'].resample('W', label='right').max().values[0]
final_df.iloc[7,11] =hsi[hsi.High == high_value].Volume.values

#T-2周
'''
vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=19),end - dt.timedelta(days=13))
final_df.iloc[2,1] =vix['Low'].resample('W', label='right').min().values
'''

uvxy = web.DataReader('UVXY', 'yahoo',end - dt.timedelta(days=19),end - dt.timedelta(days=13))
final_df.iloc[8,3] =uvxy['Volume'].resample('W', label='right').last().values

dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=19),end - dt.timedelta(days=13))
final_df.iloc[2,5] =dji['Low'].resample('W', label='right').min().values
final_df.iloc[3,5] =dji['Close'].resample('W', label='right').last().values

ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=19),end - dt.timedelta(days=13))
final_df.iloc[1,7] =ixic['Low'].resample('W', label='right').min().values
low_value = ixic['Low'].resample('W', label='right').min().values[0]
final_df.iloc[4,7] =ixic[ixic.Low == low_value].Volume.values
final_df.iloc[5,7] =ixic['Volume'].resample('W', label='right').last().values

'''
spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=19),end - dt.timedelta(days=13))
final_df.iloc[5,9] =spx['Close'].resample('W', label='right').last().values
final_df.iloc[8,9] =spx['Volume'].resample('W', label='right').last().values
'''

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=19),end - dt.timedelta(days=13))
final_df.iloc[4,11] =hsi['Close'].resample('W', label='right').last().values
low_value = hsi['Low'].resample('W', label='right').min().values[0]
final_df.iloc[6,11] =hsi[hsi.Low == low_value].Volume.values
final_df.iloc[7,11] =hsi['Volume'].resample('W', label='right').last().values

#T-3周
'''
vix= web.DataReader('^VIX', 'yahoo',end - dt.timedelta(days=26),end - dt.timedelta(days=20))
final_df.iloc[4,1] =vix['Close'].resample('W', label='right').last().values
'''
'''
uvxy = web.DataReader('UVXY', 'yahoo',end - dt.timedelta(days=26),end - dt.timedelta(days=20))
final_df.iloc[8,3] =uvxy['Volume'].resample('W', label='right').last().values
'''
'''
dji= web.DataReader('^DJI', 'yahoo',end - dt.timedelta(days=26),end - dt.timedelta(days=20))
high_value = dji['High'].resample('W', label='right').max().values[0]
final_df.iloc[7,5] =dji[dji.High == high_value].Volume.values
'''
'''
ixic= web.DataReader('^IXIC', 'yahoo',end - dt.timedelta(days=26),end - dt.timedelta(days=20))
high_value = ixic['High'].resample('W', label='right').max().values[0]
final_df.iloc[4,7] =ixic[ixic.High == high_value].Volume.values
low_value = ixic['Low'].resample('W', label='right').min().values[0]
final_df.iloc[5,7] =ixic[ixic.Low == low_value].Volume.values
'''
spx= web.DataReader('^GSPC', 'yahoo',end - dt.timedelta(days=26),end - dt.timedelta(days=20))
final_df.iloc[2,9] =spx['High'].resample('W', label='right').max().values
final_df.iloc[3,9] =spx['Low'].resample('W', label='right').min().values

hsi= web.DataReader('^HSI', 'yahoo',end - dt.timedelta(days=26),end - dt.timedelta(days=20))
final_df.iloc[3,11] =hsi['Low'].resample('W', label='right').min().values

final_df = final_df.applymap(lambda y: np.nan if y == 0 else y )

final_df.to_excel(r'C:/Users/Tao/Desktop/投资/uvxy/取数.xlsx' , index=True, encoding='utf_8_sig')