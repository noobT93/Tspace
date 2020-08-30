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
import calendar

#后向最小二乘法
def backward_regression(X, y,threshold_out,verbose=False):
    included=list(X.columns)
    X = sm.add_constant(pd.DataFrame(X[included]))
    included = list(X.columns)
    while True:
        changed=False
        model = sm.OLS(y, pd.DataFrame(X[included])).fit()
        # use all coefs except intercept
        pvalues = model.pvalues.iloc[:]

        worst_pval = pvalues.max() # null if pvalues is empty
        if worst_pval > threshold_out:
            changed=True
            worst_feature = pvalues.idxmax()
            included.remove(worst_feature)
            if verbose:
                print('Drop {:30} with p-value {:.6}'.format(worst_feature, worst_pval))
        if not changed:
            break
    return [model,included]

vixinit_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/final/vix基本数据.xlsx',sheet_name=0)
vixinit_df.set_index('Date', inplace=True)

#先靠倒数第二个数据定位
#测试集为2年，倒数第二往前推2年的下月1日，不确定最后一周是否为下一月，得到测试集起点

teststart_date = dt.datetime(vixinit_df.index[-2].year-2,vixinit_df.index[-2].month+1,1,0,0,0)
teststart_point = [i for i in vixinit_df.index if i >= teststart_date][0]

#获取index 年月数据
date_init = [str(i.year) + str(i.month).zfill(2) for i in vixinit_df.index]
date_init = sorted(set(date_init),key = date_init.index)
#测试起点至2年后终点，不可直接-1，不确定最后一周
date_init= date_init[date_init.index(str(teststart_point.year) + str(teststart_point.month).zfill(
    2)):1+date_init.index(str(teststart_point.year+2) + str(teststart_point.month).zfill(2))]

traintest_df = pd.DataFrame(index = [date_init[0] +'-'+date_init[-1]],columns = ['width3freq1','width3freq3'])

#训练集样本为3年
#每月更新
data_width = 3
for data_freq in [1,3]:
    #vix不需要调整
    locals()['width3freq' + str(data_freq)] = pd.DataFrame(index = vixinit_df.index[
        list(vixinit_df.index).index(teststart_point):-1],columns = ['vixmaxcoe','foreresult',
        'realjudge'])
    locals()['width3freq' + str(data_freq)].vixmaxcoe = [math.exp(i)for i in vixinit_df[
        teststart_point<=vixinit_df.index][:-1]['^VIXmaxcoe']]
    locals()['width3freq' + str(data_freq)].realjudge = 0

    if data_freq ==3:
        date_init = [date_init[i*3] for i in range(0,9,1)]
    for datenum in date_init:
        if data_freq ==1:
            if int(datenum[4:]) == 12:
                endyear = int(datenum[:4]) + 1
                endmonth = 1
            else:
                endyear = int(datenum[:4])
                endmonth = int(datenum[4:]) + data_freq
        else:
            if int(datenum[4:]) == 10:
                endyear = int(datenum[:4]) + 1
                endmonth = 1
            else:
                endyear = int(datenum[:4])
                endmonth = int(datenum[4:]) + data_freq

        #生成训练集
        traindata_df = vixinit_df[dt.datetime(
            int(datenum[:4])-data_width,int(datenum[4:]),1,0,0,0) <=vixinit_df.index]
        traindata_df = traindata_df[traindata_df.index<dt.datetime(
            int(datenum[:4]),int(datenum[4:]),1,0,0,0)][3:-1]
        y_value = traindata_df.iloc[:, 0].values
        data = traindata_df.iloc[:, 1:]
        #测试集宽度
        testdata_df = vixinit_df[dt.datetime(
            int(datenum[:4]),int(datenum[4:]),1,0,0,0) <=vixinit_df.index][1:]
        testdata_df = testdata_df[testdata_df.index <dt.datetime(
            endyear,endmonth,1,0,0,0) ]
        #获取自变量
        if 'const' not in backward_regression(data, y_value, 0.05, verbose=False)[-1]:
            testvariable = backward_regression(data, y_value, 0.05, verbose=False)[-1]
            testresult = [math.exp(i) for i in backward_regression(
                data, y_value, 0.05, verbose=False)[0].predict(testdata_df[testvariable])]
        else:
            testvariable = backward_regression(data, y_value, 0.05, verbose=False)[-1]
            testresult = [math.exp(i)  for i in backward_regression(
                data, y_value, 0.05, verbose=False)[0].predict(sm.add_constant(testdata_df)[testvariable])]

        locals()['width3freq' + str(data_freq)].loc[testdata_df.index,'foreresult']= list(testresult)
        for i in testdata_df.index:
            if (locals()['width3freq' + str(data_freq)].loc[i,'foreresult']>=1.2)*(
                    locals()['width3freq' + str(data_freq)].loc[i,'vixmaxcoe']>=1.2) != 0 or (
                    locals()['width3freq' + str(data_freq)].loc[i,'foreresult']<=1.2)*(
                    locals()['width3freq' + str(data_freq)].loc[i,'vixmaxcoe']<=1.2) !=0:
                locals()['width3freq' + str(data_freq)].loc[i, 'realjudge'] =1
    print (locals()['width3freq' + str(data_freq)])
    finalresult = locals()['width3freq' + str(data_freq)].realjudge.sum() / len(
        locals()['width3freq' + str(data_freq)].index)
    traintest_df['width3freq' + str(data_freq)] = finalresult


print (traintest_df)
traintest_df.to_excel((r'C:/Users/Tao/Desktop/投资/final/vix回测结果.xlsx'))




