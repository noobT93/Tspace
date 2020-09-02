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

spyinit_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/final/spy基本数据.xlsx',sheet_name=0)
spyinit_df.set_index('Date', inplace=True)

#先靠倒数第二个数据定位
#测试集为2年，倒数第二往前推2年的下月1日，不确定最后一周是否为下一月，得到测试集起点

teststart_date = dt.datetime(spyinit_df.index[-2].year-2,spyinit_df.index[-2].month+1,1,0,0,0)
teststart_point = [i for i in spyinit_df.index if i >= teststart_date][0]

#获取index 年月数据
date_init = [str(i.year) + str(i.month).zfill(2) for i in  spyinit_df.index]
date_init = sorted(set(date_init),key = date_init.index)
#测试起点至2年后终点，不可直接-1，不确定最后一周
date_init= date_init[date_init.index(str(teststart_point.year) + str(teststart_point.month).zfill(
    2)):1+date_init.index(str(teststart_point.year+2) + str(teststart_point.month).zfill(2))]

traintest_df = pd.DataFrame(index = [date_init[0] +'-'+date_init[-1]],columns = ['width3freq1','width3freq3'])

#训练集样本为3年
#每月更新
data_width = 3
for data_freq in [1,3]:
    #spy不需要调整，直接判断最大正向波动
    locals()['width3freq' + str(data_freq)] = pd.DataFrame(index = spyinit_df.index[
        list(spyinit_df.index).index(teststart_point):-1],columns = ['spymaxcoe','foreresult',
        'forejudge','realjudge','adjresult'])
    locals()['width3freq' + str(data_freq)].forejudge = 0
    locals()['width3freq' + str(data_freq)].adjjudge = 0
    locals()['width3freq' + str(data_freq)].spymaxcoe = [abs(math.exp(i)-1) for i in spyinit_df[
        teststart_point<=spyinit_df.index][:-1].spymaxcoe]
    for reali in locals()['width3freq' + str(data_freq)].index:
        if locals()['width3freq' + str(data_freq)].loc[reali,'spymaxcoe'] <= 0.02:
            locals()['width3freq' + str(data_freq)].loc[reali, 'realjudge'] = 1
        else:
            locals()['width3freq' + str(data_freq)].loc[reali, 'realjudge'] = 0
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
        traindata_df = spyinit_df[dt.datetime(
            int(datenum[:4])-data_width,int(datenum[4:]),1,0,0,0) <=spyinit_df.index]
        traindata_df = traindata_df[traindata_df.index<dt.datetime(
            int(datenum[:4]),int(datenum[4:]),1,0,0,0)][3:-1]
        y_value = traindata_df.iloc[:, 0].values
        data = traindata_df.iloc[:, 1:]
        #测试集宽度
        testdata_df = spyinit_df[dt.datetime(
            int(datenum[:4]),int(datenum[4:]),1,0,0,0) <=spyinit_df.index][1:]
        testdata_df = testdata_df[testdata_df.index <dt.datetime(
            endyear,endmonth,1,0,0,0) ]
        #获取自变量
        if 'const' not in backward_regression(data, y_value, 0.05, verbose=False)[-1]:
            testvariable = backward_regression(data, y_value, 0.05, verbose=False)[-1]
            testresult = [abs(math.exp(i) - 1) for i in backward_regression(
                data, y_value, 0.05, verbose=False)[0].predict(testdata_df[testvariable])]
            # 生成preforereulst
            traindata_df['adjfore'] = [abs(math.exp(i) - 1) for i in backward_regression(
                data, y_value, 0.05, verbose=False)[0].predict(traindata_df[testvariable])]
        else:
            testvariable = backward_regression(data, y_value, 0.05, verbose=False)[-1]
            testresult = [abs(math.exp(i) - 1) for i in backward_regression(
                data, y_value, 0.05, verbose=False)[0].predict(sm.add_constant(testdata_df)[testvariable])]
            # 生成preforereulst
            traindata_df['adjfore'] = [abs(math.exp(i) - 1) for i in backward_regression(
                data, y_value, 0.05, verbose=False)[0].predict(sm.add_constant(traindata_df)[testvariable])]

        traindata_df['adjdiff'] = 1
        # 生成adjratio
        for i in traindata_df.index:
            if math.exp(traindata_df.loc[i, traindata_df.columns[0]])-1 > 0.02 > traindata_df.loc[i, 'adjfore']:
                traindata_df.loc[i, 'adjdiff'] = (math.exp(traindata_df.loc[i, traindata_df.columns[0]])-1) / traindata_df.loc[
                    i, 'adjfore']

        adjratio = st.gmean(traindata_df.adjdiff)

        locals()['width3freq' + str(data_freq)].loc[testdata_df.index,'foreresult']= list(testresult)
        for i in testdata_df.index:
            locals()['width3freq' + str(data_freq)].loc[i, 'adjresult'] = locals()['width3freq' + str(
                data_freq)].loc[i, 'foreresult'] * adjratio

            if 0.004<=locals()['width3freq' + str(data_freq)].loc[i,'adjresult']<=0.02:
                locals()['width3freq' + str(data_freq)].loc[i, 'forejudge'] =1
    print (locals()['width3freq' + str(data_freq)])
    finalresult_df = locals()['width3freq' + str(data_freq)][locals()['width3freq' + str(data_freq)].forejudge == 1]
    finalresult = finalresult_df.realjudge.sum() / len(finalresult_df.index)
    traintest_df['width3freq' + str(data_freq)] = finalresult


print (traintest_df)
traintest_df.to_excel((r'C:/Users/Tao/Desktop/投资/final/spy回测结果adj.xlsx'))




