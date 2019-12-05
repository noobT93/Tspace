import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import datetime as dt
from chinese_calendar import is_workday
import copy

matplotlib.rcParams['font.sans-serif'] = ['SimHei']
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['axes.unicode_minus'] = False

#设置回测区间
startdate = '2018-12-28'
enddate = '2019-11-01'

#获取期间内工作日和每月最后一日
def get_transferday(x,y):
# 获取起止之间的天数差
    daydiff = str(dt.datetime.strptime(y, '%Y-%m-%d') - dt.datetime.strptime(x, '%Y-%m-%d')).split(' ')[0]

# 获取所有日期
    end_date = dt.datetime.strptime(y, '%Y-%m-%d')
    oneday = dt.timedelta(days=1)
    adate = [y]
    for i in range(int(daydiff), 0, -1):
        end_date = end_date - oneday
        adate.append(str(end_date).split(' ')[0])
    td_df = pd.DataFrame(pd.Series(adate), columns=['日期'])
    td_df['非节假日'] = td_df['日期'].map(lambda x: is_workday(dt.datetime.strptime(x, '%Y-%m-%d')))
    td_df['周末'] = td_df['日期'].map(
        lambda x: dt.datetime.strptime(x, '%Y-%m-%d').isoweekday() == 6 or dt.datetime.strptime(x,
                                                                                                '%Y-%m-%d').isoweekday() == 7)
    td_df = td_df.drop(td_df[td_df['非节假日'] == False].index.values, axis=0)
    td_df = td_df.drop(td_df[td_df['周末'] == True].index.values, axis=0)
    wday = td_df['日期'].values.tolist()[::-1]  # list倒序用法

# 根据月份判断每月最后一个交易日
    tday = []
    for i in wday[:-1]:
        if i[5:7] != wday[int(wday.index(i)) + 1][5:7]:
            tday.append(i)
    tday.append(wday[-1])
# 日期模块没更新5.2 5.3
    if '2019-05-02' in wday:
        wday.remove('2019-05-02')
        wday.remove('2019-05-03')
    return [wday,tday]

#调仓组合总收益df
portreturn1 = pd.DataFrame(columns = ['日期','总收益'])
portreturn1['日期'] = get_transferday(startdate,enddate)[0]
portreturn1['总收益'] = 0.1
portreturn3 = pd.DataFrame(columns = ['日期','总收益'])
portreturn3['日期'] = get_transferday(startdate,enddate)[0]
portreturn3['总收益'] = 0.1

#一号调仓权重
#调仓调整后组合，获取组合每日收益率，x为调仓日，至下一个调仓日
#A：5日移动平均组合换手率*100 B：500指数/2500
#基准仓位50%，如果观察日前一日A<B，观察日A>=B，剩余仓位加仓30%
#若已经加仓，t日A>B,剩余仓位加仓20%
#若已经加仓，次日观察前一日A<B,剩余仓位加仓部分清仓
adjust_weight = pd.read_csv(r'D:/previous d/adjust_weight/调仓模型1.csv',engine = 'python',encoding='utf_8_sig').iloc[:,1:]
adjust_weight = adjust_weight.reset_index(drop = True)

#二号调仓权重
#调仓日当天，根据组合中个股周平均换手率（wind函数）超过3%的股票占组合个股总数（不含停牌）的比例来计算组合仓位调整幅度
#组合基准仓位为70%，周平均换手率超过3%的股票基准占比为20%。若周平均换手率超过3%的股票占比达到40%，则达到满仓。若周平均换手率超过3%的股票占比为10%，则仓位为半仓。通过利用分段二次函数进行构造模型。
#上升：(-5)*(x-0.2)^2+2.5*（x-0.2)
#下降：(-10)*(x-0.2)^2+1*(x-0.2)
weekturn_df = pd.read_csv(r'D:/previous d/daily_turn/换手率3%仓位.csv',engine = 'python',encoding='utf_8_sig')

#三号调仓权重
newportreturn = pd.DataFrame(columns = ['日期','总收益'])
newportreturn['日期'] = get_transferday(startdate,enddate)[0]
newportreturn['总收益'] = 0.1
newportreturn1 = pd.DataFrame(columns = ['日期','总收益'])
newportreturn1['日期'] = get_transferday(startdate,enddate)[0]
newportreturn1['总收益'] = 0.1
posdf = pd.read_csv(r'D:/previous d/non rehabilitation/totalpure仓位表（20181228）ver6.csv',engine = 'python',encoding='utf_8_sig')
posdf= posdf.set_index(posdf.columns[0])

#基准收益
zz500list = pd.read_csv(r'D:/previous d/daily_return/中证500行情.csv',engine = 'python',encoding='utf_8_sig').iloc[6:,1:]
zz500list.iloc[:,1] = zz500list.iloc[:,1] / zz500list.iloc[0,1]
zz500list = zz500list.reset_index(drop = True)
HS300list = pd.read_csv(r'D:/previous d/daily_return/沪深300行情.csv',engine = 'python',encoding='utf_8_sig').iloc[6:,1:]
HS300list.iloc[:,1] = HS300list.iloc[:,1] / HS300list.iloc[0,1]
HS300list = HS300list.reset_index(drop = True)

#获取组合每日收益率，x为调仓日，至下一个调仓日
def get_portreturn(x):
    score_time = x
    #半年内最高净值,半年内最低值,5天内最低值
    if score_time == get_transferday(startdate, enddate)[1][0]:
        maxvalue = 1
        minvalue = -1
        weekminvalue = 0
    else:
        maxvalue = max(zz500list.iloc[max(zz500list[zz500list['时间'] == score_time].index.values[0] - 126,0):zz500list[
                            zz500list['时间'] == score_time].index.values[0] , 1].values)
        minvalue = min(zz500list.iloc[max(zz500list[zz500list['时间'] == score_time].index.values[0] - 126,0):zz500list[
                            zz500list['时间'] == score_time].index.values[0] , 1].values)
        weekminvalue = min(zz500list.iloc[max(zz500list[zz500list['时间'] == score_time].index.values[0] - 7,0):zz500list[
                            zz500list['时间'] == score_time].index.values[0] , 1].values)
    weightlist = pd.read_csv(r'D:/previous d/final 500list/最终500数据%s.csv' % score_time, engine='python', encoding='utf_8_sig').T
    newweightlist = pd.read_csv(r'D:/previous d/final 500list/纯粹500数据%s.csv'% score_time,engine = 'python',encoding='utf_8_sig').T
    newweightlist1 = copy.deepcopy(newweightlist)
    if score_time == get_transferday(startdate,enddate)[1][0]:
        weightlist1 = weightlist.loc['权重', :].values * 1
        weightlist3 = weightlist.loc['权重', :].values * 1
        newweightlist = newweightlist.loc['权重', :].values * 1
        newweightlist1 = newweightlist1.loc['权重', :].values * 1
    else:
        weightlist1 = weightlist.loc['权重', :].values * portreturn1.iloc[
            portreturn1[portreturn1['日期'] == score_time].index.values - 1, 1].values
        weightlist3 = weightlist.loc['权重', :].values * portreturn3.iloc[
            portreturn3[portreturn3['日期'] == score_time].index.values - 1, 1].values
        newweightlist = newweightlist.loc['权重', :].values * newportreturn.iloc[
            newportreturn[newportreturn['日期'] == score_time].index.values - 1, 1].values
        newweightlist1 = newweightlist1.loc['权重', :].values * newportreturn1.iloc[
            newportreturn1[newportreturn1['日期'] == score_time].index.values - 1, 1].values
    portvalue = pd.read_csv(r'D:/previous d/华商银行/理财/放表/daily_return/500每日行情%s.csv'% score_time,engine = 'python',encoding='utf_8_sig')
    portvalue.iloc[1:,1:] = portvalue.iloc[1:,1:].applymap(lambda y : round(float(y),5))
    newportvalue = pd.read_csv(r'D:/previous d/华商银行/理财/放表/daily_return/纯粹500每日行情%s.csv'% score_time,engine = 'python',encoding='utf_8_sig')
    newportvalue.iloc[1:, 1:] = newportvalue.iloc[1:, 1:].applymap(lambda y: round(float(y), 5))
    newportvalue1 = copy.deepcopy(newportvalue)
    newportvalue1.columns = newportvalue1.iloc[0, :]
    newportvalue1 = newportvalue1.drop(columns=newportvalue1.columns[0])
    newportvalue1 = newportvalue1.drop(0)
    newportvalue1 = newportvalue1.reset_index(drop=True)

    # 判断调仓日是否为涨停板，迭代遍历columns，若涨幅超过9.95%，则从首个数据至该列首个涨跌幅低于9.95%数据全部赋值为1
    for z in [i for i, x in enumerate(portvalue.columns[1:],start = 1)]:
        if portvalue.iloc[1, z] >= 1.095:
            try:
                portvalue.iloc[1:portvalue.iloc[1:,z][portvalue.iloc[1:,z] < 1.095].index.values[0], z] = 1
            except:
                portvalue.iloc[1:,z] = 1
    for z in [i for i, x in enumerate(newportvalue.columns[1:],start = 1)]:
        if newportvalue.iloc[1, z] >= 1.095:
            try:
                newportvalue.iloc[1:newportvalue.iloc[1:,z][newportvalue.iloc[1:,z] < 1.095].index.values[0], z] = 1
            except:
                newportvalue.iloc[1:,z] = 1
    portvalue1 = copy.deepcopy(portvalue)
    portvalue3 = copy.deepcopy(portvalue)
    for port1 in portvalue1.index[1:]:
        try:
            portvalue1.iloc[port1, 1:] = (portvalue1.iloc[port1, 1:].values-1)*adjust_weight.iloc[
                adjust_weight[adjust_weight['时间'] ==portvalue1.iloc[port1,0]].index.values,4].values + 1
        except:
            portvalue1.iloc[port1, 1:] = portvalue1.iloc[port1, 1:]
    for port3 in portvalue3.index[1:]:
        portvalue3.iloc[port3, 1:] = (portvalue3.iloc[port3, 1:].values-1)*weekturn_df.iloc[
            weekturn_df[weekturn_df['时间'] ==portvalue3.iloc[port3,0]].index.values,1].values + 1
    #涨跌幅进行调仓调整
    if score_time == get_transferday(startdate,enddate)[1][0]:
        portvalue1.iloc[7, 1:] = portvalue1.iloc[7, 1:] * weightlist1
        portvalue1.iloc[7:, 1:] = portvalue1.iloc[7:, 1:].cumprod(axis=0)
        portvalue1.iloc[1:7,1:] = 0
        portvalue3.iloc[7, 1:] = portvalue3.iloc[7, 1:] * weightlist3
        portvalue3.iloc[7:, 1:] = portvalue3.iloc[7:, 1:].cumprod(axis=0)
        portvalue3.iloc[1:7, 1:] = 0
        newportvalue.iloc[7, 1:] = newportvalue.iloc[7, 1:] * newweightlist
        newportvalue.iloc[7:, 1:] = newportvalue.iloc[7:, 1:].cumprod(axis=0)
        newportvalue.iloc[1:7, 1:] = 0
        for num in newportvalue1.index[6:]:
            for columnsname in newportvalue1.columns:
                newportvalue1.iloc[num, newportvalue1.columns.tolist().index(
                    columnsname)] = (newportvalue1.iloc[num, newportvalue1.columns.tolist().index(
                    columnsname)] - 1) * posdf.iloc[posdf.index.tolist().index(
                    columnsname), posdf.columns.tolist().index(score_time) + num] + 1
        newportvalue1.iloc[6, :] = newportvalue1.iloc[6, :] * newweightlist1
        newportvalue1.iloc[6:, :] = newportvalue1.iloc[6:, :].cumprod(axis=0)
        newportvalue1.iloc[:6, :] = 0
    else:
        portvalue1.iloc[1, 1:] = portvalue1.iloc[1, 1:] * weightlist1
        portvalue1.iloc[1:, 1:] = portvalue1.iloc[1:, 1:].cumprod(axis=0)
        portvalue3.iloc[1, 1:] = portvalue3.iloc[1, 1:] * weightlist3
        portvalue3.iloc[1:, 1:] = portvalue3.iloc[1:, 1:].cumprod(axis=0)
        newportvalue.iloc[1, 1:] = newportvalue.iloc[1, 1:] * newweightlist
        newportvalue.iloc[1:, 1:] = newportvalue.iloc[1:, 1:].cumprod(axis=0)
        #分段逻辑，计算个股收益，判断中证500指数
        # 大于0.95*半年最高净值满仓（不需要处理），小于0.95*半年最高净值、大于0.85*半年最高净值按照三号调仓，
        # 小于0.85*半年最高净值按照换手指数/小于0.85*半年最高净值空仓（当日净值高于5日内最低净值且5日内最低净值高于半年最低净值按照三号调仓）
        for num in newportvalue1.index:
            if maxvalue*0.95>zz500list.iloc[zz500list[zz500list['时间'] == score_time].index.values[0] +num, 1] >= 0.85* maxvalue:
                for columnsname in newportvalue1.columns:
                    newportvalue1.iloc[num, newportvalue1.columns.tolist().index(
                        columnsname)] = (newportvalue1.iloc[num, newportvalue1.columns.tolist().index(
                        columnsname)] - 1) * posdf.iloc[posdf.index.tolist().index(
                        columnsname), posdf.columns.tolist().index(score_time) + num] + 1
            '''if 0.85* maxvalue>zz500list.iloc[zz500list[zz500list['时间'] == score_time].index.values[0] +num, 1]:
                newportvalue1.iloc[num,:] = (newportvalue1.iloc[num,:].values-1)*adjust_weight.iloc[
                int(adjust_weight[adjust_weight['时间'] ==score_time].index.values)+num,4] + 1'''
            if zz500list.iloc[zz500list[zz500list['时间'] == score_time].index.values[0] +num, 1]< 0.85* maxvalue:
                if zz500list.iloc[zz500list[zz500list['时间'] == score_time].index.values[0] +num,
                                  1] > weekminvalue >= minvalue :
                    for columnsname in newportvalue1.columns:
                        newportvalue1.iloc[num, newportvalue1.columns.tolist().index(
                            columnsname)] = (newportvalue1.iloc[num, newportvalue1.columns.tolist().index(
                            columnsname)] - 1) * posdf.iloc[posdf.index.tolist().index(
                            columnsname), posdf.columns.tolist().index(score_time) + num] + 1
                else:
                    newportvalue1.iloc[num, :] = 1

        newportvalue1.iloc[0, :] = newportvalue1.iloc[0, :] * newweightlist1
        newportvalue1.iloc[:, :] = newportvalue1.iloc[:, :].cumprod(axis=0)
    portvalue1['组合收益'] = portvalue1.iloc[1:, 1:].sum(axis=1)
    portvalue3['组合收益'] = portvalue3.iloc[1:, 1:].sum(axis=1)
    newportvalue1['组合收益'] = newportvalue1.iloc[:, :].sum(axis=1)
    newportvalue['组合收益'] = newportvalue.iloc[1:, 1:].sum(axis=1)
    return [portvalue1,portvalue3,newportvalue,newportvalue1]

for i in get_transferday(startdate,enddate)[1]:
    if i == get_transferday(startdate, enddate)[1][-1]:
        portreturn1.iloc[int(portreturn1[portreturn1['日期'] == i].index.values):, 1] = get_portreturn(i)[0].loc[1:,'组合收益'].values
        portreturn3.iloc[int(portreturn3[portreturn3['日期'] == i].index.values):, 1] = get_portreturn(i)[1].loc[1:,
                                                                                      '组合收益'].values
        newportreturn.iloc[int(newportreturn[newportreturn['日期'] == i].index.values):, 1] = get_portreturn(i)[2].loc[1:,
                                                                                      '组合收益'].values
        newportreturn1.iloc[int(newportreturn1[newportreturn1['日期'] == i].index.values):, 1] = get_portreturn(i)[3].loc[:,
                                                                                      '组合收益'].values
    else:
        next_x = get_transferday(startdate, enddate)[1][get_transferday(startdate, enddate)[1].index(i) + 1]
        portreturn1.iloc[int(portreturn1[portreturn1['日期'] == i].index.values):int(portreturn1[
            portreturn1['日期'] == next_x].index.values), 1] = get_portreturn(i)[0].iloc[1:,-1].values
        portreturn3.iloc[int(portreturn3[portreturn3['日期'] == i].index.values):int(portreturn3[
            portreturn3['日期'] == next_x].index.values), 1] = get_portreturn(i)[1].iloc[1:,-1].values
        newportreturn.iloc[int(newportreturn[newportreturn['日期'] == i].index.values):int(newportreturn[
            newportreturn['日期'] == next_x].index.values), 1] = get_portreturn(i)[2].iloc[1:,-1].values
        newportreturn1.iloc[int(newportreturn1[newportreturn1['日期'] == i].index.values):int(newportreturn1[
            newportreturn1['日期'] == next_x].index.values), 1] = get_portreturn(i)[3].iloc[:,-1].values

portreturn1 = portreturn1.iloc[6:,:]
portreturn1.iloc[:,1] = portreturn1.iloc[:,1] / portreturn1.iloc[0,1]
portreturn3 = portreturn3.iloc[6:,:]
portreturn3.iloc[:,1] = portreturn3.iloc[:,1] / portreturn3.iloc[0,1]
newportreturn = newportreturn.iloc[6:,:]
newportreturn.iloc[:,1] = newportreturn.iloc[:,1] / newportreturn.iloc[0,1]
newportreturn1 = newportreturn1.iloc[6:,:]
newportreturn1.iloc[:,1] = newportreturn1.iloc[:,1] / newportreturn1.iloc[0,1]

# 获取收益率
def backtest_plot(x):
    backtestframe = pd.DataFrame(index = portreturn1['日期'].map(lambda x: dt.datetime.strptime(x, '%Y-%m-%d')))
    indexlen = len(backtestframe.index.values)//6
    xlabel1 = [backtestframe.index[0],backtestframe.index[indexlen],backtestframe.index[indexlen*2],backtestframe.index[indexlen*3],
                                 backtestframe.index[indexlen*4],backtestframe.index[indexlen*5],backtestframe.index[-1]]
    xlabel2 = [dt.datetime.strftime(i, '%Y-%m-%d') for i in xlabel1]

    backtestframe['中证500'] = zz500list['500涨跌'].values
    backtestframe['沪深300'] = HS300list['沪深300涨跌'].values
    backtestframe['纯粹策略(3个月观察期）'] = newportreturn1['总收益'].values
    backtestframe['模拟组合-换手/指数比较'] = portreturn1['总收益'].values
    backtestframe['模拟组合-换手3%'] = portreturn3['总收益'].values

    time_delta = (dt.datetime.strptime(enddate, '%Y-%m-%d') - dt.datetime.strptime(startdate, '%Y-%m-%d')).days
    if time_delta >= 365:
        abs_return_adjust1 = round(((backtestframe.iloc[-1, 2]) ** (1 / (time_delta / 365)) - 1) * 100, 4)
        abs_return_benchmark = round(((backtestframe.iloc[-1,0])**(1/(time_delta/365)) -1)*100,4)
    else:
        abs_return_adjust1 = round(((backtestframe.iloc[-1, 2]) - 1) * 100, 4)
        abs_return_benchmark = round(((backtestframe.iloc[-1,0]) -1)*100,4)
    relative_return1 = round(abs_return_adjust1 - abs_return_benchmark,4)
    backtestframe.plot(xlim = [backtestframe.index[0],backtestframe.index[-1]],figsize = (15,8))
    plt.xticks(xlabel1,xlabel2)
    plt.text(dt.datetime.strptime('2019-05-08', '%Y-%m-%d'), 1.35,
             "纯粹策略绝对收益:%s%%\n纯粹策略相对收益:%s%%" % (
             abs_return_adjust1, relative_return1,), size=10, style="italic",bbox=dict(alpha=0.2))
    backtestframe.to_csv(r'D:/previous d/adjust_weight/收益对比.csv' , index=True, encoding='utf_8_sig')
    plt.show()
# 输出回测图像
backtest_plot(get_transferday(startdate,enddate))
