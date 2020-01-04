import pandas as pd
import datetime as dt
import copy
import numpy as np
import time
start_time = time.time()
real_df =pd.read_excel(r'C:/Users/Tao/Desktop/投资/牛熊/1231/牛熊数据计算.xlsx',sheet_name= 4)
real_df.set_index(real_df['日期'],inplace=True)
real_df = real_df.drop(columns = ['日期'])
real_df.index = real_df.index.map(lambda y : dt.datetime.strftime(y, '%Y-%m-%d'))
real_df['止盈止损最高-振幅'] = real_df['止盈止损预测最高'] - real_df['止盈止损预测振幅']
real_df['止盈止损最低+振幅'] = real_df['止盈止损预测最低'] + real_df['止盈止损预测振幅']
real_df['方向最高-振幅'] = real_df['方向预测最高'] - real_df['方向预测振幅']
real_df['方向最低+振幅'] = real_df['方向预测最低'] + real_df['方向预测振幅']
real_df.dropna(inplace= True)

finalresult = pd.DataFrame(index=['highlow','high-swing','low+swing'], columns=['单边最大期望', '单边牛止盈', '单边牛止损',
                            '单边熊止盈','单边熊止损','盈损单边模型','双边最大期望',
                            '双边牛止盈','双边牛止损', '双边熊止盈','双边熊止损', '盈损双边模型','合并最大期望'])
finalresult.fillna(0, inplace=True)

endresult = pd.DataFrame(index=[i/10 for i in range(19,26, 1)], columns=['合并最大期望', '方向模型','单边牛止盈', '单边牛止损',
                            '单边熊止盈','单边熊止损','双边牛止盈','双边牛止损', '双边熊止盈','双边熊止损',
                                                                         '盈损单边模型','盈损双边模型'])
endresult.fillna(0, inplace=True)

for bearsl in [i/10 for i in range(19,26, 1)]:
    num = range(19,26, 1).index(int(bearsl*10))
    finalresult.iloc[:, :] = 0
    for model in ['highlow', 'high-swing', 'low+swing']:
        singledf = pd.DataFrame(columns=real_df.columns)
        if model == 'highlow':
            high = 5
            low = 6
        if model == 'high-swing':
            high = 5
            low = 10
        if model == 'low+swing':
            high = 11
            low = 6
        singledf = real_df
        for bull in [i / 10 for i in range(12,26, 1)]:
            for bear in [i / 10 for i in range(12,26, 1)]:
                for bullsl in [i / 10 for i in range(12,26, 1)]:
                    for sinmodel in ['highlow', 'high-swing', 'low+swing']:
                        singledf['结果'] = 0
                        # 确定单边模型
                        if sinmodel == 'highlow':
                            sinhigh = 0
                            sinlow = 1
                        if sinmodel == 'high-swing':
                            sinhigh = 0
                            sinlow = 8
                        if sinmodel == 'low+swing':
                            sinhigh = 9
                            sinlow = 1
                        for sin in singledf.index:
                            #设置同边对边止损系数，若为对边则为1，若为同边则为2，默认为1
                            sindou = 1
                            # 牛
                            if abs(singledf.iloc[singledf.index.tolist().index(sin), high] - 1) - abs(
                                    singledf.iloc[singledf.index.tolist().index(sin), low] - 1) > 0:
                                if singledf.iloc[singledf.index.tolist().index(sin), low] > 1:
                                    sindou =2
                                else:
                                    sindou = 1
                                # 对边实际偏离大于止损系数*对边预测偏离且实际偏离小于预测偏离*压缩系数，亏止损系数*对边预计偏离度
                                if (1 - singledf.iloc[
                                    singledf.index.tolist().index(sin), sinlow]) * bullsl * sindou < 1 - \
                                        singledf.iloc[
                                            singledf.index.tolist().index(sin), 4] and singledf.iloc[
                                    singledf.index.tolist().index(
                                        sin), 3] - 1 < (
                                        singledf.iloc[singledf.index.tolist().index(sin), sinhigh] - 1) * bull:
                                    singledf.iloc[singledf.index.tolist().index(sin), -1] = -bullsl * (
                                            1 - singledf.iloc[singledf.index.tolist().index(sin), sinlow])
                                # 对边实际偏离大于止损系数*对边预测偏离且实际偏离大于预测偏离*压缩系数，亏0.5*（压缩系数*预计偏离度 - 止损系数*对边预测偏离）
                                if (1 - singledf.iloc[
                                    singledf.index.tolist().index(sin), sinlow]) * bullsl * sindou < 1 - \
                                        singledf.iloc[
                                            singledf.index.tolist().index(sin), 4] and singledf.iloc[
                                    singledf.index.tolist().index(
                                        sin), 3] - 1 > (
                                        singledf.iloc[singledf.index.tolist().index(sin), sinhigh] - 1) * bull:
                                    singledf.iloc[singledf.index.tolist().index(sin), -1] = 0.5 * (
                                            (singledf.iloc[
                                                 singledf.index.tolist().index(sin), sinhigh] - 1) * bull - (
                                                    1 - singledf.iloc[
                                                singledf.index.tolist().index(sin), sinlow]) * bullsl)
                                # 实际偏离大于预测偏离*压缩系数且对边实际偏离小于止损系数*对边预测偏离，赚压缩系数
                                if (1 - singledf.iloc[
                                    singledf.index.tolist().index(sin), sinlow]) * bullsl * sindou > 1 - \
                                        singledf.iloc[
                                            singledf.index.tolist().index(sin), 4] and (
                                        singledf.iloc[singledf.index.tolist().index(
                                            sin), sinhigh] - 1) * bull < singledf.iloc[
                                    singledf.index.tolist().index(sin), 3] - 1:
                                    singledf.iloc[singledf.index.tolist().index(
                                        sin), -1] = (singledf.iloc[
                                                         singledf.index.tolist().index(
                                                             sin), sinhigh] - 1) * bull
                            # 熊
                            if abs(singledf.iloc[singledf.index.tolist().index(sin), high] - 1) - abs(
                                    singledf.iloc[singledf.index.tolist().index(sin), low] - 1) < 0:
                                # 设置同边对边止损系数，若为对边则为1，若为同边则为2，默认为1
                                if singledf.iloc[singledf.index.tolist().index(sin), low] > 1:
                                    sindou =2
                                else:
                                    sindou = 1
                                # 对边实际偏离大于止损系数*对边预测偏离且实际偏离小于预测偏离*压缩系数，亏止损系数*对边预计偏离度
                                if (singledf.iloc[
                                        singledf.index.tolist().index(sin), sinhigh] - 1) * sindou * 2 < \
                                        singledf.iloc[
                                            singledf.index.tolist().index(sin), 3] - 1 and 1 - singledf.iloc[
                                    singledf.index.tolist().index(
                                        sin), 4] < (
                                        1 - singledf.iloc[singledf.index.tolist().index(sin), sinhigh]) * bear:
                                    singledf.iloc[singledf.index.tolist().index(sin), -1] = -bearsl * (
                                            singledf.iloc[singledf.index.tolist().index(sin), sinhigh] - 1)
                                # 对边实际偏离大于止损系数*对边预测偏离且实际偏离大于预测偏离*压缩系数，亏0.5*（压缩系数*预计偏离度 - 止损系数*对边预测偏离）
                                if (singledf.iloc[
                                        singledf.index.tolist().index(sin), sinhigh] - 1) * sindou * 2 < \
                                        singledf.iloc[
                                            singledf.index.tolist().index(sin), 3] - 1 and 1 - singledf.iloc[
                                    singledf.index.tolist().index(
                                        sin), 4] > (
                                        1 - singledf.iloc[singledf.index.tolist().index(sin), sinlow]) * bear:
                                    singledf.iloc[singledf.index.tolist().index(sin), -1] = 0.5 * (
                                            (1 - singledf.iloc[
                                                singledf.index.tolist().index(sin), sinlow]) * bear - (
                                                    singledf.iloc[
                                                        singledf.index.tolist().index(
                                                            sin), sinhigh] - 1) * bearsl)
                                # 实际偏离大于预测偏离*压缩系数且对边实际偏离小于止损系数*对边预测偏离，赚压缩系数
                                if (singledf.iloc[
                                        singledf.index.tolist().index(sin), sinhigh] - 1) * sindou * 2 > \
                                        singledf.iloc[
                                            singledf.index.tolist().index(sin), 3] - 1 and (
                                        1 - singledf.iloc[singledf.index.tolist().index(
                                    sin), sinlow]) * bear < 1 - singledf.iloc[
                                    singledf.index.tolist().index(sin), 4]:
                                    singledf.iloc[singledf.index.tolist().index(
                                        sin), -1] = (1 - singledf.iloc[
                                        singledf.index.tolist().index(sin), sinlow]) * bear
                        sinsum = singledf['结果'].values.sum()
                        if sinsum > finalresult.iloc[finalresult.index.tolist().index(model), 0]:
                            finalresult.iloc[finalresult.index.tolist().index(model), 0] = copy.deepcopy(sinsum)
                            finalresult.iloc[finalresult.index.tolist().index(model), 1] = copy.deepcopy(bull)
                            finalresult.iloc[finalresult.index.tolist().index(model), 2] = copy.deepcopy(bullsl)
                            finalresult.iloc[finalresult.index.tolist().index(model), 3] = copy.deepcopy(bear)
                            finalresult.iloc[finalresult.index.tolist().index(model), 4] = copy.deepcopy(bearsl)
                            finalresult.iloc[finalresult.index.tolist().index(model), 5] = copy.deepcopy(
                                sinmodel)

        finalresult.iloc[finalresult.index.tolist().index(model), 12] = copy.deepcopy(
            finalresult.iloc[finalresult.index.tolist().index(model), 0]) + copy.deepcopy(
            finalresult.iloc[finalresult.index.tolist().index(model), 6])
    endresult.iloc[num, 0] = copy.deepcopy(max(finalresult['合并最大期望'].values))
    endresult.iloc[num, 1] = copy.deepcopy(finalresult[finalresult['合并最大期望'] ==endresult.iloc[num, 0]].index[0])
    endresult.iloc[num, 2] = copy.deepcopy(finalresult.iloc[finalresult.index.tolist().index(
        finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 1])
    endresult.iloc[num, 3] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 2])
    endresult.iloc[num, 4] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 3])
    endresult.iloc[num, 5] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 4])
    endresult.iloc[num, 6] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 7])
    endresult.iloc[num, 7] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 8])
    endresult.iloc[num, 8] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 9])
    endresult.iloc[num, 9] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 10])
    endresult.iloc[num, 10] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 5])
    endresult.iloc[num, 11] = copy.deepcopy(
        finalresult.iloc[finalresult.index.tolist().index(
            finalresult[finalresult['合并最大期望'] == endresult.iloc[num, 0]].index[0]), 11])
endresult.to_csv(r'C:/Users/tao/Desktop/投资/牛熊/pool/牛熊结果（1231-2）（1.2-2.5,1.2-2.5）.csv' , index=True, encoding='utf_8_sig')
print('运行时间%s' % (time.time() - start_time))
