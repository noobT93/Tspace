import pandas as pd
import statsmodels.api as sm
from stepwise_regression import step_reg
import numpy as np
import math

price_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/下次用这个/stoploss数据整理（以前一日收盘价为底） - 对比.xlsx',
                         sheet_name=0).iloc[:, :]
vol_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/下次用这个/stoploss数据整理（以前一日收盘价为底） - 对比.xlsx',
                       sheet_name=1).iloc[:, :]
df_final = pd.DataFrame(index = price_df.index,columns = ['spycoe','closevarspy0','closevarspy1',
                'closevarspy2','maxvarspy0','maxvarspy1','maxvarspy2',
                'closevarvix0','closevarvix1','closevarvix2','maxvarvix0','maxvarvix1','maxvarvix2',
                'closevardji0','closevardji1','closevardji2','maxvardji0','maxvardji1','maxvardji2',
                'closevarixic0','closevarixic1','closevarixic2','maxvarixic0','maxvarixic1','maxvarixic2',
                'closevarspx0','closevarspx1','closevarspx2','maxvarspx0','maxvarspx1','maxvarspx2',
                'closevarhsi0','closevarhsi1','closevarhsi2','maxvarhsi0','maxvarhsi1','maxvarhsi2',
                'closevarsh0','closevarsh1','closevarsh2','maxvarsh0','maxvarsh1','maxvarsh2',
                'closevarsz0', 'closevarsz1', 'closevarsz2', 'maxvarsz0','maxvarsz1', 'maxvarsz2',
                'closevarcy0', 'closevarcy1', 'closevarcy2', 'maxvarcy0','maxvarcy1', 'maxvarcy2',
                'closevolspy0','closevolspy1','closevolspy2','maxvolspy0','maxvolspy1','maxvolspy2',
                'closevoldji0','closevoldji1','closevoldji2','maxvoldji0','maxvoldji1','maxvoldji2',
                'closevolixic0','closevolixic1','closevolixic2','maxvolixic0','maxvolixic1','maxvolixic2',
                'closevolspx0','closevolspx1','closevolspx2','maxvolspx0','maxvolspx1','maxvolspx2',
                'closevolhsi0','closevolhsi1','closevolhsi2','maxvolhsi0','maxvolhsi1','maxvolhsi2',
                'closevolsh0','closevolsh1','closevolsh2','maxvolsh0','maxvolsh1','maxvolsh2',
                'closevolsz0','closevolsz1','closevolsz2','maxvolsz0','maxvolsz1','maxvolsz2',
                'closevolcy0', 'closevolcy1', 'closevolcy2', 'maxvolcy0','maxvolcy1', 'maxvolcy2'])
df_final.fillna(0,inplace =True)
df_final['spycoe'] = price_df.iloc[:,list(price_df.columns).index('spycoe')].values

#美股
for i in ['spy','vix','dji','ixic','spx','hsi','sh','sz','cy']:
    #计算收盘价波动率
    df_final.loc[1:,'closevar'+i+'0'] = list(price_df.loc[
    1:, 'pct'+i+'b'].values / price_df.loc[:len(price_df.index)-2, 'pct'+ i + 'b'].values)
    df_final.loc[2:,'closevar'+i+'1'] = df_final.loc[1:,'closevar'+i+'0'].values[:-1]
    df_final.loc[3:,'closevar'+i+'2'] = df_final.loc[1:,'closevar'+i+'0'].values[:-2]

    #计算周内最大波动率
    for numi in df_final.index[1:]:
        df_final.loc[numi, 'maxvar' + i + '0'] = price_df.loc[
    numi, 'high'+i+'b'] / price_df.loc[numi-1, 'pct'+ i + 'b'] if price_df.loc[
    numi, 'high'+i+'b'] + price_df.loc[numi, 'low'+i+'b'] >= 2*price_df.loc[
    numi-1, 'pct'+i+'b'] else price_df.loc[numi, 'low'+i+'b'] / price_df.loc[numi-1, 'pct'+ i + 'b']
    df_final.loc[2:,'maxvar'+i+'1'] = df_final.loc[1:,'maxvar'+i+'0'].values[:-1]
    df_final.loc[3:,'maxvar'+i+'2'] = df_final.loc[1:,'maxvar'+i+'0'].values[:-2]

    if i != 'vix':
        #计算周间最后一个交易日的交易量波动率
        df_final.loc[1:, 'closevol' + i + '0'] = vol_df.loc[
                                                1:, 'cvol' + i + '1'].values / vol_df.loc[:len(price_df.index) - 2,
                                                                               'cvol' + i + '1'].values
        df_final.loc[2:, 'closevol' + i + '1'] = df_final.loc[1:, 'closevol' + i + '0'].values[:-1]
        df_final.loc[3:, 'closevol' + i + '2'] = df_final.loc[1:, 'closevol' + i + '0'].values[:-2]
        #选取周内最大波动率对应交易量，作为交易量波动率
        for yy in df_final.index[1:]:
            if price_df.loc[yy, 'high'+i+'b'] + price_df.loc[yy, 'low'+i+'b'] >= 2*price_df.loc[
                yy-1, 'pct'+i+'b']:
                df_final.loc[yy, 'maxvol' + i + '0'] = vol_df.loc[
                yy, 'hvol' + i + '1'] / vol_df.loc[yy - 1, 'cvol' + i + '1']
            else:
                df_final.loc[yy, 'maxvol' + i + '0'] = vol_df.loc[
                yy, 'lvol' + i + '1'] / vol_df.loc[yy - 1, 'cvol' + i + '1']
        df_final.loc[2:, 'maxvol' + i + '1'] = df_final.loc[1:, 'maxvol' + i + '0'].values[:-1]
        df_final.loc[3:, 'maxvol' + i + '2'] = df_final.loc[1:, 'maxvol' + i + '0'].values[:-2]

df_final.iloc[:,1:] = df_final.iloc[:,1:].applymap(lambda y:math.log(y) if y != 0 else y)

df_final.to_csv(r'C:/Users/Tao/Desktop/投资/spy/stoploss检查用111.csv', index=False, encoding='utf_8_sig')

y_value = df_final.iloc[3:,0].values
data = df_final.iloc[3:,1:]

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
    print (model.summary())
    return included

finallist =['spycoe']
if 'const' in backward_regression(data,y_value,0.05,verbose = False)[:]:
    finallist.extend(backward_regression(data, y_value, 0.05, verbose=False)[:])
    finallist.remove('const')
    data = df_final[finallist]
    i_value = list((data == 0).any(axis=1)).index(False)
    y_value = data.iloc[i_value:, 0]
    data = data.iloc[i_value:, 1:]
    data = sm.add_constant(data)
else:
    finallist.extend(backward_regression(data,y_value,0.05,verbose = False)[:])
    data = df_final[finallist]
    i_value = list((data == 0).any(axis=1)).index(False)
    y_value = data.iloc[i_value:, 0]
    data = data.iloc[i_value:, 1:]

model = sm.OLS(y_value, data).fit()
result_df = pd.DataFrame(index=range(0, len(model.predict()), 1), columns=['equation', 'expresult'])
final_equ = ['(' + str(model.params[i]) + '*' + i + ')' if i != 'const' else str(
    model.params[i]) for i in model.params.keys()]

result_df.iloc[0, 0] = '=exp(' + ('+'.join(final_equ)) + ')'
paralist = list(model.params.keys())
try:
    paralist.remove('const')
except:
    pass

result_df.iloc[1:1+len(paralist),0] = paralist

result_df.iloc[:, 1] = [math.exp(i) for i in model.predict()]
result_df.to_csv(r'C:/Users/Tao/Desktop/投资/spy/stoplossresult111.csv', index=False, encoding='utf_8_sig')

















