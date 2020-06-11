import pandas as pd
import statsmodels.api as sm
from stepwise_regression import step_reg
import numpy as np
import math

price_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/0331/stoploss数据整理（以前一日收盘价为底）.xlsx', sheet_name=0).iloc[:, :]
vol_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/spy/0331/stoploss数据整理（以前一日收盘价为底）.xlsx', sheet_name=1).iloc[:, :]
df_final = pd.DataFrame(index = price_df.index,columns = ['spycoe','highspy0','highspy1','highspy2',
                'highspy3','lowspy0','lowspy1','lowspy2','lowspy3','pctspy1','pctspy2','pctspy3','highvix0',
                'highvix1','highvix2','highvix3','lowvix0','lowvix1','lowvix2','lowvix3','pctvix1','pctvix2','pctvix3',
                'highdji0','highdji1','highdji2','highdji3','lowdji0','lowdji1','lowdji2','lowdji3','pctdji1',
                'pctdji2','pctdji3','highixic0','highixic1','highixic2','highixic3','lowixic0','lowixic1',
                'lowixic2','lowixic3','pctixic1','pctixic2','pctixic3','highspx0','highspx1','highspx2',
                'highspx3','lowspx0','lowspx1','lowspx2','lowspx3','pctspx1','pctspx2','pctspx3',
                'highhsi0','highhsi1','highhsi2','highhsi3','lowhsi0','lowhsi1','lowhsi2','lowhsi3','pcthsi1',
                'pcthsi2','pcthsi3','highsh0','highsh1','highsh2','highsh3','lowsh0','lowsh1','lowsh2',
                'lowsh3','pctsh1','pctsh2','pctsh3','highvolspy1','highvolspy2','highvolspy3','lowvolspy1',
                'lowvolspy2','lowvoluspy3','pctvolspy1','pctvolspy2','pctvolspy3','highvoldji1',
                'highvoldji2','highvoldji3','lowvoldji1','lowvoldji2','lowvoldji3','pctvoldji1','pctvoldji2',
                'pctvoldji3','highvolixic1','highvolixic2','highvolixic3','lowvolixic1','lowvolixic2','lowvolixic3',
                'pctvolixic1','pctvolixic2','pctvolixic3','highvolspx1','highvolspx2','highvolspx3','lowvolspx1',
                'lowvolspx2','lowvolspx3','pctvolspx1','pctvolspx2','pctvolspx3','highvolhsi1','highvolhsi2',
                'highvolhsi3','lowvolhsi1','lowvolhsi2','lowvolhsi3','pctvolhsi1','pctvolhsi2','pctvolhsi3',
                'highvolsh1','highvolsh2','highvolsh3','lowvolsh1','lowvolsh2','lowvolsh3','pctvolsh1','pctvolsh2',
                'pctvolsh3'])
df_final.fillna(0,inplace =True)
df_final['spycoe'] = price_df.iloc[:,list(price_df.columns).index('spycoe')].values
#美股，只有本周四前数据
for i in ['spy','vix','dji','ixic','spx']:
    df_final['high'+i+'0'] = price_df.iloc[:, list(price_df.columns).index('high' + i +'a')].values / price_df.iloc[
                            :, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[1:,'high'+i+'1'] = price_df.iloc[:-1, list(price_df.columns).index('high'+i+'b')].values / price_df.iloc[
                            1:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[2:,'high'+i+'2'] = price_df.iloc[:-2, list(price_df.columns).index('high'+i+'b')].values / price_df.iloc[
                            2:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[3:,'high'+i+'3'] = price_df.iloc[:-3, list(price_df.columns).index('high'+i+'b')].values / price_df.iloc[
                            3:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final['low'+i+'0'] = price_df.iloc[:, list(price_df.columns).index('low' + i +'a')].values / price_df.iloc[
                            :, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[1:,'low'+i+'1'] = price_df.iloc[:-1, list(price_df.columns).index('low'+i+'b')].values / price_df.iloc[
                            1:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[2:,'low'+i+'2'] = price_df.iloc[:-2, list(price_df.columns).index('low'+i+'b')].values / price_df.iloc[
                            2:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[3:,'low'+i+'3'] = price_df.iloc[:-3, list(price_df.columns).index('low'+i+'b')].values / price_df.iloc[
                            3:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[1:,'pct'+i+'1'] = price_df.iloc[:-1, list(price_df.columns).index('pct'+i+'b')].values / price_df.iloc[
                            1:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[2:,'pct'+i+'2'] = price_df.iloc[:-2, list(price_df.columns).index('pct'+i+'b')].values / price_df.iloc[
                            2:, list(price_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[3:,'pct'+i+'3'] = price_df.iloc[:-3, list(price_df.columns).index('pct'+i+'b')].values / price_df.iloc[
                            3:, list(price_df.columns).index('pct'+ i + 'a')].values

    if i != 'vix':
        df_final.loc[1:,'highvol' + i + '1'] = vol_df.iloc[:-1,list(vol_df.columns).index(
            'hvol' + i + '1')].values / vol_df.iloc[1:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[2:,'highvol' + i + '2'] = vol_df.iloc[:-2,list(vol_df.columns).index(
            'hvol' + i + '1')].values / vol_df.iloc[2:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[3:,'highvol' + i + '3'] = vol_df.iloc[:-3,list(vol_df.columns).index(
            'hvol' + i + '1')].values / vol_df.iloc[3:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[1:,'lowvol' + i + '1'] = vol_df.iloc[:-1,list(vol_df.columns).index(
            'lvol' + i + '1')].values / vol_df.iloc[1:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[2:,'lowvol' + i + '2'] = vol_df.iloc[:-2,list(vol_df.columns).index(
            'lvol' + i + '1')].values / vol_df.iloc[2:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[3:,'lowvol' + i + '3'] = vol_df.iloc[:-3,list(vol_df.columns).index(
            'lvol' + i + '1')].values / vol_df.iloc[3:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[1:,'pctvol' + i + '1'] = vol_df.iloc[:-1,list(vol_df.columns).index(
            'cvol' + i + '1')].values / vol_df.iloc[1:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[2:,'pctvol' + i + '2'] = vol_df.iloc[:-2,list(vol_df.columns).index(
            'cvol' + i + '1')].values / vol_df.iloc[2:, list(vol_df.columns).index('cvol' + i + '0')].values
        df_final.loc[3:,'pctvol' + i + '3'] = vol_df.iloc[:-3,list(vol_df.columns).index(
            'cvol' + i + '1')].values / vol_df.iloc[3:, list(vol_df.columns).index('cvol' + i + '0')].values

for i in ['hsi','sh']:
    df_final['high'+i+'0'] = price_df.iloc[:, list(price_df.columns).index('high' + i +'b')].values / price_df.iloc[
                            :, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[1:,'high'+i+'1'] = price_df.iloc[:-1, list(price_df.columns).index('high'+i+'b')].values / price_df.iloc[
                            1:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[2:,'high'+i+'2'] = price_df.iloc[:-2, list(price_df.columns).index('high'+i+'b')].values / price_df.iloc[
                            2:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[3:,'high'+i+'3'] = price_df.iloc[:-3, list(price_df.columns).index('high'+i+'b')].values / price_df.iloc[
                            3:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final['low'+i+'0'] = price_df.iloc[:, list(price_df.columns).index('low' + i +'b')].values / price_df.iloc[
                            :, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[1:,'low'+i+'1'] = price_df.iloc[:-1, list(price_df.columns).index('low'+i+'b')].values / price_df.iloc[
                            1:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[2:,'low'+i+'2'] = price_df.iloc[:-2, list(price_df.columns).index('low'+i+'b')].values / price_df.iloc[
                            2:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[3:,'low'+i+'3'] = price_df.iloc[:-3, list(price_df.columns).index('low'+i+'b')].values / price_df.iloc[
                            3:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[1:,'pct'+i+'1'] = price_df.iloc[:-1, list(price_df.columns).index('pct'+i+'b')].values / price_df.iloc[
                            1:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[2:,'pct'+i+'2'] = price_df.iloc[:-2, list(price_df.columns).index('pct'+i+'b')].values / price_df.iloc[
                            2:, list(price_df.columns).index('pct'+ i + 'b')].values
    df_final.loc[3:,'pct'+i+'3'] = price_df.iloc[:-3, list(price_df.columns).index('pct'+i+'b')].values / price_df.iloc[
                            3:, list(price_df.columns).index('pct'+ i + 'b')].values

    df_final.loc[1:,'highvol' + i + '1'] = vol_df.iloc[:-1, list(vol_df.columns).index(
        'hvol' + i + '1')].values / vol_df.iloc[1:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[2:,'highvol' + i + '2'] = vol_df.iloc[:-2, list(vol_df.columns).index(
        'hvol' + i + '1')].values / vol_df.iloc[2:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[3:,'highvol' + i + '3'] = vol_df.iloc[:-3, list(vol_df.columns).index(
        'hvol' + i + '1')].values / vol_df.iloc[3:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[1:,'lowvol' + i + '1'] = vol_df.iloc[:-1, list(vol_df.columns).index(
        'lvol' + i + '1')].values / vol_df.iloc[1:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[2:,'lowvol' + i + '2'] = vol_df.iloc[:-2, list(vol_df.columns).index(
        'lvol' + i + '1')].values / vol_df.iloc[2:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[3:,'lowvol' + i + '3'] = vol_df.iloc[:-3, list(vol_df.columns).index(
        'lvol' + i + '1')].values / vol_df.iloc[3:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[1:,'pctvol' + i + '1'] = vol_df.iloc[:-1, list(vol_df.columns).index(
        'cvol' + i + '1')].values / vol_df.iloc[1:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[2:,'pctvol' + i + '2'] = vol_df.iloc[:-2, list(vol_df.columns).index(
        'cvol' + i + '1')].values / vol_df.iloc[2:, list(vol_df.columns).index('cvol' + i + '1')].values
    df_final.loc[3:,'pctvol' + i + '3'] = vol_df.iloc[:-3, list(vol_df.columns).index(
        'cvol' + i + '1')].values / vol_df.iloc[3:, list(vol_df.columns).index('cvol' + i + '1')].values

df_final.iloc[:,1:] = df_final.iloc[:,1:].applymap(lambda y:math.log(y) if y != 0 else y)
df_final.to_csv(r'C:/Users/Tao/Desktop/投资/spy/stoploss检查用.csv', index=False, encoding='utf_8_sig')

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
result_df.to_csv(r'C:/Users/Tao/Desktop/投资/spy/stoplossresult.csv', index=False, encoding='utf_8_sig')

















