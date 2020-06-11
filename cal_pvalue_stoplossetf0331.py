import pandas as pd
import statsmodels.api as sm
from stepwise_regression import step_reg
import numpy as np
import math

init_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/HS300/0331/stoplossvol数据整理.xlsx', sheet_name=0).iloc[:, :]
#变量太多，不用创业板指
df_final = pd.DataFrame(index = init_df.index,columns = ['etfcoe','highetf0','highetf1','highetf2','highetf3',
        'lowetf0','lowetf1','lowetf2','lowetf3','pctetf1','pctetf2','pctetf3','highsh0','highsh1','highsh2','highsh3',
        'lowsh0','lowsh1','lowsh2','lowsh3','pctsh1','pctsh2','pctsh3','highsz0','highsz1','highsz2','highsz3',
        'lowsz0','lowsz1','lowsz2','lowsz3','pctsz1','pctsz2','pctsz3','highhsi0','highhsi1','highhsi2','highhsi3',
        'lowhsi0','lowhsi1','lowhsi2','lowhsi3','pcthsi1','pcthsi2','pcthsi3','highdji0','highdji1','highdji2',
        'highdji3','lowdji0','lowdji1','lowdji2','lowdji3','pctdji1','pctdji2','pctdji3','highixic0','highixic1',
        'highixic2','highixic3','lowixic0','lowixic1','lowixic2','lowixic3','pctixic1','pctixic2','pctixic3',
        'highspx0','highspx1','highspx2','highspx3','lowspx0','lowspx1','lowspx2','lowspx3','pctspx1','pctspx2',
        'pctspx3','hvoletf1','hvoletf2','hvoletf3','lvoletf1','lvoletf2','lvoletf3','cvoletf1','cvoletf2','cvoletf3',
        'hvolsh1','hvolsh2','hvolsh3','lvolsh1','lvolsh2','lvolsh3','cvolsh1','cvolsh2','cvolsh3','hvolsz1','hvolsz2',
        'hvolsz3','lvolsz1','lvolsz2','lvolsz3','cvolsz1','cvolsz2','cvolsz3','hvolhsi1','hvolhsi2','hvolhsi3',
        'lvolhsi1','lvolhsi2','lvolhsi3','cvolhsi1','cvolhsi2','cvolhsi3','hvoldji1','hvoldji2','hvoldji3',
        'lvoldji1','lvoldji2','lvoldji3','cvoldji1','cvoldji2','cvoldji3','hvolixic1','hvolixic2','hvolixic3',
        'lvolixic1','lvolixic2','lvolixic3','cvolixic1','cvolixic2','cvolixic3','hvolspx1','hvolspx2','hvolspx3',
        'lvolspx1','lvolspx2','lvolspx3','cvolspx1','cvolspx2','cvolspx3'])


df_final.fillna(0,inplace =True)
df_final['etfcoe'] = init_df.iloc[:,list(init_df.columns).index('etfcoe')].values

for i in ['etf','sh','sz','hsi','dji','ixic','spx']:
    df_final['high'+i+'0'] = init_df.iloc[:, list(init_df.columns).index('high' + i +'a')].values / init_df.iloc[
                            :, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[1:,'high'+i+'1'] = init_df.iloc[:-1, list(init_df.columns).index('high'+i+'a')].values / init_df.iloc[
                            1:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[2:,'high'+i+'2'] = init_df.iloc[:-2, list(init_df.columns).index('high'+i+'a')].values / init_df.iloc[
                            2:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[3:,'high'+i+'3'] = init_df.iloc[:-3, list(init_df.columns).index('high'+i+'a')].values / init_df.iloc[
                            3:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final['low'+i+'0'] = init_df.iloc[:, list(init_df.columns).index('low' + i +'a')].values / init_df.iloc[
                            :, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[1:,'low'+i+'1'] = init_df.iloc[:-1, list(init_df.columns).index('low'+i+'a')].values / init_df.iloc[
                            1:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[2:,'low'+i+'2'] = init_df.iloc[:-2, list(init_df.columns).index('low'+i+'a')].values / init_df.iloc[
                            2:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[3:,'low'+i+'3'] = init_df.iloc[:-3, list(init_df.columns).index('low'+i+'a')].values / init_df.iloc[
                            3:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[1:,'pct'+i+'1'] = init_df.iloc[:-1, list(init_df.columns).index('pct'+i+'a')].values / init_df.iloc[
                            1:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[2:,'pct'+i+'2'] = init_df.iloc[:-2, list(init_df.columns).index('pct'+i+'a')].values / init_df.iloc[
                            2:, list(init_df.columns).index('pct'+ i + 'a')].values
    df_final.loc[3:,'pct'+i+'3'] = init_df.iloc[:-3, list(init_df.columns).index('pct'+i+'a')].values / init_df.iloc[
                            3:, list(init_df.columns).index('pct'+ i + 'a')].values

    df_final.loc[1:,'hvol' + i + '1'] = init_df.iloc[:-1, list(init_df.columns).index(
        'hvol' + i + 'a')].values / init_df.iloc[1:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[2:,'hvol' + i + '2'] = init_df.iloc[:-2, list(init_df.columns).index(
        'hvol' + i + 'a')].values / init_df.iloc[2:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[3:,'hvol' + i + '3'] = init_df.iloc[:-3, list(init_df.columns).index(
        'hvol' + i + 'a')].values / init_df.iloc[3:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[1:,'lvol' + i + '1'] = init_df.iloc[:-1, list(init_df.columns).index(
        'lvol' + i + 'a')].values / init_df.iloc[1:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[2:,'lvol' + i + '2'] = init_df.iloc[:-2, list(init_df.columns).index(
        'lvol' + i + 'a')].values / init_df.iloc[2:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[3:,'lvol' + i + '3'] = init_df.iloc[:-3, list(init_df.columns).index(
        'lvol' + i + 'a')].values / init_df.iloc[3:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[1:,'cvol' + i + '1'] = init_df.iloc[:-1, list(init_df.columns).index(
        'cvol' + i + 'a')].values / init_df.iloc[1:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[2:,'cvol' + i + '2'] = init_df.iloc[:-2, list(init_df.columns).index(
        'cvol' + i + 'a')].values / init_df.iloc[2:, list(init_df.columns).index('cvol' + i + 'a')].values
    df_final.loc[3:,'cvol' + i + '3'] = init_df.iloc[:-3, list(init_df.columns).index(
        'cvol' + i + 'a')].values / init_df.iloc[3:, list(init_df.columns).index('cvol' + i + 'a')].values

df_final.iloc[:,1:] = df_final.iloc[:,1:].applymap(lambda y:math.log(y) if y != 0 else y)
df_final.to_csv(r'C:/Users/Tao/Desktop/投资/HS300/stoplossvol检查用.csv', index=False, encoding='utf_8_sig')
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

finallist =['etfcoe']
if 'const' in backward_regression(data,y_value,0.05,verbose = False)[:]:
    finallist.extend(backward_regression(data, y_value, 0.05, verbose=False)[:])
    finallist.remove('const')
    data = df_final[finallist]
    y_value = data.iloc[3:, 0]
    data = data.iloc[3:, 1:]
    data = sm.add_constant(data)
else:
    finallist.extend(backward_regression(data, y_value, 0.05, verbose=False)[:])
    data = df_final[finallist]
    i_value = list((data == 0).any(axis=1)).index(False)
    y_value = data.iloc[3:, 0]
    data = data.iloc[3:, 1:]

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
result_df.to_csv(r'C:/Users/Tao/Desktop/投资/HS300/stoplossvolresult.csv', index=False, encoding='utf_8_sig')







