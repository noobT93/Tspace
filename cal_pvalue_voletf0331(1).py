import pandas as pd
import statsmodels.api as sm
from stepwise_regression import step_reg
import numpy as np
import math

init_df = pd.read_excel(r'C:/Users/Tao/Desktop/投资/HS300/0331/vol数据整理 - 副本.xlsx', sheet_name=0).iloc[:, :]
#变量太多，不用创业板指
df_final = pd.DataFrame(index = init_df.index,columns = ['etfcoe',
        'openvaretf1','openvaretf2','openvaretf3','maxvaretf1','maxvaretf2','maxvaretf3',
        'openvarsh1','openvarsh2','openvarsh3','maxvarsh1','maxvarsh2','maxvarsh3',
        'openvarsz1','openvarsz2','openvarsz3','maxvarsz1','maxvarsz2','maxvarsz3',
        'openvarcy1', 'openvarcy2', 'openvarcy3', 'maxvarcy1','maxvarcy2', 'maxvarcy3',
        'openvarhsi1','openvarhsi2','openvarhsi3','maxvarhsi1','maxvarhsi2','maxvarhsi3',
        'openvardji1','openvardji2','openvardji3','maxvardji1','maxvardji2','maxvardji3',
        'openvarixic1','openvarixic2','openvarixic3','maxvarixic1','maxvarixic2','maxvarixic3',
        'openvarspx1','openvarspx2','openvarspx3','maxvarspx1','maxvarspx2','maxvarspx3',
        'openvoletf1','openvoletf2','openvoletf3','maxvoletf1','maxvoletf2','maxvoletf3',
        'openvolsh1','openvolsh2','openvolsh3','maxvolsh1','maxvolsh2','maxvolsh3',
        'openvolsz1','openvolsz2','openvolsz3','maxvolsz1','maxvolsz2','maxvolsz3',
        'openvolcy1', 'openvolcy2', 'openvolcy3', 'maxvolcy1','maxvolcy2', 'maxvolcy3',
        'openvolhsi1','openvolhsi2','openvolhsi3','maxvolhsi1','maxvolhsi2','maxvolhsi3',
        'openvoldji1','openvoldji2','openvoldji3','maxvoldji1','maxvoldji2','maxvoldji3',
        'openvolixic1','openvolixic2','openvolixic3','maxvolixic1','maxvolixic2','maxvolixic3',
        'openvolspx1','openvolspx2','openvolspx3','maxvolspx1','maxvolspx2','maxvolspx3',
        'highetf', 'lowetf', 'highsh', 'lowsh', 'highsz', 'lowsz','highcy', 'lowcy',
        'highdji', 'lowdji','highixic', 'lowixic', 'highspx', 'lowspx','highhsi', 'lowhsi'])


df_final.fillna(0,inplace =True)
df_final['etfcoe'] = init_df.iloc[:,list(init_df.columns).index('etfcoe')].values

for i in ['etf','sh','sz','cy','hsi','dji','ixic','spx']:
    df_final.loc[1:, 'high' + i ] = [abs(y-1) for y in list(init_df.loc[
    1:, 'high'+i+'a'].values / init_df.loc[:len(init_df.index)-2, 'pct'+ i + 'a'].values)]
    df_final.loc[1:, 'low' + i ] = [abs(y-1) for y in list(init_df.loc[
    1:, 'low'+i+'a'].values / init_df.loc[:len(init_df.index)-2, 'pct'+ i + 'a'].values)]

    df_final.loc[1:,'openvar'+i+'1'] = [abs(y-1) for y in list(init_df.loc[
    1:, 'open'+i+'a'].values / init_df.loc[:len(init_df.index)-2, 'pct'+ i + 'a'].values)]
    df_final.loc[2:,'openvar'+i+'2'] = df_final.loc[1:,'openvar'+i+'1'].values[:-1]
    df_final.loc[3:,'openvar'+i+'3'] = df_final.loc[1:,'openvar'+i+'1'].values[:-2]

    df_final.loc[1:,'maxvar'+i+'1'] = df_final.loc[1:,['high' + i ,'low' + i ]].max(axis =1)
    df_final.loc[2:,'maxvar'+i+'2'] = df_final.loc[1:,'maxvar'+i+'1'].values[:-1]
    df_final.loc[3:,'maxvar'+i+'3'] = df_final.loc[1:,'maxvar'+i+'1'].values[:-2]

    df_final.loc[1:, 'openvol' + i + '1'] = init_df.loc[
    1:, 'cvol' + i + 'a'].values / init_df.loc[:len(init_df.index) - 2,'cvol' + i + 'a'].values
    df_final.loc[2:, 'openvol' + i + '2'] = df_final.loc[1:, 'openvol' + i + '1'].values[:-1]
    df_final.loc[3:, 'openvol' + i + '3'] = df_final.loc[1:, 'openvol' + i + '1'].values[:-2]
    for yy in df_final.index[1:]:
        if df_final.loc[yy, 'maxvar' + i + '1'] == df_final.loc[yy, 'high' + i]:
            df_final.loc[yy, 'maxvol' + i + '1'] = init_df.loc[
            yy, 'hvol' + i + 'a'] / init_df.loc[yy - 1, 'cvol' + i + 'a']
        else:
            df_final.loc[yy, 'maxvol' + i + '1'] = init_df.loc[
            yy, 'lvol' + i + 'a'] / init_df.loc[yy - 1, 'cvol' + i + 'a']
    df_final.loc[2:, 'maxvol' + i + '2'] = df_final.loc[1:, 'maxvol' + i + '1'].values[:-1]
    df_final.loc[3:, 'maxvol' + i + '3'] = df_final.loc[1:, 'maxvol' + i + '1'].values[:-2]


df_final.iloc[:,1:] = df_final.iloc[:,1:].applymap(lambda y:math.log(y) if y != 0 else y)
df_final.drop(['highetf', 'lowetf', 'highsh', 'lowsh', 'highsz', 'lowsz','highcy', 'lowcy',
        'highdji', 'lowdji','highixic', 'lowixic', 'highspx', 'lowspx','highhsi', 'lowhsi'],axis=1,inplace = True)

df_final.to_csv(r'C:/Users/Tao/Desktop/投资/HS300/vol检查用111.csv', index=False, encoding='utf_8_sig')

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
result_df.to_csv(r'C:/Users/Tao/Desktop/投资/HS300/volresult111.csv', index=False, encoding='utf_8_sig')







