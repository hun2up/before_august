import os
import pandas as pd

df_fa = pd.read_excel(r'D:/project/regist/fa.xlsx')
df_fa.rename(columns={'영업가족CD':'영업가족코드'}, inplace=True)
df_fa = df_fa[['사원번호','영업가족CD','입사일자(사원)','퇴사일자(사원)','유자격(사용인)','금소법교육이수','재직여부']]
df_fa['입사일자(사원)'] = pd.to_datetime(df_fa['입사일자(사원)'])
df_fa['퇴사일자(사원)'] = pd.to_datetime(df_fa['퇴사일자(사원)'])

df_branch = pd.read_excel(r'D:/project/regist/branch.xlsx')
df_branch = df_branch[['영업가족코드','소속부서','영업가족명','코드구분']]
# 데이터 정리 (과정코드)
for modify in range(df_branch.shape[0]):
    if len(df_branch.iloc[modify,2].split("<"))>5:
        df_branch.iloc[modify,1] = df_branch.iloc[modify,2].split("<")[0].replace('(','')
df_merge = pd.merge(df_fa, df_branch, on=['영업가족코드'])

cutoff_date = pd.to_datetime('2023-08-31')

#8월
filtered_august = df_fa[df_fa['입사일자(사원)'] <= cutoff_date]
filtered_august = df_fa[(df_fa['퇴사일자(사원)'] > cutoff_date) & (df_fa['퇴사일자(사원)'] != None)]
#입문과정
#코드등록


