import pandas as pd
from datetime import datetime

# 영업가족직원관리 파일 호출
df_fa = pd.read_csv('/Users/huni/Documents/Python/registered_fa/fa.csv')
# 컬럼명 변경
df_fa.rename(columns={'영업가족CD':'영업가족코드'}, inplace=True)
# 컬럼 재선택
df_fa = df_fa[['사원번호','영업가족코드','입사일자(사원)','퇴사일자(사원)','유자격(사용인)','입문과정수료여부','직무수행교육이수','재직여부']]
# 데이터 삭제 (금소법 미수료)
df_fa = df_fa.drop(df_fa[(df_fa.iloc[:,5] == None) & (df_fa.iloc[:,6] == None)].index)
# 데이터 삭제 (등록여부: 미등록, 손보말소, 생보말소, 손생말소)
# df_fa = df_fa.drop(df_fa[(df_fa.iloc[:,4] != '생보등록') & (df_fa.iloc[:,4] != '손보등록') & (df_fa.iloc[:,4] != '손생등록')].index)
# df_attend: 컬럼 추가 및 데이터 삽입 (입사연차)
df_fa['입사연차'] = (datetime.now().year%100 + 1 - df_fa['사원번호'].astype(str).str[:2].astype(int, errors='ignore')).apply(lambda x: f'{x}년차')
# 입사일자 및 퇴사일자 데이터 시간 형식으로 변경
df_fa['입사일자(사원)'] = pd.to_datetime(df_fa['입사일자(사원)'])
df_fa['퇴사일자(사원)'] = pd.to_datetime(df_fa['퇴사일자(사원)'])

# 영업가족관리 파일 호출
df_branch = pd.read_csv('/Users/huni/Documents/Python/registered_fa/branch.csv')
# 컬럼명 변경 (소속부서 -> 파트너)
df_branch.rename(columns={'소속부서':'파트너'}, inplace=True)
# 컬럼 재선택
df_branch = df_branch[['영업가족코드','파트너','영업가족명','코드구분']]
# 데이터 정리 (소속)
df_branch.insert(loc=1, column='소속부문', value=None)
df_branch.insert(loc=2, column='소속총괄', value=None)
df_branch.insert(loc=3, column='소속부서', value=None)
# 소속 정리
for modify in range(df_branch.shape[0] -1,-1,-1):
    split = df_branch.iloc[modify,4].split(">")
    if len(split) > 4:
        df_branch.iloc[modify,1] = df_branch.iloc[modify,4].split(">")[2]
        df_branch.iloc[modify,2] = df_branch.iloc[modify,4].split(">")[3]
        df_branch.iloc[modify,3] = df_branch.iloc[modify,4].split(">")[4]
    elif split[1] == '다이렉트부문총괄':
        if len(split) > 3:
            df_branch.iloc[modify,1] = df_branch.iloc[modify,4].split(">")[2]
            df_branch.iloc[modify,2] = df_branch.iloc[modify,4].split(">")[3]
            df_branch.iloc[modify,3] = df_branch.iloc[modify,4].split(">")[3]
        else:
            df_branch.iloc[modify,1] = df_branch.iloc[modify,4].split(">")[2]
            df_branch.iloc[modify,2] = df_branch.iloc[modify,4].split(">")[2]
            df_branch.iloc[modify,3] = df_branch.iloc[modify,4].split(">")[2]
    else:
        df_branch.drop([modify], inplace=True)
# 컬럼 삭제
df_branch = df_branch.drop(columns='파트너')
# dataframe 병합
df_merge = pd.merge(df_fa, df_branch, on=['영업가족코드'], how='left')

# 월별로 정리
df_registered = pd.DataFrame([])
dates = [
    '2023-01-31',
    '2023-02-28',
    '2023-03-31',
    '2023-04-30',
    '2023-05-31',
    '2023-06-30',
    '2023-07-31',
    '2023-08-31',
    ]
for register in range(len(dates)):
    # 기준일 생성
    cutoff = pd.to_datetime(dates[register])
    # 기준일 이전 입사자만 추리기
    df_cutoff = df_merge[df_merge['입사일자(사원)'] <= cutoff]
    # 기준일 이전 퇴사자 삭제하기
    df_cutoff = df_cutoff.drop(df_cutoff[df_cutoff.iloc[:,3] <= cutoff].index)
    # 인덱스 재설정
    df_cutoff.reset_index(drop=True, inplace=True)
    # 소속부문별 재적인원
    df_channel = df_cutoff.groupby(['소속부문'])['사원번호'].count().reset_index(name='재적인원')
    # 컬럼명 재설정 (소속부문 -> 항목)
    df_channel.rename(columns={'소속부문':'항목'}, inplace=True)
    # 구분 컬럼 생성
    df_channel['구분'] = '소속부문'
    # 입사연차별 재적인원
    df_career = df_cutoff.groupby(['입사연차'])['사원번호'].count().reset_index(name='재적인원')
    # 컬럼명 재설정 (입사연차 -> 항목)
    df_career.rename(columns={'입사연차':'항목'}, inplace=True)
    # 구분 컬럼 생성
    df_career['구분'] = '입사연차'
    # 소속부문+입사연차
    df_concat = pd.concat([df_channel, df_career], axis=0, ignore_index=True)
    # 월 값 삽입
    df_concat['월'] = f'{register+1}월'
    df_registered = pd.concat([df_registered, df_concat], axis=0, ignore_index=True)

df_registered = df_registered[['월','구분','항목','재적인원']]
df_channel = df_channel.groupby(['구분'])['재적인원'].sum().reset_index(name='합계')
print(df_channel)

# df_registered.to_excel('registered_fa.xlsx', index=False)