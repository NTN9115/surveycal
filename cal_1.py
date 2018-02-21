import pandas as pd
import numpy as np
from pandas import DataFrame, Series
from os import listdir
from datetime import datetime


def readdir():
    filelist = listdir('data')
    df = DataFrame()
    for file in filelist:
        tmp = pd.read_excel('data/' + file)
        df = df.append(tmp)
    a=check(df)

    if a == True:
        if df.columns[3] == '汇报时长 [10min内完成“亮点展示”之前环节的汇报]':
            proc1(df)
        else:
            proc2(df)
    else:
        print('未完成或出现错误')
    


def proc1(df=DataFrame()):
    conf = pd.read_csv('conf.csv')
    total = df.iloc[:,
                    3] * 0.01 + df.iloc[:,
                                        4] * 0.04 + df.iloc[:,
                                                            5] * 0.05 + df.iloc[:,
                                                                                6] * 0.2 + df.iloc[:,
                                                                                                   7] * 0.3 + df.iloc[:,
                                                                                                                      8] * 0.05 + df.iloc[:,
                                                                                                                                          9] * 0.35
    df.insert(11, '综合得分', total)
    list1 = list(conf['公司领导'])
    list1.extend(list(conf['公司主管领导']))
    list1.extend(list(conf['部门领导']))
    df_1 = df[df['名'] == list(conf['公司领导'])[0]].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_2 = df[df['名'] == list(conf['公司领导'])[1]].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_3 = df[df['名'] == list(conf['公司领导'])[2]].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_4 = df[df['名'].isin(list(conf['公司主管领导']))].loc[:,
                                                      ['汇报部门', '汇报人', '综合得分']]
    df_5 = df[~df['名'].isin(list1)].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_1 = df_1.groupby(['汇报部门', '汇报人']).mean() * 0.1
    df_2 = df_2.groupby(['汇报部门', '汇报人']).mean() * 0.1
    df_3 = df_3.groupby(['汇报部门', '汇报人']).mean() * 0.1
    df_4 = df_4.groupby(['汇报部门', '汇报人']).mean() * 0.45
    df_5 = df_5.groupby(['汇报部门', '汇报人']).mean() * 0.25
    df_n = df_1 + df_2 + df_3 + df_4 + df_5
    print(df_n)
    dt = datetime.now()
    df_n.to_excel(dt.strftime('%Y-%m-%d') + '运营分析报告统计.xlsx')


def proc2(df=DataFrame()):
    conf = pd.read_csv('conf.csv')
    total = df.iloc[:,
                    3] * 0.01 + df.iloc[:,
                                        4] * 0.04 + df.iloc[:,
                                                            5] * 0.05 + df.iloc[:,
                                                                                6] * 0.2 + df.iloc[:,
                                                                                                   7] * 0.3 + df.iloc[:,
                                                                                                                      8] * 0.05 + df.iloc[:,
                                                                                                                                          9] * 0.35
    df.insert(11, '综合得分', total)
    list1 = list(conf['公司领导'])
    list1.extend(list(conf['公司主管领导']))
    list1.extend(list(conf['部门领导']))
    df_1 = df[df['名'] == list(conf['公司领导'])[0]].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_2 = df[df['名'] == list(conf['公司领导'])[1]].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_3 = df[df['名'] == list(conf['公司领导'])[2]].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_4 = df[df['名'].isin(list(conf['公司主管领导']))].loc[:,
                                                      ['汇报部门', '汇报人', '综合得分']]
    df_5 = df[df['名'].isin(list(conf['部门领导']))].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_6 = df[~df['名'].isin(list1)].loc[:, ['汇报部门', '汇报人', '综合得分']]
    df_1 = df_1.groupby(['汇报部门', '汇报人']).mean() * 0.1
    df_2 = df_2.groupby(['汇报部门', '汇报人']).mean() * 0.1
    df_3 = df_3.groupby(['汇报部门', '汇报人']).mean() * 0.1
    df_4 = df_4.groupby(['汇报部门', '汇报人']).mean() * 0.30
    df_5 = df_5.groupby(['汇报部门', '汇报人']).mean() * 0.25
    df_6 = df_6.groupby(['汇报部门', '汇报人']).mean() * 0.15
    df_n = df_1 + df_2 + df_3 + df_4 + df_5 + df_6
    print(df_n)
    dt = datetime.now()
    df_n.to_excel(dt.strftime('%Y-%m-%d') + '销售大会述职报告统计.xlsx')

def check(df=DataFrame()):
    if '建议及点评' in df.columns:
        del df['建议及点评']
    df = df.groupby(['汇报部门', '汇报人']).count()
    test=df.iat[0,0]
    test1 = df==test
    test2=test1[test1['提交的日期']==False]
    if test2.empty:
        return True
    else:
        print(test2)
        return False
if __name__ == '__main__':
    readdir()