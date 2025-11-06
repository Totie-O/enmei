import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta, date
import re
import time



dir_name = '2025-11-06第一场'

CWD = f'D:/桌面/直播后清单/{dir_name}'

p = Path(CWD)

file_sales_details = list(p.glob('*销售*.xlsx'))
print("file_sales_details:", file_sales_details)

# 当日销售主题
df_sales_today = pd.read_excel(file_sales_details[0]) 

############################################################
# def df_merge(path_ls_all):
#     merged_df = pd.DataFrame()

#     for i in path_ls_all:
#         df = pd.read_excel(i, header = 0)
#         if '款式编号' in df.columns:
#             df = df.rename(columns={'款式编号':'款式编码'}, inplace=True)
        
#         merged_df = pd.concat([merged_df, df], ignore_index = True)


#     return merged_df

# df = df_merge(file_sales_details)



d = {
#     '10-31首场':{'开始时间':'2025-10-31 9:00','结束时间':'2025-10-31 15:20'},
#     '10-31晚场':{'开始时间':'2025-10-31 19:00','结束时间':'2025-11-01 00:26'},
#     '11-03首场':{'开始时间':'2025-11-03 09:00','结束时间':'2025-11-03 14:55'},
#     '11-03晚场':{'开始时间':'2025-11-03 19:00','结束时间':'2025-11-04 00:33'},
    '11-06首场':{'开始时间':'2025-11-06 09:00','结束时间':'2025-11-06 15:30'},
}

def df_solve(df, session_dict=d):
    df['付款日期'] = pd.to_datetime(df['付款日期'], errors='coerce')
    df['日期'] = df['付款日期'].dt.date
    df['场次标签'] = np.nan  # 新列

    df['达人编号'] = df['达人编号'].astype(str).str.strip()
    df = df[df['达人编号'] == '6524430296']


    # 新增小时段列：示例 10.31 17:00-18:00
    df['小时段'] = (
        df['付款日期'].dt.strftime('%m.%d %H:00') + '-' +
        (df['付款日期'] + pd.Timedelta(hours=1)).dt.strftime('%H:00')
    )

    for label, span in session_dict.items():
        start_dt = pd.to_datetime(span['开始时间'])
        end_dt = pd.to_datetime(span['结束时间'])
        mask = (df['付款日期'] >= start_dt) & (df['付款日期'] <= end_dt)
        df.loc[mask, '场次标签'] = label

    df.to_excel(f'{CWD}/直播清单合并.xlsx', index=False)

    # 过滤掉场次标签为空的行
    df = df[df['场次标签'].notna()].copy()
    
    df_solve_pivot = pd.pivot_table(
        df,
        index=['款式编号'],
        columns=['小时段'],
        values='销售数量',
        aggfunc='sum',
        fill_value=0
    )

    # 行总计并按降序排序
    df_solve_pivot['当场累计销量'] = df_solve_pivot.sum(axis=1)
    df_solve_pivot = df_solve_pivot.sort_values('当场累计销量', ascending=False).reset_index()

    df_solve_pivot['图片'] = ''
    df_solve_pivot['品类'] = ''

    hour_cols = [c for c in df_solve_pivot.columns if re.match(r'\d{2}\.\d{2} \d{2}:00-\d{2}:00', c)]
    # 按开始时间排序（月份.日期 小时）

    def _key(s):
        date_part, span = s.split()
        start_hour = span.split('-')[0]
        return datetime.strptime(date_part + ' ' + start_hour, '%m.%d %H:%M')
    hour_cols = sorted(hour_cols, key=_key)

    base_cols = ['款式编号','图片','品类','当场累计销量']
    ordered = base_cols + hour_cols
    # 仅重排存在列
    ordered_existing = [c for c in ordered if c in df_solve_pivot.columns]
    df_solve_pivot = df_solve_pivot[ordered_existing]


    return df_solve_pivot

df_1 = df_solve(df_sales_today)
df_1.to_excel('D:/桌面/测试.xlsx', index=False)
print('输出列顺序:', df_1.columns.tolist())