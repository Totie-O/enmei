import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import time

# 记录开始时间
start_time = time.time()

# 固定文件
df_old_report  = pd.read_excel(r'D:/桌面/模板/固定文件/22年-24年销售记录.xlsx', sheet_name='总表')

# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

CWD = f'D:/桌面/模板/{today}'

p = Path(CWD)

file_report = list(p.glob('*销售记录*.xlsx'))

print("file_file_report:", file_report)

df_report = pd.read_excel(file_report[0])


def report(df_old_report=df_old_report, df_report=df_report):
    df_report = pd.concat([df_old_report, df_report], ignore_index=True)
    #根据分隔符分成两个字段
    extracted = df_report['颜色规格'].str.extract(r'^(.*?);(.*)$')
    df_report['颜色'] = extracted[0].fillna(df_report['颜色规格'])  # 无分号时用原值
    df_report['规格'] = extracted[1]  # 无分号时为NaN

    df_report['款色'] = df_report['款式编码'] + df_report['颜色']

    fill_cols = ['销售数量', '退货数量']
    for col in fill_cols:
        df_report[col] = pd.to_numeric(df_report[col], errors='coerce').fillna(0)

    df_report['净销量'] = df_report['销售数量'] - df_report['退货数量']

    df_report_pivot = pd.pivot_table(
        df_report,
        index=['款色', '款式编码', '颜色'],
        values=['销售数量', '退货数量','净销量'],
        aggfunc='sum'
    ).reset_index()

    return df_report_pivot

df_report_pivot = report()

df_report_pivot.to_clipboard(index=False)



