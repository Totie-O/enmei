import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta, date
import re
import time


# 记录开始时间
start_time = time.time()

# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

CWD = f'D:/桌面/模板/{today}'

p = Path(CWD)

# 固定文件
df_new_category = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='品类')
# df_new_channel = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='渠道')

file_ndim_payment_time = list(p.glob('*销售多维付款时间*.xlsx'))
print("file_ndim_payment_time:", file_ndim_payment_time)

## 销售付款时间
df_ndim_payment_time = pd.read_excel(file_ndim_payment_time[0], usecols=["渠道", "店铺","日期","商品编码","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额"])

def payment_time(df_ndim_payment_time=df_ndim_payment_time):
    df_ndim_payment_time.rename(columns={'渠道':'聚水潭渠道'}, inplace=True)

    df_ndim_payment_time = df_ndim_payment_time.merge(df_new_category[['产品分类','新品类（企划版）']], on='产品分类', how='left')
    # df_ndim_payment_time = df_ndim_payment_time.merge(df_new_channel, on='店铺', how='left')

    # df_ndim_payment_time = pd.to_datetime(df_ndim_payment_time['日期'], errors='coerce')
    df_ndim_payment_time['日期'] = pd.to_datetime(df_ndim_payment_time['日期'], errors='coerce').dt.strftime('%Y-%m-%d')



    #根据分隔符分成两个字段，创建款色
    extracted = df_ndim_payment_time['颜色规格'].str.extract(r'^(.*?);(.*)$')
    df_ndim_payment_time['颜色'] = extracted[0].fillna(df_ndim_payment_time['颜色规格'])  # 无分号时用原值
    df_ndim_payment_time['规格'] = extracted[1]  # 无分号时为NaN
    df_ndim_payment_time['款色'] = df_ndim_payment_time['款式编码'] + df_ndim_payment_time['颜色']

    # 填充缺失值
    df_ndim_payment_time['销售数量'] = df_ndim_payment_time['销售数量'].fillna(0)
    df_ndim_payment_time['退货数量'] = df_ndim_payment_time['退货数量'].fillna(0)
    df_ndim_payment_time['实发数量'] = df_ndim_payment_time['实发数量'].fillna(0)
    df_ndim_payment_time['实退数量'] = df_ndim_payment_time['实退数量'].fillna(0)


    # 透视当天数据
    df_ndim_payment_time_day_pivot = pd.pivot_table(
        df_ndim_payment_time,
        index=['日期', '款色','新品类（企划版）'],
        values=['销售数量','实发数量','实退数量','退货数量'],
        aggfunc='sum'
    ).reset_index()
    
    num_cols = ['销售数量','实发数量','实退数量','退货数量']
    df_ndim_payment_time_day_pivot[num_cols] = (
        df_ndim_payment_time_day_pivot[num_cols]
        .apply(pd.to_numeric, errors='coerce')
        .fillna(0)
    )
    
    # df_ndim_payment_time_day_pivot['仅退款数量'] = df_ndim_payment_time_day_pivot['销售数量'] - df_ndim_payment_time_day_pivot['实发数量']

    df_ndim_payment_time_day_pivot = df_ndim_payment_time_day_pivot.sort_values(['日期','款色'])
    df_ndim_payment_time_day_pivot['累计退货量'] = df_ndim_payment_time_day_pivot.groupby(['款色'])['退货数量'].cumsum()
    df_ndim_payment_time_day_pivot['累计销售量'] = df_ndim_payment_time_day_pivot.groupby(['款色'])['销售数量'].cumsum()
    df_ndim_payment_time_day_pivot['累计实发量'] = df_ndim_payment_time_day_pivot.groupby(['款色'])['实发数量'].cumsum()
    df_ndim_payment_time_day_pivot['累计实退量'] = df_ndim_payment_time_day_pivot.groupby(['款色'])['实退数量'].cumsum()
    # df_ndim_payment_time_day_pivot['累计仅退款数量'] = df_ndim_payment_time_day_pivot.groupby(['日期', '款色'])['仅退款数量'].cumsum()
    df_ndim_payment_time_day_pivot['累计仅退款数量'] = df_ndim_payment_time_day_pivot['累计销售量'] - df_ndim_payment_time_day_pivot['累计实发量']

    df_ndim_payment_time_day_pivot['累计退货率'] = round(df_ndim_payment_time_day_pivot['累计退货量'] / df_ndim_payment_time_day_pivot['累计销售量'], 2)
    df_ndim_payment_time_day_pivot['累计仅退款率'] = round(df_ndim_payment_time_day_pivot['累计仅退款数量'] / df_ndim_payment_time_day_pivot['累计销售量'], 2)
    df_ndim_payment_time_day_pivot['累计实发退货率'] = round(df_ndim_payment_time_day_pivot['累计实退量'] / df_ndim_payment_time_day_pivot['累计实发量'], 2)

    # df_ndim_payment_time_day_pivot.to_excel(f'/Users/totie_o/Desktop/退货率明细1.xlsx', index=False)
    # 合并近两月销售数量
    df_ndim_payment_time_month_pivot = pd.pivot_table(
        df_ndim_payment_time,
        index=['款色'],
        values=['销售数量','实发数量','实退数量','退货数量'],
        aggfunc='sum'
    ).reset_index()

    df_ndim_payment_time_month_pivot.rename(columns={'销售数量':'近两月销售数量','实发数量':'近两月实发数量','实退数量':'近两月实退数量','退货数量':'近两月退货数量'}, inplace=True)
    df_ndim_payment_time_month_pivot['近两月仅退款数量'] = df_ndim_payment_time_month_pivot['近两月销售数量'] - df_ndim_payment_time_month_pivot['近两月实发数量']
    df_ndim_payment_time_month_pivot['近两月仅退款率'] = round(df_ndim_payment_time_month_pivot['近两月仅退款数量'] / df_ndim_payment_time_month_pivot['近两月销售数量'], 2)
    df_ndim_payment_time_month_pivot['近两月整体退货率'] = round(df_ndim_payment_time_month_pivot['近两月退货数量'] / df_ndim_payment_time_month_pivot['近两月销售数量'], 2)
    df_ndim_payment_time_month_pivot['近两月实发退货率'] = round(df_ndim_payment_time_month_pivot['近两月实退数量'] / df_ndim_payment_time_month_pivot['近两月实发数量'], 2)

    df_ndim_payment_time_day_pivot = df_ndim_payment_time_day_pivot.merge(df_ndim_payment_time_month_pivot, on='款色', how='left')

    df_ndim_payment_time_day_pivot = df_ndim_payment_time_day_pivot[['日期','款色','新品类（企划版）','累计退货率','近两月退货数量','累计仅退款率','累计实发退货率','近两月整体退货率', '近两月仅退款率','近两月实发退货率','近两月销售数量','近两月实发数量','近两月实退数量','近两月仅退款数量']]
    df_ndim_payment_time_day_pivot['图片'] = ''

    df_ndim_payment_time_day_pivot = df_ndim_payment_time_day_pivot.set_index(['日期','款色','图片','新品类（企划版）','近两月销售数量','近两月退货数量','近两月整体退货率', '近两月仅退款数量','近两月仅退款率','近两月实发数量','近两月实退数量','近两月实发退货率']).sort_index()

    df_ndim_payment_time_day_pivot = df_ndim_payment_time_day_pivot.unstack('日期')
    df_ndim_payment_time_day_pivot = df_ndim_payment_time_day_pivot.stack(0)
    df_ndim_payment_time_day_pivot.index = df_ndim_payment_time_day_pivot.index.set_names(
        list(df_ndim_payment_time_day_pivot.index.names[:-1]) + ['指标']
    )
    df_ndim_payment_time_day_pivot = df_ndim_payment_time_day_pivot.reset_index()

    return df_ndim_payment_time_day_pivot

df = payment_time()
df.to_excel(f'D:/桌面/总退款率.xlsx', index=False)