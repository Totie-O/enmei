import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta, date
import re
import time

df_daren = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='达人')

# 记录开始时间
start_time = time.time()

# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

CWD = f'D:/桌面/模板/{today}'

p = Path(CWD)

file_today = list(p.glob('*当日销售*.xlsx'))
file_sales_details = list(p.glob('*销售明细*.xlsx'))

print("file_sales_details:", file_sales_details)
print("file_today:", file_today)

## 销售明细
df_sales_details = pd.read_excel(file_sales_details[0], usecols=[ "内部订单号","店铺","付款日期","商品编码","达人编号","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额"])

## 当日销售
df_today_sales = pd.read_excel(file_today[0])

def today_sales(df_today_sales=df_today_sales):
    df_today_sales['销售数量'] = df_today_sales['销售数量'].fillna(0)
    df_today_sales['销售金额'] = df_today_sales['销售金额'].fillna(0)

    df_today_sales_pivot = pd.pivot_table(
        df_today_sales,
        index=['款式编号','达人编号'],
        values=['销售数量','销售金额'],
        aggfunc='sum'
    ).reset_index()

    df_today_sales_pivot.rename(columns={'款式编号': '款式编码'}, inplace=True)

    return df_today_sales_pivot

def sales_details(df_sales_details=df_sales_details):
    # 获取当前日期和时间
    now = datetime.now()

    # 计算前天日期
    day_before_yesterday = now.date() - timedelta(days=2)

    # 计算昨天日期
    yesterday = now.date() - timedelta(days=1)

    # 创建昨天8:30的时间点
    yesterday_830 = datetime.combine(yesterday, datetime.strptime('08:30:00', '%H:%M:%S').time())

    # 创建昨天结束的时间点（今天午夜24:00）
    yesterday_end = datetime.combine(yesterday, datetime.strptime('23:59:59', '%H:%M:%S').time())

    # 先将付款日期列转为datetime（无法解析的设为 NaT），再做区间过滤
    if '付款日期' in df_sales_details.columns:
        df_sales_details['付款日期'] = pd.to_datetime(df_sales_details['付款日期'], errors='coerce')

        # 筛选指定日期的订单
        mask1 = df_sales_details['付款日期'].dt.date == day_before_yesterday
        df_sales_details1 = df_sales_details[mask1].copy()

        # 使用 between 做区间筛选，NaT 会被视为 False
        mask = df_sales_details['付款日期'].between(yesterday_830, yesterday_end)
        df_sales_details = df_sales_details[mask].copy()

        # df_sales_details = df_sales_details[df_sales_details['付款日期'].dt.date == yesterday]
        # 如果后续只需要日期部分，可以转换为 date
        # df_sales_details['付款日期'] = df_sales_details['付款日期'].dt.date
    else:
        # 如果没有该列，返回空表
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_sales_details['销售数量'] = df_sales_details['销售数量'].fillna(0)
    df_sales_details['销售金额'] = df_sales_details['销售金额'].fillna(0)

    df_sales_details_pivot = pd.pivot_table(
        df_sales_details,
        index=['款式编码','达人编号'],
        values=['销售数量','销售金额'],
        aggfunc='sum'
    ).reset_index()

    return df_sales_details, df_sales_details_pivot, df_sales_details1

df_today = today_sales(df_today_sales)

df_details, df_details_pivot, df_sales_details1 = sales_details(df_sales_details)

df_concat = pd.concat([df_today, df_details_pivot], axis=0, ignore_index=True)

df_concat_pviot = pd.pivot_table(
    df_concat,
    index=['款式编码','达人编号'],
    values=['销售数量','销售金额'],
    aggfunc='sum'
).reset_index()


# 统一达 人 编 号的数据类型，避免合并时报类型不匹配错误
if '达人编号' in df_concat_pviot.columns:
    df_concat_pviot['达人编号'] = df_concat_pviot['达人编号'].astype(str).str.strip()
if '达人编号' in df_daren.columns:
    df_daren['达人编号'] = df_daren['达人编号'].astype(str).str.strip()

df_concat_pviot = df_concat_pviot.merge(df_daren, on='达人编号', how='left')


with pd.ExcelWriter(f'D:/桌面/达人数据更新/{today}达人销售数据.xlsx', engine='openpyxl') as writer:
    df_details.to_excel(writer, sheet_name='昨天销售明细', index=False)
    df_today.to_excel(writer, sheet_name='当日销售明细', index=False)
    df_concat_pviot.to_excel(writer, sheet_name='红人昨天销售', index=False)
    df_sales_details1.to_excel(writer, sheet_name='前天销售明细', index=False)

# 记录结束时间并计算耗时
end_time = time.time()
run_time = end_time - start_time

print(f"程序运行了 {run_time:.2f} 秒")