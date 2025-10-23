import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta, date
import re
import time
from pic_download import download_and_compress_image_plus
from dayly_pic_download import open_protected_excel_safe

# 记录开始时间
start_time = time.time()

# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

CWD = f'D:/桌面/模板/{today}'

p = Path(CWD)

# 获取当天日期
today_small = datetime.now().strftime("%y.%m.%d")
print(today_small)

CWD = f'D:/桌面/过往货盘表/'

p = Path(CWD)
file_path = list(p.glob(f'*{today_small}*.xlsx'))[0]
print(file_path)
df_huo = open_protected_excel_safe(file_path, password="789", sheet_name='Sheet1')

def main(df_huo):
    # 取第二行作为列名（索引1），再去掉前两行
    new_header = df_huo.iloc[1].astype(str).str.strip()
    df_huo = df_huo.iloc[2:].reset_index(drop=True)
    df_huo.columns = new_header

    df_huo = df_huo.iloc[:, 0:100]

    df_huo = df_huo[df_huo['类目'].notna()]
    mask = df_huo['类目'].eq('美妆')
    df_huo.loc[mask, '大类'] = '美妆'

    df_huo['上新日'] = pd.to_datetime(df_huo['上新日'], errors='coerce')
    create_dt = pd.to_datetime(df_huo['创建日期'], errors='coerce')
    missing = df_huo['上新日'].isna()
    df_huo.loc[missing, '上新日'] = create_dt.loc[missing]

    # 将“最后一次下单日期”缺失值改用上新日填充
    df_huo['最后一次下单日期'] = pd.to_datetime(df_huo['最后一次下单日期'], errors='coerce')
    mask_last_order = df_huo['最后一次下单日期'].isna()
    df_huo.loc[mask_last_order, '最后一次下单日期'] = df_huo.loc[mask_last_order, '上新日']
    df_huo['最后一次下单日期'] = df_huo['最后一次下单日期'].dt.date


    today_dt = pd.Timestamp.today().normalize()
    df_huo['今天日期'] = today_dt.date()
    last_order = pd.to_datetime(df_huo['最后一次下单日期'], errors='coerce')
    df_huo['距最后一次下单天数'] = (today_dt - last_order).dt.days

    # 计算总天数差，然后换算成月份（按30天为一个月）
    days_gap = (today_dt - last_order).dt.days
    months_gap = days_gap / 30.44  # 平均每月30.44天
    bins = [-1, 6, 12, 18, 24, np.inf]
    labels = ['0-6个月', '7-12个月', '13-18个月', '19-24个月', '24个月以上']
    df_huo['距最后一次下单区间'] = pd.cut(months_gap, bins=bins, labels=labels)

    df_huo = df_huo[df_huo['可用数'] >= 1]

    df_huo['库存成本金额'] = df_huo['在仓总库存'] * df_huo['成本']


    df_huo_pivot = pd.pivot_table(
        df_huo,
        index=['大类','距最后一次下单区间' ],
        values=['货号+色号','在仓总库存', '库存成本金额'],
        aggfunc={'货号+色号':'count','在仓总库存': 'sum', '库存成本金额': 'sum'},
        ).reset_index()
    
    stage_totals = (
        df_huo_pivot.groupby('大类', as_index=False)[['货号+色号', '在仓总库存', '库存成本金额']]
        .sum()
        .assign(距最后一次下单区间='总计')
    )
    df_huo_pivot = pd.concat([df_huo_pivot, stage_totals], ignore_index=True)

    df_huo_pivot['库存成本金额'] = pd.to_numeric(df_huo_pivot['库存成本金额'], errors='coerce').fillna(0)
    df_huo_pivot['库存成本金额'] = df_huo_pivot['库存成本金额'].round()
    total_cost = df_huo_pivot['库存成本金额'].sum()
    df_huo_pivot['库存成本金额占比'] = 0 if total_cost == 0 else (df_huo_pivot['库存成本金额'] / total_cost).round(4)

    df_huo_pivot['库存成本金额占比'] = df_huo_pivot['库存成本金额占比'] * 2

    totals = {
        '大类': '总计',
        '距最后一次下单区间': '合计',
        '货号+色号': df_huo_pivot['货号+色号'].sum()/2,
        '在仓总库存': df_huo_pivot['在仓总库存'].sum()/2,
        '库存成本金额': df_huo_pivot['库存成本金额'].sum()/2,
        '库存成本金额占比': df_huo_pivot['库存成本金额占比'].sum()/2,
    }
    df_huo_pivot = pd.concat([df_huo_pivot, pd.DataFrame([totals])], ignore_index=True)


    df_huo_pivot.reset_index(drop=True, inplace=True)

    df_huo_pivot = df_huo_pivot.reindex(columns=['大类','距最后一次下单区间',  '货号+色号', '在仓总库存', '库存成本金额', '库存成本金额占比'])
    df_huo_pivot.rename(columns={
        '货号+色号': 'SKC数',
        '距最后一次下单区间':'库龄'
    }, inplace=True)

    primary_order = ['服装', '配饰', '鞋子', '美妆','总计']
    category_order = primary_order + [cat for cat in df_huo_pivot['大类'].unique() if cat not in primary_order]
    stage_order = ['0-6个月', '7-12个月', '13-18个月', '19-24个月', '24个月以上', '总计', '合计']

    df_huo_pivot['大类'] = pd.Categorical(df_huo_pivot['大类'], categories=category_order, ordered=True)
    df_huo_pivot['库龄'] = pd.Categorical(df_huo_pivot['库龄'], categories=stage_order, ordered=True)

    df_huo_pivot = df_huo_pivot.sort_values(['大类', '库龄']).reset_index(drop=True)



    output_file = f'D:/桌面/库龄表/{today}库龄表1.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_huo_pivot.to_excel(writer, index=False, sheet_name='汇总')
        df_huo.to_excel(writer, index=False, sheet_name='明细')

main(df_huo)
