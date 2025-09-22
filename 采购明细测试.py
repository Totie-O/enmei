import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import time

df_old_procurement  = pd.read_excel(r'D:/桌面/模板/固定文件/23年-24年采购明细.xlsx', sheet_name='Sheet1')

# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

CWD = f'D:/桌面/模板/{today}'

p = Path(CWD)
file_psi = list(p.glob('*库存视角*.xlsx'))
file_product = [i for i in list(p.glob('*商品资料*.xlsx')) if '库存视角' not in str(i)]

file_caigou = list(p.glob('*采购*.xlsx'))
file_sales_details = list(p.glob('*销售明细*.xlsx'))

file_ndim_payment_time = list(p.glob('*付款时间*.xlsx'))
file_ndim_deliver_time = list(p.glob('*发货时间*.xlsx'))

file_today = list(p.glob('*当日销售*.xlsx'))


print("file_today:", file_today)
print("file_psi:", file_psi)
print("file_product:", file_product)
print("file_caigou:", file_caigou)
print("file_sales_details:", file_sales_details)
print("file_ndim_payment_time:", file_ndim_payment_time)
print("file_ndim_deliver_time:", file_ndim_deliver_time)



df_caigou = pd.read_excel(file_caigou[0]) 






def CaiGou(df_caigou=df_caigou, df_old_procurement=df_old_procurement):
    df_caigou = pd.concat([df_old_procurement, df_caigou], ignore_index=True)

    second_column_name = df_caigou.columns[1]
    df_caigou = df_caigou.drop(columns=second_column_name)

    df_caigou[['颜色', '规格']] = df_caigou['颜色规格'].str.split(';', n=1, expand=True)
    df_caigou['规格'] = df_caigou['规格'].fillna('')
    df_caigou['款色'] = df_caigou['款式编码'] + df_caigou['颜色']


    df_caigou = df_caigou[df_caigou['数据类型'] == '明细']
    df_caigou = df_caigou[df_caigou['状态'].isin(['完成', '已确认'])]
    # 采购入库数量透视
    df_caigou_pivot = pd.pivot_table(
        df_caigou,
        index=['款色'],
        values=['采购数量'],
        aggfunc='sum'
    ).reset_index()


    df_caigou = df_caigou[df_caigou['标记|多标签'] != '返修退货']
    
    df_caigou['剔除行'] = df_caigou['备注'].str.contains(r'次品|返修', na=False).map({True: '是', False: '否'})
    df_caigou = df_caigou[df_caigou['剔除行'] != '是']



    # 1. 将采购日期转换为日期格式
    df_caigou['采购日期'] = pd.to_datetime(df_caigou['采购日期'])

    # 2. 按照款色、采购日期、采购供应商分组，并对采购数量进行汇总
    # df_grouped = df_caigou.groupby(['款色', '采购日期', '采购供应商'], as_index=False)['采购数量'].sum()
    df_grouped = df_caigou.groupby(['款色', '采购日期', '采购供应商'], as_index=False).agg({
        '采购数量': 'sum',
        '总入库量': 'sum'  # 假设你的数据中有入库数量列
    })


    # 3. 按照款色升序、日期降序排序
    df_sorted = df_grouped.sort_values(['款色', '采购日期'], ascending=[True, False])

    # 重置索引（可选）
    df_final = df_sorted.reset_index(drop=True)

    return df_final, df_caigou_pivot


df_CaiGou, df_caigou_pivot = CaiGou(df_caigou=df_caigou, df_old_procurement=df_old_procurement)

with pd.ExcelWriter(f'D:/桌面/{today}采购.xlsx', engine='openpyxl') as writer:
    df_CaiGou.to_excel(writer, sheet_name='采购管理', index=False)
    df_caigou_pivot.to_excel(writer, sheet_name='入库数量', index=False)
