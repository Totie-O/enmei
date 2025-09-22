import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import time

# 记录开始时间
start_time = time.time()

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

# 读取文件



## 销售付款时间
df_ndim_payment_time = pd.read_excel(file_ndim_payment_time[0],usecols=["渠道", "店铺","日期","商品编码","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额"])

## 商品资料
df_product = pd.read_excel(file_product[0], usecols=["图片","款式编码","商品编码","颜色","规格","基本售价","成本价","创建时间","分类","年份","季节","商品名称",'供应商名称'])


# 固定文件
df_new_category = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='品类')
df_new_channel = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='渠道')
df_dalei = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='大类')
df_size = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='规则')
df_model = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='模块')
df_old_procurement  = pd.read_excel(r'D:/桌面/模板/固定文件/23年-24年采购明细.xlsx', sheet_name='Sheet1')







def Payment_Time_All(df_ndim_payment_time=df_ndim_payment_time, df_product=df_product, df_new_channel=df_new_channel):
    
    df_product['商品编码'] = df_product['商品编码'].str.strip().str.upper()
    df_ndim_payment_time['商品编码'] = df_ndim_payment_time['商品编码'].str.strip().str.upper()

    df_product['款色'] = df_product['款式编码'] + df_product['颜色']
    
    df_ndim_payment_time['日期'] = pd.to_datetime(df_ndim_payment_time['日期']).dt.date

    # 合并商品编码对应的款色信息
    df_ndim_payment_time = df_ndim_payment_time.merge(
        df_product[['商品编码', '款色']],
        on=['商品编码'],
        how='left'
    )

   # 合并商品编码对应的款色信息
    df_ndim_payment_time = pd.merge(df_ndim_payment_time,
        df_product[['商品编码', '款色']],
        on='商品编码',
        how='left'
    )




    # 修改报表渠道名称
    df_ndim_payment_time.rename(columns={'渠道': '聚水潭报表渠道'}, inplace=True)
    
    # 合并渠道信息
    df_ndim_payment_time = df_ndim_payment_time.merge(
        df_new_channel[['店铺', '渠道']],
        on=['店铺'],
        how='left'
    )

    fill_cols = ['销售金额', '退货金额','销售数量', '退货数量']
    df_ndim_payment_time[fill_cols] = df_ndim_payment_time[fill_cols].fillna(0)


    df_ndim_payment_time['净销成本价'] = df_ndim_payment_time['成本价'] * (df_ndim_payment_time['销售数量'] - df_ndim_payment_time['退货数量'])

    new_columns_order = ["款色","规格","规格终","渠道","店铺","日期","商品编码","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额","净销成本价","聚水潭报表渠道"]
    df_ndim_payment_time = df_ndim_payment_time.reindex(columns=new_columns_order)


    ## 支付时间透视
    df_ndim_payment_time_pivot = pd.pivot_table(
        df_ndim_payment_time,
        index= ['款色'],
        values=['销售金额', '退货金额','销售数量', '退货数量', '净销成本价'],
        aggfunc='sum'
    ).reset_index()

    df_ndim_payment_time_pivot['净销额'] = df_ndim_payment_time_pivot['销售金额'] - df_ndim_payment_time_pivot['退货金额']
    df_ndim_payment_time_pivot['毛利'] = round(1 - (df_ndim_payment_time_pivot['净销成本价'] / df_ndim_payment_time_pivot['净销额']), 4)
    df_ndim_payment_time_pivot['件单价'] = round(df_ndim_payment_time_pivot['销售金额'] / df_ndim_payment_time_pivot['销售数量'], 0)
    
    ## 最早支付时间
    # df_ndim_payment_time_new = df_ndim_payment_time[['款色', '日期']]
    # df_ndim_payment_time_new = df_ndim_payment_time_new.sort_values(by='日期')

    # 按款色和日期分组，计算每天的销售数量
    daily_sales = df_ndim_payment_time.groupby(['款色', '日期'])['销售数量'].sum().reset_index()
    # 筛选出每天销售数量大于10的记录
    high_sales_days = daily_sales[daily_sales['销售数量'] > 10]

    # 对每个款色，找到最早的出现销售数量大于10的日期
    df_ndim_payment_time_new = high_sales_days.groupby('款色')['日期'].min().reset_index()
    
    df_ndim_payment_time_new = df_ndim_payment_time_new.sort_values(by='日期')



    return df_ndim_payment_time, df_ndim_payment_time_pivot, df_ndim_payment_time_new


df = Payment_Time_All()
df.to_excel(f'D:/桌面/付款时间.xlsx', index=False)