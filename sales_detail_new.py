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

# CWD = f'D:/桌面/模板/{today}'

# p = Path(CWD)

# file_sales_details = list(p.glob('*销售明细*.xlsx'))
# print("file_sales_details:", file_sales_details)

## 销售明细 file_sales_details[0]
df_sales_details = pd.read_excel('D:/桌面/模板/2025-10-20/SJDZGQ6S2510销售明细.xlsx', usecols=["付款日期","商品编码","达人编号","款式编码","颜色规格","产品分类","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额",'售后登记日期','售后确认日期','售后分类'])
df_psi = pd.read_excel('D:/桌面/进销存/SJDZGQ6S2510.xlsx', sheet_name='Sheet1', usecols=["日期","商品编码","款色","期初数量"])


def sales_details(df_sales_details=df_sales_details, df_psi=df_psi):
    # df_spi数据清洗
    ## 转换为日期格式
    df_psi['日期'] = pd.to_datetime(df_psi['日期'], errors='coerce')
    df_psi['日期'] = df_psi['日期'].dt.date
    df_psi_pivot = pd.pivot_table(
        df_psi,
        index=['日期','款色'],
        values=['期初数量'],
        aggfunc='sum'
    ).reset_index()
    df_psi_pivot.rename(columns={'期初数量':'当日库存量'}, inplace=True)


    # 仅筛选溶总uid
    df_sales_details["达人编号"] = df_sales_details["达人编号"].astype(str).str.strip()
    df_sales_details = df_sales_details[df_sales_details["达人编号"] != '6524430296']

    #根据分隔符分成两个字段，创建款色
    extracted = df_sales_details['颜色规格'].str.extract(r'^(.*?);(.*)$')
    df_sales_details['颜色'] = extracted[0].fillna(df_sales_details['颜色规格'])  # 无分号时用原值
    df_sales_details['规格'] = extracted[1]  # 无分号时为NaN
    df_sales_details['款色'] = df_sales_details['款式编码'] + df_sales_details['颜色']

    # 转换为日期格式
    df_sales_details['付款日期'] = pd.to_datetime(df_sales_details['付款日期'], errors='coerce')
    df_sales_details['售后登记日期'] = pd.to_datetime(df_sales_details['售后登记日期'], errors='coerce')

    df_sales_details['付款-日期'] = df_sales_details['付款日期'].dt.date
    df_sales_details['售后-日期'] = df_sales_details['售后登记日期'].dt.date

    # 填充缺失值
    df_sales_details['销售数量'] = df_sales_details['销售数量'].fillna(0)
    df_sales_details['退货数量'] = df_sales_details['退货数量'].fillna(0)
    df_sales_details['实发数量'] = df_sales_details['实发数量'].fillna(0)
    df_sales_details['实退数量'] = df_sales_details['实退数量'].fillna(0)

    df_sales_details['销售金额'] = df_sales_details['销售金额'].fillna(0)

    # 计算退货时长（天）
    df_sales_details['退货时长（天）'] =  np.ceil(
        (df_sales_details['售后登记日期'] - df_sales_details['付款日期']).dt.total_seconds() / 86400
    )

    # 筛选仅退款
    df_sales_details_sub = df_sales_details[df_sales_details['售后分类'] == '仅退款'].copy()
    
    # 按付款日期、款色、日期间隔汇总退货数量，统计每日退货数量
    df_sales_details_sub_pivot_day = pd.pivot_table(
        df_sales_details_sub,
        index=['付款-日期','款色','退货时长（天）'],
        values=['退货数量'],
        aggfunc='sum'
    ).reset_index()

    # 先按付款日期、款色、日期间隔排序，再在同一付款日期+款色内累计退货数量
    df_sales_details_sub_pivot_day = df_sales_details_sub_pivot_day.sort_values(['付款-日期','款色','退货时长（天）'])
    df_sales_details_sub_pivot_day['当日累计仅退货量'] = df_sales_details_sub_pivot_day.groupby(['付款-日期', '款色'])['退货数量'].cumsum()


    # 合并当天销售数量
    df_sales_details_pivot = pd.pivot_table(
        df_sales_details,
        index=['付款-日期','款色','产品分类'],
        values=['销售数量','实发数量','实退数量','退货数量'],
        aggfunc='sum'
    ).reset_index()

    df_sales_details_pivot = df_sales_details_pivot.merge(
        df_psi_pivot,
        left_on=['付款-日期','款色'],
        right_on=['日期','款色'],
        how='left'
    )

    df_sales_details_pivot.rename(columns={'销售数量':'当日销售数量','实发数量':'当日实发数量','实退数量':'当日实退数量','退货数量':'当日退货数量'}, inplace=True)
    df_sales_details_pivot['当日发货后退款率'] = round(df_sales_details_pivot['当日实退数量'] / df_sales_details_pivot['当日实发数量'], 2)
    df_sales_details_pivot['当日完整退货率'] = round(df_sales_details_pivot['当日退货数量'] / df_sales_details_pivot['当日销售数量'], 2)
    df_sales_details_pivot.drop(columns=['当日实退数量','当日退货数量'], inplace=True)

    df_aggregated = df_sales_details_sub_pivot_day.merge(
        df_sales_details_pivot,
        on=['付款-日期','款色'],
        how='left'
    )
        

    # 合并当天退货数量
    df_sales_details_pivot = pd.pivot_table(
        df_sales_details_sub,
        index=['付款-日期','款色'],
        values=['退货数量'],
        aggfunc='sum'
    ).reset_index()
    df_sales_details_pivot.rename(columns={'退货数量':'当日仅退款数量'}, inplace=True)

    df_aggregated = df_aggregated.merge(
        df_sales_details_pivot,
        on=['付款-日期','款色'],
        how='left'
    )
    df_aggregated['当日发货前仅退款率'] = round(df_aggregated['当日仅退款数量'] / df_aggregated['当日销售数量'], 2)

    df_aggregated['仅退款率累加占比'] = round(df_aggregated['当日累计仅退货量'] / df_aggregated['当日销售数量'], 2)

    df_aggregated.to_excel(f'D:/桌面/模板/2025-10-20/发货前总退款率中间表.xlsx')

    df_aggregated['是否25天内退款'] = (df_aggregated['退货时长（天）'] <= 25).map({True: '是', False: '否'})
    df_aggregated = df_aggregated[df_aggregated['是否25天内退款'] == '是']
    df_aggregated.drop(columns=['是否25天内退款','日期'], inplace=True)


    df_aggregated = df_aggregated.set_index(['付款-日期','款色','产品分类','退货时长（天）','当日库存量', '当日销售数量', '当日仅退款数量','当日实发数量','当日发货前仅退款率','当日发货后退款率','当日完整退货率']).sort_index()
    df_aggregated = df_aggregated.unstack('退货时长（天）')

    df_aggregated.reset_index(inplace=True)
    df_aggregated.sort_values(by=['款色','付款-日期'], ascending=[False, True], inplace=True)


    return df_aggregated

df = sales_details()
df.to_excel(f'D:/桌面/模板/2025-10-20/发货前总退款率.xlsx')
