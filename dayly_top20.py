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

file_sales_details = list(p.glob('*销售明细*.xlsx'))
print("file_sales_details:", file_sales_details)

file_product = [i for i in list(p.glob('*商品资料*.xlsx')) if '库存视角' not in str(i)]
print("file_product:", file_product)



## 销售明细
df_sales_details = pd.read_excel(file_sales_details[0], usecols=[ "内部订单号","店铺","付款日期","商品编码","达人编号","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额",'线上商品名'])

## 商品资料
df_product = pd.read_excel(file_product[0], usecols=["图片","款式编码"])
df_product['款式编码'] = df_product['款式编码'].astype(str).str.upper()



def sales_details(df_sales_details=df_sales_details):
    # 获取当前日期和时间
    now = datetime.now()

    # 计算昨天日期
    yesterday = now.date() - timedelta(days=1)
    # 计算一周前的日期
    # 保持类型一致：全部使用 date 进行比较，避免 datetime 与 date 混比导致 TypeError
    one_week_ago = (now - timedelta(days=7)).date()

    df_sales_details['付款日期'] = pd.to_datetime(df_sales_details['付款日期'], errors='coerce')

    df_sales_details['销售数量'] = df_sales_details['销售数量'].fillna(0)
    df_sales_details['销售金额'] = df_sales_details['销售金额'].fillna(0)

    df_sales_details_yesterday = df_sales_details[df_sales_details['付款日期'].dt.date == yesterday]


    ####################################切片######################################
    df_sales_details_yesterday_copy = df_sales_details_yesterday.copy()
    _title = df_sales_details_yesterday_copy['线上商品名'].astype(str)
    df_sales_details_yesterday_copy['商品简称'] = _title.str.extract(r'【([^】]+)】', expand=False)
    df_sales_details_yesterday_copy['商品简称'].fillna(_title, inplace=True)
    df_sales_details_yesterday_copy_name = df_sales_details_yesterday_copy[['款式编码',  '商品简称','线上商品名']].drop_duplicates(subset=['款式编码'], keep='first')
    df_sales_details_yesterday_copy_name_first = df_sales_details_yesterday_copy_name.groupby('款式编码').first().reset_index()


    df_sales_details_yesterday_copy_pivot = pd.pivot_table(
        df_sales_details_yesterday_copy,
        index=['款式编码'],
        values=['销售数量', '销售金额'],
        aggfunc='sum'
    ).reset_index()

    # 按销售金额降序并计算占比，保留前36
    df_sales_details_yesterday_copy_pivot = (
        df_sales_details_yesterday_copy_pivot
        .sort_values('销售金额', ascending=False)
        .reset_index(drop=True)
    )
    _tot_amt = df_sales_details_yesterday_copy_pivot['销售金额'].sum()
    _tot_qty = df_sales_details_yesterday_copy_pivot['销售数量'].sum()
    df_sales_details_yesterday_copy_pivot['昨日销售金额占比'] = (
        df_sales_details_yesterday_copy_pivot['销售金额'] / _tot_amt
    ).round(4)
    df_sales_details_yesterday_copy_pivot['昨日销售数量占比'] = (
        df_sales_details_yesterday_copy_pivot['销售数量'] / _tot_qty
    ).round(4)
    df_sales_details_yesterday_copy_pivot = df_sales_details_yesterday_copy_pivot.head(36)
    df_sales_details_yesterday_copy_pivot['排名'] = df_sales_details_yesterday_copy_pivot.index + 1

    df_sales_details_yesterday_copy_pivot = df_sales_details_yesterday_copy_pivot[
        ['排名','款式编码','昨日销售金额占比','昨日销售数量占比']
    ]
    df_sales_details_yesterday_copy_pivot = df_sales_details_yesterday_copy_pivot.merge(
        df_sales_details_yesterday_copy_name_first,
        on='款式编码',
        how='left'
    )

    df_sales_details_yesterday_copy_pivot['图片'] = ''
    df_sales_details_yesterday_copy_pivot.reindex(
        columns=['排名','款式编码','图片','昨日销售金额占比','昨日销售数量占比','商品简称','线上商品名']
    )
    ####################################切片######################################

    # 筛选过去一周的数据（包含昨天）
    df_week = df_sales_details[
        (df_sales_details['付款日期'].dt.date >= one_week_ago) &
        (df_sales_details['付款日期'].dt.date <= yesterday)
    ]
    # 如果后续只需要日期部分，可以转换为 date
    # df_sales_details['付款日期'] = df_sales_details['付款日期'].dt.date

    def df_pivot(df, num=20):
        df_sales_details_pivot = (
            pd.pivot_table(
                df,
                index=['款式编码'],
                values=['销售数量', '销售金额'],
                aggfunc='sum'
            )
            .reset_index()
            .sort_values('销售金额', ascending=False)  # 按销售金额降序
            .head(num)  # 取前20
            .reset_index(drop=True)
        )

        sort = ["排名", "款式编码","上新时间","类目","售价(参考)","销售金额","销售数量"]
        df_sales_details_pivot['排名'] = df_sales_details_pivot.index + 1
        df_sales_details_pivot = df_sales_details_pivot.reindex(columns=sort)

        df_sales_details_pivot['销售数量'] = df_sales_details_pivot['销售数量'].astype(int)

        return df_sales_details_pivot
    
    df_sales_details_pivot_30 = df_pivot(df_week, num=30)
    df_sales_details_pivot_20 = df_pivot(df_sales_details_yesterday, num=20)

    return df_sales_details_yesterday, df_sales_details_pivot_20, df_week, df_sales_details_pivot_30, df_sales_details_yesterday_copy_pivot

df_sales_details_yesterday, df_sales_details_pivot_20, df_week, df_sales_details_pivot_30, df_sales_details_yesterday_copy_pivot = sales_details()



def pic_dl(df_product=df_product, df_sales_details_pivot=df_sales_details_pivot_20):
    df_product = df_product[df_product['图片'].notna()]
    df_product = df_product.drop_duplicates(subset=['款式编码'], keep='first')
    df_merged = pd.merge(df_sales_details_pivot, df_product, on='款式编码', how='left',suffixes=('', '_y'))

    for n, row in df_merged.iterrows():
        download_and_compress_image_plus(row['图片'], row['款式编码'], output_dir="D:\\图片800_800", high=1000, wide=1000)
    
    return df_merged

if __name__ == "__main__":

    df_merged_1 = pic_dl(df_product=df_product, df_sales_details_pivot=df_sales_details_pivot_20)
    df_merged_2 = pic_dl(df_product=df_product, df_sales_details_pivot=df_sales_details_pivot_30)
    df_merged_3 = pic_dl(df_product=df_product, df_sales_details_pivot=df_sales_details_yesterday_copy_pivot)
    # df_merged.to_clipboard(index=False)

    # 获取当天日期
    today_small = datetime.now().strftime("%y.%m.%d")
    print(today_small)

    CWD = f'D:/桌面/过往货盘表/'

    p = Path(CWD)
    file_path = list(p.glob(f'*{today_small}*.xlsx'))[0]
    print(file_path)
    df_huo = open_protected_excel_safe(file_path, password="789", sheet_name='Sheet1')
    df_huo_new_time = df_huo.iloc[:, [1,9]]

    # 统一将视觉空值转换为真正的 NaN
    df_huo_new_time['Unnamed: 9'] = (
        df_huo_new_time['Unnamed: 9']
        .astype(str)
        .str.strip()
        .replace({'': np.nan, 'NaN': np.nan, 'nan': np.nan, 'None': np.nan})
    )

    df_huo_new_time = df_huo_new_time[df_huo_new_time['Unnamed: 9'].notna()]
    df_huo_new_time['Unnamed: 9'] = pd.to_datetime(df_huo_new_time['Unnamed: 9'], errors='coerce').dt.date
    df_huo_new_time = df_huo_new_time.sort_values(by='Unnamed: 9', ascending=True)
    df_huo_new_time = df_huo_new_time.drop_duplicates(subset=['公式在此行保管，保持第4行有公式其他都复制成值'], keep='first')
    df_huo_new_time.columns = ['款号', '上新日期']
    df_huo_new_time['款号'] = df_huo_new_time['款号'].astype(str).str.upper()
    # print(df_huo_new_time)


    with pd.ExcelWriter(f'D:/桌面//模板/{today}/{today}销售top20.xlsx', engine='openpyxl') as writer:
        # df_sales_details_yesterday.to_excel(writer, sheet_name='昨日销售明细', index=False)

        df_sales_details_pivot = df_merged_1.merge(df_huo_new_time, left_on='款式编码', right_on='款号', how='left')
        df_sales_details_pivot.to_excel(writer, sheet_name='昨日销售top20', index=False)

        # df_week.to_excel(writer, sheet_name='过往一周销售明细', index=False)
        df_sales_details_pivot_week = df_merged_2.merge(df_huo_new_time, left_on='款式编码', right_on='款号', how='left')
        df_sales_details_pivot_week.to_excel(writer, sheet_name='过往销售top30', index=False)

        df_sales_details_yesterday_copy_pivot.to_excel(writer, sheet_name='切片top36', index=False)








