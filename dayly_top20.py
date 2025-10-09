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
print(today)

CWD = f'D:/桌面/模板/{today}'

p = Path(CWD)

file_sales_details = list(p.glob('*销售明细*.xlsx'))
print("file_sales_details:", file_sales_details)

file_product = [i for i in list(p.glob('*商品资料*.xlsx')) if '库存视角' not in str(i)]
print("file_product:", file_product)



## 销售明细
df_sales_details = pd.read_excel(file_sales_details[0], usecols=[ "内部订单号","店铺","付款日期","商品编码","达人编号","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额"])

## 商品资料
df_product = pd.read_excel(file_product[0], usecols=["图片","款式编码"])


def sales_details(df_sales_details=df_sales_details):
    # 获取当前日期和时间
    now = datetime.now()

    # 计算昨天日期
    yesterday = now.date() - timedelta(days=1)


    df_sales_details['付款日期'] = pd.to_datetime(df_sales_details['付款日期'], errors='coerce')

    df_sales_details = df_sales_details[df_sales_details['付款日期'].dt.date == yesterday]
    # 如果后续只需要日期部分，可以转换为 date
    # df_sales_details['付款日期'] = df_sales_details['付款日期'].dt.date


    df_sales_details['销售数量'] = df_sales_details['销售数量'].fillna(0)
    df_sales_details['销售金额'] = df_sales_details['销售金额'].fillna(0)

    df_sales_details_pivot = (
        pd.pivot_table(
            df_sales_details,
            index=['款式编码'],
            values=['销售数量', '销售金额'],
            aggfunc='sum'
        )
        .reset_index()
        .sort_values('销售金额', ascending=False)  # 按销售金额降序
        .head(20)  # 取前20
        .reset_index(drop=True)
    )

    sort = ["排名", "款式编码","图片","上新时间","类目","售价(参考)","销售数量","销售金额"]
    df_sales_details_pivot['排名'] = df_sales_details_pivot.index + 1
    df_sales_details_pivot = df_sales_details_pivot.reindex(columns=sort)

    df_sales_details_pivot['销售数量'] = df_sales_details_pivot['销售数量'].astype(int)

    return df_sales_details, df_sales_details_pivot

df_sales_details, df_sales_details_pivot = sales_details()



def pic_dl(df_product=df_product, df_sales_details_pivot=df_sales_details_pivot):
    df_product = df_product.drop_duplicates(subset=['款式编码'], keep='first')
    df_merged = pd.merge(df_sales_details_pivot, df_product, on='款式编码', how='left',suffixes=('', '_y'))

    for n, row in df_merged.iterrows():
        download_and_compress_image_plus(row['图片_y'], row['款式编码'], output_dir="D:\\图片800_800", high=1000, wide=1000)
    
    return df_merged

if __name__ == "__main__":

    df_merged = pic_dl(df_product=df_product, df_sales_details_pivot=df_sales_details_pivot)
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

    df_huo_new_time['Unnamed: 9'] = pd.to_datetime(df_huo_new_time['Unnamed: 9'], errors='coerce').dt.date
    df_huo_new_time = df_huo_new_time.sort_values(by='Unnamed: 9', ascending=True)
    df_huo_new_time = df_huo_new_time.drop_duplicates(subset=['公式在此行保管，保持第4行有公式其他都复制成值'], keep='first')

    # print(df_huo_new_time)


    with pd.ExcelWriter(f'D:/桌面//模板/{today}/{today}销售top20.xlsx', engine='openpyxl') as writer:
        df_sales_details.to_excel(writer, sheet_name='销售明细', index=False)
        df_sales_details_pivot = df_sales_details_pivot.merge(df_huo_new_time, left_on='款式编码', right_on='公式在此行保管，保持第4行有公式其他都复制成值', how='left')
        df_sales_details_pivot.rename(columns={'Unnamed: 9': '上新时间'}, inplace=True)
        df_sales_details_pivot.to_excel(writer, sheet_name='销售top20', index=False)







