import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import time

# 记录开始时间
start_time = time.time()# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

CWD = f'D:/桌面/模板/{today}'
p = Path(CWD)

file_product = [i for i in list(p.glob('*商品资料*.xlsx')) if '库存视角' not in str(i)]

print("file_product:", file_product)

## 商品资料
df_product = pd.read_excel(file_product[0], usecols=["图片","款式编码","商品编码","颜色","规格","年份","季节",'分类','创建年份'])

df_new_category = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='品类')

df_product_merged = pd.merge(df_product, df_new_category, how='left', left_on='分类', right_on='产品分类')


df_product_merged = df_product_merged.loc[:, ['款式编码','商品编码','颜色','图片','规格','年份','季节','分类','产品分类','创建年份']]

# SPU
df_product_merged_SPU = df_product_merged.loc[:, ['款式编码','图片','年份','季节','分类','产品分类','创建年份']]
df_product_merged_SPU = df_product_merged_SPU.drop_duplicates(subset=['款式编码'])

# SKC
df_product_merged['款色'] = df_product_merged['款式编码'].astype(str) + df_product_merged['颜色'].astype(str)
df_product_merged_SKC = df_product_merged.loc[:, ['款色', '款式编码','颜色','图片','年份','季节','分类','产品分类','创建年份']]
df_product_merged_SKC = df_product_merged_SKC.drop_duplicates(subset=['款色'])

# SKU
df_product_merged_SKU = df_product_merged.loc[:, ['商品编码', '款色', '款式编码','图片','颜色','规格','年份','季节','分类','产品分类','创建年份']]


with pd.ExcelWriter(f'D:/桌面/每日商品资料/{today}SPU&SKC&SKU.xlsx', engine='openpyxl') as writer:
    df_product_merged_SPU.to_excel(writer, sheet_name='SPU', index=False)
    df_product_merged_SKC.to_excel(writer, sheet_name='SKC', index=False)
    df_product_merged_SKU.to_excel(writer, sheet_name='SKU', index=False)
