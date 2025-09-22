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
df_new_category = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='品类')
df_new_channel = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='渠道')
df_dalei = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='大类')
df_size = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='规则')
df_model = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='模块')
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

# 读取文件

## 采购
df_caigou = pd.read_excel(file_caigou[0]) 

## 销售发货时间
df_ndim_deliver_time = pd.read_excel(file_ndim_deliver_time[0], usecols=["渠道", "店铺","日期","商品编码","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额"])

## 销售付款时间
df_ndim_payment_time = pd.read_excel(file_ndim_payment_time[0],usecols=["渠道", "店铺","日期","商品编码","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额"])

## 商品资料
df_product = pd.read_excel(file_product[0], usecols=["图片","款式编码","商品编码","颜色","规格","基本售价","成本价","创建时间","分类","年份","季节","商品名称",'供应商名称'])

## 销售明细
df_sales_details = pd.read_excel(file_sales_details[0], usecols=[ "内部订单号","店铺","付款日期","商品编码","达人编号","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额"])

## 当日销售
df_today_sales = pd.read_excel(file_today[0])

## 库存视角
df_psi = pd.read_excel(file_psi[0], usecols=[ "款式编码","商品编码","颜色","规格","实际库存数","订单占有数","销退仓库存","进货仓库存","采购在途数","销退在途数","虚拟库存数","可用数","公有可用数"])


def Product_Output(df_product=df_product):
    new_columns_order = ["款式编码","图片","商品编码","颜色","规格","基本售价","成本价","创建时间",'供应商名称',"分类","年份","季节","商品名称"]
    df_product = df_product.reindex(columns=new_columns_order)

    df_product['创建时间'] = pd.to_datetime(df_product['创建时间'])

    return df_product

# 筛选出货盘报表

def filter(df_psi=df_psi, df_product=df_product, 
           df_ndim_payment_time=df_ndim_payment_time, 
           df_new_category=df_new_category):
    """
    数据处理与过滤函数
    返回按款式编码和颜色去重后的商品子集
    """
    # ==================== 3. 销售数据透视 ====================

    df_ndim_payment_time['销售数量'] = df_ndim_payment_time['销售数量'].fillna(0)
    # 创建款色组合列
    df_product['款色'] = df_product['款式编码'] + df_product['颜色']
    
    df_product['商品编码'] = df_product['商品编码'].str.strip().str.upper()
    df_ndim_payment_time['商品编码'] = df_ndim_payment_time['商品编码'].str.strip().str.upper()
    

    # 合并商品编码对应的款色信息
    df_ndim_payment_time = df_ndim_payment_time.merge(
        df_product[['商品编码', '款色']],
        on=['商品编码'],
        how='left'
    )
    
    # 创建销售数据透视表
    df_ndim_payment_time_pivot = pd.pivot_table(
        df_ndim_payment_time,
        index='款色',
        values=['销售数量'],
        aggfunc={'销售数量': 'sum'}
    ).reset_index()

    # ==================== 2. 库存数据透视 ====================
    df_psi['款色'] = df_psi['款式编码'] + df_psi['颜色']

    psi_fill_cols = ['实际库存数', '订单占有数', '采购在途数']
    df_psi[psi_fill_cols] = df_psi[psi_fill_cols].fillna(0)

    df_psi_pivot = pd.pivot_table(
        df_psi,
        index='款色',
        values=['实际库存数', '订单占有数', '销退仓库存', '进货仓库存',
                '采购在途数', '销退在途数', '虚拟库存数', '可用数', '公有可用数'],
        aggfunc='sum'
    ).reset_index()
    
    # ==================== 1. 商品资料处理 ====================
    # 款色去重
    df_product = df_product.drop_duplicates(
        subset=['款式编码', '颜色'],
        keep='first'
    )
    # 剔除辅料和邮费
    df_product = df_product[(df_product['分类'] != '辅料') & (df_product['分类'] != '邮费')]
    # 颜色不为空
    df_product = df_product[(df_product['颜色'].notna())]

    # ==================== 4. 数据合并 ====================
    df_product = df_product.merge(df_psi_pivot, on='款色', how='left')
    df_product = df_product.merge(
        df_ndim_payment_time_pivot[['款色', '销售数量']],
        on='款色',
        how='left'
    )

    # ==================== 5. 数据过滤 ====================
    fill_cols = ['实际库存数', '订单占有数', '采购在途数', '销售数量']
    df_product[fill_cols] = df_product[fill_cols].fillna(0)
        
    df_product = df_product[
        # 条件1：三个字段都不为0
        (
            (df_product['实际库存数'] != 0) |
            (df_product['订单占有数'] != 0) |
            (df_product['销售数量'] != 0) |
            (df_product['采购在途数'] != 0)
        )       
    ]
    
    # # ==================== 6. 去重处理 ====================
    df_product_sub = df_product.loc[:, ['款色', '款式编码',  '颜色','图片']]
    df_product_sub['颜色'] = df_product_sub['颜色'].str.replace('_x002B_', '+')
    df_product_sub['款色'] = df_product_sub['款色'].str.replace('_x002B_', '+')

    df_product_sub = df_product_sub[(df_product_sub['款式编码'] != 'SJPRRYMC1779W')]

    return df_product_sub


def Psi(df_psi=df_psi, df_size=df_size):

    df_psi['款色'] = df_psi['款式编码'] + df_psi['颜色']
    df_psi['规格'] = df_psi['规格'].str.replace(r'\s*\([^)]*\)', '', regex=True)
    df_psi['规格'] = df_psi['规格'].str.replace(r'\s*【[^】]*】', '', regex=True)
    df_psi['规格'] = df_psi['规格'].str.replace(r'\s*（[^）]*）', '', regex=True)
    df_psi['规格'] = df_psi['规格'].str.replace(r'\s*\([^）]*）', '', regex=True)


    df_psi = df_psi.merge(
        df_size[['规格', '规格终']],
        on=['规格'],
        how='left'
    )

    new_columns_order = ['款色',"款式编码","商品编码","颜色","规格","规格终","实际库存数","订单占有数","销退仓库存","进货仓库存","采购在途数","销退在途数","虚拟库存数","可用数","公有可用数"]
    df_psi = df_psi.reindex(columns=new_columns_order)

    return df_psi


df_psi_output = Psi(df_psi=df_psi, df_size=df_size)


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
        values=['总入库量'],
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
        '采购数量': 'sum'
    })


    # 3. 按照款色升序、日期降序排序
    df_sorted = df_grouped.sort_values(['款色', '采购日期'], ascending=[True, False])

    # 重置索引（可选）
    df_final = df_sorted.reset_index(drop=True)

    return df_final, df_caigou_pivot


def today_sales(df_today_sales=df_today_sales, df_product=df_product):
    df_today_sales['销售数量'] = df_today_sales['销售数量'].fillna(0)
    df_today_sales.drop(columns=['款色'], inplace=True)
    
    df_product['商品编码'] = df_product['商品编码'].str.strip().str.upper()
    df_today_sales['商品编码'] = df_today_sales['商品编码'].str.strip().str.upper()


    df_product['款色'] = df_product['款式编码'] + df_product['颜色']

    df_today_sales = df_today_sales.merge(
        df_product[['商品编码', '款色']],
        on=['商品编码'],
        how='left'
    )

    df_today_sales_pivot = pd.pivot_table(
        df_today_sales,
        index='款色',
        values=['销售数量'],
        aggfunc='sum'
    ).reset_index()

    return df_today_sales_pivot


def Sales_Details(df_sales_details=df_sales_details, df_product=df_product, df_new_channel=df_new_channel,df_model=df_model):

    df_sales_details['销售数量'] = df_sales_details['销售数量'].fillna(0)

    df_product['商品编码upper'] = df_product['商品编码'].str.strip().str.upper()
    df_sales_details['商品编码upper'] = df_sales_details['商品编码'].str.strip().str.upper()

    df_product['款色'] = df_product['款式编码'] + df_product['颜色']

    df_sales_details = df_sales_details.merge(
        df_product[['商品编码upper', '款色']],
        on=['商品编码upper'],
        how='left'
    )

    df_model['UID'] = df_model['UID'].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)
    df_model['UID'] = df_model['UID'].str.strip()

    df_sales_details['付款日期'] = pd.to_datetime(df_sales_details['付款日期']).dt.date


    df_sales_details['达人编号'] = df_sales_details['达人编号'].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)
    df_sales_details['达人编号'] = df_sales_details['达人编号'].str.strip()
    
    # 定义处理函数
    def process_row(row):
        if pd.isna(row['款色']) and pd.notna(row['产品分类']) and pd.notna(row['颜色规格']):
            # 拆分颜色规格
            if ';' in row['颜色规格']:
                color, spec = row['颜色规格'].split(';', 1)
                row['颜色'] = color
                row['规格'] = spec
                row['款色'] = f"{row['款式编码']}+{color}"
        return row

    # 应用处理函数
    df_sales_details = df_sales_details.apply(process_row, axis=1)

    df_sales_details = df_sales_details.merge(
        df_new_channel[['店铺', '渠道']],
        on=['店铺'],
        how='left'
    )

    new_columns_order = ['款色',"店铺",'付款日期',"渠道","款式编码","内部订单号","商品编码","颜色规格","颜色","规格","达人编号","产品分类","销售数量","实发数量","退货数量","实退数量"]
    df_sales_details = df_sales_details.reindex(columns=new_columns_order)
    df_sales_details = df_sales_details.merge(
        df_model[['模块', 'UID']],
        left_on=['达人编号'], right_on=['UID'],
        how='left'
    )
    df_sales_details['达人编号'] = df_sales_details['达人编号'].str.replace('nan', '')


    return df_sales_details


def Deliver_Time_SUM(df_ndim_deliver_time=df_ndim_deliver_time, df_product=df_product):

    df_product['商品编码'] = df_product['商品编码'].str.strip().str.upper()
    df_ndim_deliver_time['商品编码'] = df_ndim_deliver_time['商品编码'].str.strip().str.upper()

    df_product['款色'] = df_product['款式编码'] + df_product['颜色']
    
    df_ndim_deliver_time['日期'] = pd.to_datetime(df_ndim_deliver_time['日期']).dt.date

    # 合并商品编码对应的款色信息
    df_ndim_deliver_time = df_ndim_deliver_time.merge(
        df_product[['商品编码', '款色']],
        on=['商品编码'],
        how='left'
    )
    fill_cols = ['实发数量','实退数量']
    df_ndim_payment_time[fill_cols] = df_ndim_payment_time[fill_cols].fillna(0)
    
    # 创建销售数据透视表
    df_ndim_deliver_time_pivot = pd.pivot_table(
        df_ndim_deliver_time,
        index= ['日期','款色'],
        values=['实发数量','实退数量'],
        aggfunc='sum'
    ).reset_index()


    return df_ndim_deliver_time_pivot


def Deliver_Time_All(df_ndim_deliver_time=df_ndim_deliver_time, df_product=df_product, df_psi_output=df_psi_output, df_new_channel=df_new_channel):
    
    df_product['商品编码'] = df_product['商品编码'].str.strip().str.upper()
    df_ndim_deliver_time['商品编码'] = df_ndim_deliver_time['商品编码'].str.strip().str.upper()

    df_product['款色'] = df_product['款式编码'] + df_product['颜色']
    
    df_ndim_deliver_time['日期'] = pd.to_datetime(df_ndim_deliver_time['日期']).dt.date

    # 合并商品编码对应的款色信息
    df_ndim_deliver_time = df_ndim_deliver_time.merge(
        df_product[['商品编码', '款色']],
        on=['商品编码'],
        how='left'
    )

    # 合并库存尺码
    df_ndim_deliver_time = df_ndim_deliver_time.merge(
        df_psi_output[['商品编码', '规格', '规格终']],
        on=['商品编码'],
        how='left'
    )

    # 修改报表渠道名称
    df_ndim_deliver_time.rename(columns={'渠道': '聚水潭报表渠道'}, inplace=True)
    
    # 合并渠道信息
    df_ndim_deliver_time = df_ndim_deliver_time.merge(
        df_new_channel[['店铺', '渠道']],
        on=['店铺'],
        how='left'
    )

    df_ndim_deliver_time['净销成本价'] = df_ndim_deliver_time['成本价'] * (df_ndim_deliver_time['销售数量'] - df_ndim_deliver_time['退货数量'])

    new_columns_order = ["款色","规格","规格终","渠道","店铺","日期","商品编码","款式编码","颜色规格","产品分类","成本价","销售数量","实发数量","实发金额","销售金额","退货数量","实退数量","退货金额","实退金额","净销成本价","聚水潭报表渠道"]
    df_ndim_deliver_time = df_ndim_deliver_time.reindex(columns=new_columns_order)
    

    return df_ndim_deliver_time


def Payment_Time_All(df_ndim_payment_time=df_ndim_payment_time, df_product=df_product, df_size=df_size, df_new_channel=df_new_channel):
    
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

    # # 先筛选出规格为空的行
    # mask = df_ndim_payment_time['规格'].isna() & df_ndim_payment_time['颜色规格'].notna()

    # # 对这些行进行处理，添加错误处理
    # df_ndim_payment_time.loc[mask, '规格'] = df_ndim_payment_time.loc[mask, '颜色规格'].apply(
    #     lambda x: x.split(';')[1].strip() if pd.notna(x) and len(x.split(';')) > 1 else np.nan
    # )
    #根据分隔符分成两个字段
    extracted = df_ndim_payment_time['颜色规格'].str.extract(r'^(.*?);(.*)$')
    df_ndim_payment_time['颜色'] = extracted[0].fillna(df_ndim_payment_time['颜色规格'])  # 无分号时用原值
    df_ndim_payment_time['规格'] = extracted[1]  # 无分号时为NaN

    # 合并库存尺码
    df_ndim_payment_time = df_ndim_payment_time.merge(
        df_size[['规格', '规格终']],
        on=['规格'],
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




df_Paymemt_Time_All, df_ndim_payment_time_pivot, df_ndim_payment_time_new = Payment_Time_All(df_ndim_payment_time=df_ndim_payment_time, df_product=df_product, df_size=df_size, df_new_channel=df_new_channel)

df_Deliver_Time_All = Deliver_Time_All(df_ndim_deliver_time=df_ndim_deliver_time, df_product=df_product, df_psi_output=df_psi_output, df_new_channel=df_new_channel)

df_deliver_time_sum = Deliver_Time_SUM(df_ndim_deliver_time=df_ndim_deliver_time, df_product=df_product)

df_details = Sales_Details(df_sales_details=df_sales_details, df_product=df_product)


df_today = today_sales(df_today_sales, df_product)


df_CaiGou, df_caigou_pivot = CaiGou(df_caigou=df_caigou, df_old_procurement=df_old_procurement)

df_product_output = Product_Output(df_product=df_product)

df_filter = filter(df_psi=df_psi, df_product=df_product, 
           df_ndim_payment_time=df_ndim_payment_time,
              df_new_category=df_new_category)


with pd.ExcelWriter(f'D:/桌面/新版货盘表/{today}货盘表数据表.xlsx', engine='openpyxl') as writer:
    df_filter.to_excel(writer, sheet_name='款号筛选', index=False)
    df_product_output.to_excel(writer, sheet_name='商品资料', index=False)
    df_psi_output.to_excel(writer, sheet_name='库存视角', index=False)
    df_CaiGou.to_excel(writer, sheet_name='采购管理', index=False)
    df_caigou_pivot.to_excel(writer, sheet_name='入库数量', index=False)
    df_today.to_excel(writer, sheet_name='当天销售报表', index=False)
    df_details.to_excel(writer, sheet_name='过往一周销售明细', index=False)
    df_deliver_time_sum.to_excel(writer, sheet_name='近两月发货时间透视', index=False)
    df_ndim_payment_time_pivot.to_excel(writer, sheet_name='近两月付款时间透视', index=False)
    df_Deliver_Time_All.to_excel(writer, sheet_name='近两月发货时间明细', index=False)
    df_Paymemt_Time_All.to_excel(writer, sheet_name='近两月付款时间明细', index=False)
    df_ndim_payment_time_new.to_excel(writer, sheet_name='最早付款时间', index=False)

# 记录结束时间并计算耗时
end_time = time.time()
run_time = end_time - start_time

print(f"程序运行了 {run_time:.2f} 秒")







