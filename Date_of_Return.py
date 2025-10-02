import os
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import time


# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

# 固定文件
df_new_category = pd.read_excel(r'D:/桌面/模板/固定文件/企划品类.xlsx', sheet_name='品类')

# 采购报表
CWD = f'D:/桌面/货期更新/2024-09-26采购订单'
p = Path(CWD)

file_caigou = list(p.glob('*采购单*.xlsx'))
file_mingxi = list(p.glob('*明细*.xlsx'))

print("file_today:", file_caigou)
print("file_psi:", file_mingxi)

df_caigou = pd.read_excel(file_caigou[0], keep_default_na=False, na_values=['']) 

df_mingxi = pd.read_excel(file_mingxi[0], keep_default_na=False, na_values=[''])

def Date_of_Return(df_caigou=df_caigou, df_mingxi=df_mingxi, df_new_category=df_new_category):

    # 筛选采购单
    df_caigou = df_caigou[df_caigou['状态'].isin(['完成', '已确认'])]

    # 剔除 返修退货，次品，返修
    df_caigou = df_caigou[df_caigou['标记|多标签'] != '返修退货']
    df_caigou['剔除行'] = df_caigou['备注'].str.contains(r'次品|返修', na=False).map({True: '是', False: '否'})
    df_caigou = df_caigou[df_caigou['剔除行'] != '是']

    df_caigou[['颜色', '规格']] = df_caigou['颜色规格'].str.split(';', n=1, expand=True)
    df_caigou['规格'] = df_caigou['规格'].fillna('')
    df_caigou['款色'] = df_caigou['款式编码'] + df_caigou['颜色']
    df_caigou['采购款色'] = df_caigou['采购单号'].astype(str) + df_caigou['款色'].astype(str)
    df_caigou['采购编码'] = df_caigou['采购单号'].astype(str) + df_caigou['商品编码'].astype(str)

    # 将采购完成日期转换为日期格式
    df_caigou['采购单完成时间'] = pd.to_datetime(df_caigou['采购单完成时间'])
    # df_caigou['采购单完成时间new'] = pd.to_datetime(df_caigou['采购单完成时间'])
    df_caigou['采购日期'] = pd.to_datetime(df_caigou['采购日期'])

    # 方法1：直接去重
    df_caigou_status = df_caigou[['采购单号', '状态']].drop_duplicates()
    df_caigou_status = df_caigou_status[df_caigou_status['状态'].notna()]

    df_caigou = df_caigou.merge(df_caigou_status, on='采购单号', how='left', suffixes=('', '_y'))

    # 为每个采购单号创建一个完成时间的映射
    completion_map = df_caigou.dropna(subset=['采购单完成时间']).groupby('采购单号')['采购单完成时间'].first()

    # 填充缺失值
    df_caigou['采购单完成时间'] = df_caigou.apply(
        lambda row: completion_map.get(row['采购单号'], row['采购单完成时间']), 
        axis=1
    )

    ## 入仓明细
    df_mingxi = df_mingxi[df_mingxi['标记|多标签（采购入库/采购退货）'] != '调整单']
    df_mingxi['采购单号'] = pd.to_numeric(df_mingxi['采购单号'], errors='coerce').fillna(0).astype(int)
    df_mingxi['采购编码'] = df_mingxi['采购单号'].astype(str) + df_mingxi['商品编号'].astype(str)
    df_mingxi['入仓时间'] = pd.to_datetime(df_mingxi['入仓时间'])
    df_mingxi['采购日期'] = pd.to_datetime(df_mingxi['采购日期'])

    df_mingxi['采购到完成时间_明细'] = ((df_mingxi['入仓时间'] - df_mingxi['采购日期']).dt.total_seconds() / 86400).round(2)

    # df_mingxi.rename(columns={'入仓时间': '入仓时间(明细)'}, inplace=True)
    # df_mingxi_sub = df_mingxi.loc[:, ['采购编码', '入仓时间(明细)', '出入库数量']]

    # 1. 计算每一行的 (数量 * 时间)
    df_mingxi['数量_时间乘积'] = round(df_mingxi['出入库数量'] * df_mingxi['采购到完成时间_明细'], 2)

    # 2. 按 '采购编码' 分组，并分别计算乘积的总和和数量的总和
    grouped = df_mingxi.groupby('采购编码').agg(
        总数量_时间乘积=('数量_时间乘积', 'sum'),
        总入库数量=('出入库数量', 'sum')
    ).reset_index()

    grouped['总数量_时间乘积'] = round(grouped['总数量_时间乘积'], 2)
    # 3. 计算加权平均时间
    grouped['加权平均入库时间'] = round(grouped['总数量_时间乘积'] / grouped['总入库数量'], 2)

    ## 合并表格
    df_caigou = df_caigou.merge(grouped, on='采购编码', how='left')
    df_caigou['采购编码'] = df_caigou['采购编码'].replace('nan', '').astype(str)

    # latest_time = df_mingxi.groupby('采购编码')['入仓时间'].max().reset_index()
    # latest_time.rename(columns={'入仓时间': '最后入仓时间'}, inplace=True)
    # pivot_df = df_mingxi.pivot_table(
    #     index='采购编码',
    #     values='出入库数量',
    #     aggfunc='sum'
    # )

    # result = pd.merge(pivot_df, latest_time, on='采购编码', how='left')

    df_caigou = df_caigou.merge(df_new_category[['产品分类', '新品类（企划版）']].drop_duplicates(subset=['产品分类']), how='left', left_on='商品分类', right_on='产品分类')


    # df_caigou = df_caigou[df_caigou['数据类型'] == '明细']

    df_caigou_pivot = pd.pivot_table(
        df_caigou,
        index=['采购款色'],
        values=['采购数量', '总入库数量'],
        aggfunc='sum'
    ).reset_index()

    df_caigou_pivot['完成比例'] = round(df_caigou_pivot['总入库数量'] / df_caigou_pivot['采购数量'], 2)
    df_caigou_pivot['完成比例'] = df_caigou_pivot['完成比例'].replace([np.inf, -np.inf], np.nan).fillna(0)
    df_caigou_pivot['是否完结'] = np.where(df_caigou_pivot['完成比例'] >= 0.75, '完结', '未完结')

    df_caigou = df_caigou.merge(df_caigou_pivot[['采购款色', '完成比例', '是否完结']], how='left', on='采购款色')

    # def calculate_weighted_avg(group):
    #     """计算每个款的加权平均回货时间"""
    #     return np.average(group['加权平均入库时间'], weights=group['采购数量'])
    # df_caigou['平均加权回货时间'] = df_caigou.groupby('采购款色').transform(calculate_weighted_avg)


    return df_caigou

df = Date_of_Return()

df.to_clipboard(index=False)
 