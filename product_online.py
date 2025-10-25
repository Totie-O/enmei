import msoffcrypto
import pandas as pd
from io import BytesIO
import re
import time
from pathlib import Path
from datetime import datetime
import numpy as np
from dayly_pic_download import open_protected_excel_safe
import shutil

# 读取货盘表
#############################################################
# 获取当天日期
today_small = datetime.now().strftime("%y.%m.%d")
print(today_small)

CWD_huo = f'D:/桌面/过往货盘表/'

p_huo = Path(CWD_huo)
file_path = list(p_huo.glob(f'*{today_small}*.xlsx'))[0]
print(file_path)

def huo():
    password = "789"

    df = open_protected_excel_safe(file_path, password, sheet_name='Sheet1')
    # 如果是想用“第三行”当表头（常见需求）：
    df.columns = df.iloc[1].astype(str)  # 第三行作为新列名
    df = df.iloc[2:].reset_index(drop=True)  # 去掉前三行

    df_stack_10 = df[df['在仓总库存'] >= 10].copy()
    df_order_10 = df[df['订单占有数'] >= 10].copy()

    df = pd.concat([df_stack_10, df_order_10]).drop_duplicates().reset_index(drop=True)

    df_sub = df.loc[:, ['货号', '货号+色号', '颜色', '上新日', '企划类目', '在仓总库存', '订单占有数', '近3天日均销量']]
    df_sub['上新日'] = pd.to_datetime(df_sub['上新日'], errors='coerce').dt.date

    df_sub['近3天日均销量'] = df_sub['近3天日均销量'].astype(int)
    df_sub = df_sub.sort_values(by=['近3天日均销量'], ascending=False).reset_index(drop=True)


    df_sub = df_sub[df_sub['企划类目'] != '美妆']
    df_sub['前200top'] = np.where(df_sub.index < 200, '是', '否')
    # print(df.head())
    df_sub.to_excel(f'D:/桌面/店铺在架/结果输出/{today_small}_处理后.xlsx', index=False)

    return df_sub



# 读取商品资料
#############################################################
# 获取当天日期
today = datetime.now().strftime("%Y-%m-%d")
print(today)

CWD = f'D:/桌面/模板/{today}'
p = Path(CWD)
file_product = [i for i in list(p.glob('*商品资料*.xlsx')) if '库存视角' not in str(i)]
print("file_product:", file_product)

def df_read_product():
    ## 商品资料
    df_product = pd.read_excel(file_product[0], usecols=["国标码","款式编码","商品编码",'颜色'])

    return df_product



# 读取店铺在架
#############################################################
df_online = pd.read_excel('D:/桌面/店铺在架/10.21店铺在架.xlsx', usecols=["店铺名称","站点名称","平台店铺款式编码","平台店铺商品编码","线上款式编码","线上商品编码","原始商品编码","线上商品名称","线上颜色规格","店铺库存","店铺售价","是否上架","系统款式编码","系统商品编码","系统商品名称","系统颜色规格"], sheet_name='Sheet1')

def product_online(df_product, df_online=df_online):

    df_online_wei = df_online[df_online['店铺名称'] == '溶溶RongRong女装（唯品）']
    df_online_other = df_online[df_online['店铺名称'] != '溶溶RongRong女装（唯品）']

    df_online_wei = df_online_wei.merge(df_product, how='left', left_on='线上款式编码', right_on='国标码')
    df_online_wei.drop(columns=['系统款式编码','系统商品编码','颜色','国标码'], inplace=True)
    df_online_wei.rename(columns={'款式编码': '系统款式编码', '商品编码': '系统商品编码'}, inplace=True)

    df_concat = pd.concat([df_online_wei, df_online_other], axis=0, ignore_index=True)

    df_merge = df_concat.merge(df_product, how='left', left_on=['系统商品编码'], right_on=['商品编码'])

    site_channel_map = {
        "WXChannel": "微信",
        "TouTiaoFXG": "头条",
        "KWaiShop": "快手",
        "Xiaohs": "小红书",
        "Pinduoduo": "拼多多",
        "Tmall": "天猫",
        "TaobaoMarket": "淘宝",
        "Kdt": "有赞",
        "BaiduXD": "百度",
        "Jos": "京东",
        "Vipapis": "唯品",
        "DeWu": "得物",
    }
    df_merge["渠道"] = df_merge["站点名称"].map(site_channel_map).fillna("未知")

    df_merge['款色'] = df_merge['系统款式编码'] + df_merge['颜色'].astype(str)

    df_merge.to_excel(f'D:/桌面/店铺在架/结果输出/{today_small}_店铺在架-系统款式.xlsx', index=False)

    pivot = df_merge.pivot_table(
        index='款色',
        columns=['渠道','店铺名称'],
        values='是否上架',
        aggfunc='max',
        fill_value=0
    )
    # 转成 是/否
    pivot = pivot.replace({1: '是', 0: ''})
    pivot.reset_index(inplace=True)
    pivot['款色'] = pivot['款色'].str.replace('nan', '', regex=False)
    pivot['款色'] = pivot['款色'].str.replace('_x002B_', '+', regex=False)
    # 可选：扁平化列名  渠道_站点名称
    # pivot.columns = [f'{c1}_{c2}' for c1, c2 in pivot.columns]

    pivot.to_excel(f'D:/桌面/店铺在架/结果输出/{today_small}_店铺在架-pivot.xlsx')
    return pivot




if __name__ == "__main__":
    df_huo = huo()
    df_product = df_read_product()
    df_pivot = product_online(df_product)
