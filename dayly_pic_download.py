import msoffcrypto
import pandas as pd
from io import BytesIO
import re
import time
from pathlib import Path
from datetime import datetime
import numpy as np
from pic_download import download_and_compress_image_plus


# 获取当天日期
today_small = datetime.now().strftime("%y.%m.%d")
print(today_small)

CWD = f'D:/桌面/过往货盘表/'

p = Path(CWD)
file_path = list(p.glob(f'*{today_small}*.xlsx'))[0]
print(file_path)

def product():
    # 获取当天日期
    today = datetime.now().strftime("%Y-%m-%d")
    print(today)

    CWD = f'D:/桌面/模板/{today}'

    p = Path(CWD)

    file_product = [i for i in list(p.glob('*商品资料*.xlsx')) if '库存视角' not in str(i)]
    print("file_product:", file_product)

    ## 商品资料
    df_product = pd.read_excel(file_product[0], usecols=["图片","款式编码","颜色"])
    # 创建款色组合列
    df_product['款色'] = df_product['款式编码'] + df_product['颜色']
    df_product_sub = df_product.drop_duplicates(subset=['款色'], keep='first').reset_index(drop=True)

    return df_product_sub


def open_protected_excel_safe(file_path, password=789, sheet_name='Sheet1'):
    """
    安全地打开受密码保护的Excel文件
    
    参数:
    file_path: Excel文件路径
    password: 文件密码
    sheet_name: 工作表名称或索引
    
    返回:
    DataFrame 或 None（如果打开失败）
    """
    try:
        with open(file_path, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            
            # 验证文件是否需要密码
            if office_file.is_encrypted():
                office_file.load_key(password=password)
                
                decrypted = BytesIO()
                office_file.decrypt(decrypted)
                
                # 读取Excel文件
                df = pd.read_excel(decrypted, sheet_name=sheet_name, engine='openpyxl')
                return df
            else:
                print("文件未加密，直接读取")
                return pd.read_excel(file_path, sheet_name=sheet_name)
                
    except Exception as e:
        print(f"打开文件时出错: {e}")
        return None


if __name__ == "__main__":
    password = "789"

    df = open_protected_excel_safe(file_path, password, sheet_name='Sheet1')
    if df is not None:
        print("文件读取成功！")
        # print(df.head())

        kuan_pic = df.iloc[:, [2,7]]
        kuan_pic_filter = kuan_pic[kuan_pic['Unnamed: 7'].isna()]

        kuan_pic_list = kuan_pic_filter['Unnamed: 2'].to_list()[1:]

        print(kuan_pic_list)

        product_df = product()
        product_df_filter = product_df[product_df['款色'].isin(kuan_pic_list)]

        for n, row in product_df_filter.iterrows():
            download_and_compress_image_plus(row['图片'], row['款色'])
            product_df_filter.to_clipboard(index=False)

    else:
        print("文件读取失败，请检查密码或文件路径")