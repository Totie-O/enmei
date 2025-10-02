sort = ["款式编码","颜色","规格","品牌","商品编码","名称","数量","图片地址链接","订单占有","仓库待发","安全库存下限","安全库存上限","最小备货天数","最大备货天数","采购在途数","销退仓库存","进货仓库存","虚拟库存","次品库存","调拨在途数","库存锁定","运营云仓可用数","样衣仓","办公室版衣仓","仓库报废仓","待清洗仓","可用数","公有可用数"]

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

file_psi = list(p.glob('*商品库存*.xlsx'))

print("file_today:", file_psi)

df_psi = pd.read_excel(file_psi[0])

print(df_psi.columns)
df = df_psi.reindex(columns=sort)
df['款式编码'] = df['款式编码'].str.replace('_x002B_', '+')
df['商品编码'] = df['商品编码'].str.replace('_x002B_', '+')
df['颜色'] = df['颜色'].str.replace('_x002B_', '+')

df.to_excel(f'D:/桌面//模板/{today}/{today}商品库存new.xlsx', index=False)