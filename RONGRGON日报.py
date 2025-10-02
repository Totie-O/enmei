import os
import pandas as pd
import datetime
import numpy as np
import xlwings as xw
from datetime import datetime, date, timedelta

def nowtime():
    nowtime = str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    return nowtime

# 任务开始计算时间
start_time = datetime.now()
print('任务开始', start_time)

folder_path_24 = 'D:\\桌面\\日报\\sale\\2024年销售数据\\'

folder_path_25 = 'D:\\桌面\\日报\\sale\\2025年销售数据\\'

def file(table,position):
    # 初始化一个空的DataFrame，用于存储合并后的数据
    combined_df = pd.DataFrame()

    # 遍历文件夹中的文件
    for filename in os.listdir(position):
        # 检查文件名是否包含“每日库存报表”且是Excel文件
        if table in filename and filename.endswith(('.xlsx', '.xls')):
            # 构建文件的完整路径
            file_path = os.path.join(position, filename)

            # 尝试读取Excel文件
            try:
                df = pd.read_excel(file_path)
                print(f"读取文件：{filename}")

                # 将读取到的DataFrame添加到combined_df中
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            except Exception as e:
                print(f"读取文件{filename}时出错：{e}")
    return combined_df



# 2024年业绩总销售（全部）
end_time = datetime.now()
data_24 = file('销售主题分析_明细订单商品',folder_path_24)
total_time = end_time - start_time
mes_24 = '读取24年销售文件共花时间： ' + str(total_time).split('.')[0]
print(mes_24)
# 所有业绩
SUM_24 = data_24[['销售数量', '实发数量', '实发金额', '销售金额', '销售成本', '实发成本', '销售毛利', '退货数量', '实退数量', '退货金额', '退货成本', '实退成本', '实退金额']].sum().to_frame().T


# 2024年业绩总销售（剔除异常值）
data_24_new=data_24.copy()
# data_24_new=data_24_new[data_24_new['销售数量']<500]
data_24_new=data_24_new[~data_24_new['产品分类'].isin(['邮费','辅料'])]
# 所有业绩（剔除）['邮费','辅料'],单日单款销售500件以上
data_24_new_sum = data_24_new[['销售数量', '实发数量', '实发金额', '销售金额', '销售成本', '实发成本', '销售毛利', '退货数量', '实退数量', '退货金额', '退货成本', '实退成本', '实退金额']].sum().to_frame().T
data_24_new['付款日期'] = pd.to_datetime(data_24_new['付款日期']).dt.strftime('%Y-%m-%d')


# 2025年业绩总销售（全部）
end_time = datetime.now()
data_25 = file('销售主题分析_明细订单商品',folder_path_25)
total_time = end_time - start_time
mes_25 = '读取25年销售文件共花时间： ' + str(total_time).split('.')[0]
print(mes_25)
# 所有业绩
SUM_25 = data_25[['销售数量', '实发数量', '实发金额', '销售金额', '销售成本', '实发成本', '销售毛利', '退货数量', '实退数量', '退货金额', '退货成本', '实退成本', '实退金额']].sum().to_frame().T



# 2025年业绩总销售（剔除异常值）
data_25_new=data_25.copy()
# data_25_new=data_25_new[data_25_new['销售数量']<500]
data_25_new=data_25_new[~data_25_new['产品分类'].isin(['邮费','辅料'])]
# 所有业绩（剔除）['邮费','辅料'],单日单款销售500件以上
data_25_new_sum = data_25_new[['销售数量', '实发数量', '实发金额', '销售金额', '销售成本', '实发成本', '销售毛利', '退货数量', '实退数量', '退货金额', '退货成本', '实退成本', '实退金额']].sum().to_frame().T

data_25['付款日期'] = pd.to_datetime(data_25['付款日期']).dt.strftime('%Y-%m-%d')
data_25['达人编号'] = data_25['达人编号'].astype(str).str.replace(" ", "")


# data_25['年月'] = data_25['付款日期'].str[:7]
#
# aaa = data_25.groupby(['年月'], as_index=True)[['销售数量','销售金额',]].sum().reset_index()
#
# aaa = data_25_new_sum['销售金额'].sum()
# print(aaa)






#店铺渠道
QD = pd.read_excel('D:\\桌面\\日报\\模块&店铺_对照表.xlsx',sheet_name='渠道划分')

#UID 模块
MK = pd.read_excel('D:\\桌面\\日报\\模块&店铺_对照表.xlsx',sheet_name='最总整合UID')
MK['UID'] = MK['UID'].astype(str).str.replace(" ", "")

MK = MK.drop_duplicates(subset=['UID'], keep='first')

# 匹配渠道
data_1=pd.merge(data_25,QD,on='店铺',how='left')


# 根据UID匹配模块
data_1=pd.merge(data_1,MK[['UID','项目','模块']],left_on='达人编号',right_on='UID',how='left')
# 如果项目为空则等于渠道
data_1['项目'] = data_1['项目'].fillna(data_1['渠道'])

data_1['达人编号'] = pd.to_numeric(data_1['达人编号'], errors='coerce')  # 非数字转NaN



# # 外部切片公司
# qiepian=data_1[data_1['模块'].isin(['抖音切片','快手切片','视频号切片'])]
# #输出excel
#
# qiepian.to_excel(r'C:\Users\PC\Desktop\内部切片明细.xlsx',index=False)


# 定义条件和对应的赋值
conditions = [
    data_1['达人编号'].isna() & data_1['模块'].isna(),   # 条件1：达人编号空 & 模块空
    data_1['达人编号'].notna() & data_1['模块'].isna()  # 条件2：达人编号非空 & 模块空
]
choices = [
    '商品卡',   # 条件1成立时赋值
    '其它UID'   # 条件2成立时赋值
]
# 执行条件赋值
data_1['模块'] = np.select(conditions, choices, default=data_1['模块'])

# 求和
df = data_1.groupby(['项目', '模块'], as_index=True)[['销售数量', '实发数量', '实发金额', '销售金额', '销售成本', '实发成本', '销售毛利', '退货数量', '实退数量', '退货金额', '退货成本', '实退成本', '实退金额']].sum().reset_index()
df['辅助列'] =df['项目'].astype(str)+df['模块'].astype(str)


# df.to_excel(r'C:\Users\PC\Desktop\cs.xlsx',index=False)




# aaaa=data_1[['店铺','付款日期','销售数量','销售金额','达人编号','达人名称','渠道','UID','项目','模块']]
# aaaa.to_excel(r'C:\Users\PC\Desktop\商品卡明细.xlsx',index=False)
# # 外部切片公司
# qiepian=data_1[data_1['模块'].isin(['商品卡'])]
# #输出excel
# qiepian.to_excel(r'C:\Users\PC\Desktop\商品卡明细.xlsx',index=False)


# 溶溶时尚旗舰店（外部合作店铺）
# 商品卡数据

# 获取系统当前日期减1天的年月
yesterday_year_month = (datetime.now() - timedelta(days=1)).strftime('%Y-%m')



hangzhou = data_1.loc[(data_1['店铺'] == '溶溶时尚旗舰店（外部合作店铺）') & (data_1['模块'] == '商品卡')]
hangzhou1 = hangzhou.groupby(['付款日期'], as_index=True)[['销售数量', '销售金额']].sum().reset_index()
hangzhou1['年月'] = hangzhou1['付款日期'].str[:7].fillna('')

# 先将付款日期列转换为 datetime 类型
hangzhou1['付款日期'] = pd.to_datetime(hangzhou1['付款日期'])

# 然后使用 dt 访问器
hangzhou1['是否当月'] = ''
hangzhou1.loc[hangzhou1['付款日期'].notna() & (hangzhou1['付款日期'].dt.strftime('%Y-%m') == yesterday_year_month), '是否当月'] = '当月'




# 抖音主店日业绩
# 商品卡数据
douyin=data_1[(data_1['店铺'].isin(['溶溶女装旗舰店1','溶溶官方旗舰店'])) & (data_1['模块'] == '商品卡')]

douyin1 = douyin.groupby(['付款日期', '店铺'], as_index=True)[['销售数量', '销售金额']].sum().reset_index()

douyin1['年月'] = douyin1['付款日期'].str[:7].fillna('')


# 先将付款日期列转换为 datetime 类型
douyin1['付款日期'] = pd.to_datetime(douyin1['付款日期'])

# 然后使用 dt 访问器
douyin1['是否当月'] = ''
douyin1.loc[douyin1['付款日期'].notna() & (douyin1['付款日期'].dt.strftime('%Y-%m') == yesterday_year_month), '是否当月'] = '当月'






#---------------------------------------------------------------- 有新添加项目（在这里加）
# 1. 定义渠道优先级
specified_order = ['溶总溶总_抖音直播','溶总溶总_快手双开直播','溶总溶总_视频号双开直播','小晶小晶_抖音直播','小晶小晶_快手双开直播','小晶小晶_视频号双开直播','苏苏苏苏_抖音直播','米米_抖音直播','抖音商品卡','抖音抖音切片','抖音其它UID','抖音多来米','抖音有屿','抖音美际','抖音蝉选','抖音麦满分','快手商品卡','快手快手切片','快手其它UID','视频号商品卡','视频号视频号切片','百度推广商品卡','百度推广其它UID','小红书商品卡','唯品会商品卡','有赞商品卡','淘宝商品卡','天猫商品卡','拼多多商品卡','京东商品卡','线下商品卡','爱奇艺商品卡']

order_mapping = {v: i for i, v in enumerate(specified_order)}

# 2. 按映射排序（缺失值会被排到最后）
df_sum = df.sort_values('辅助列', key=lambda x: x.map(order_mapping))

# del df_sum['辅助列']

df_sum['净销售额']=df_sum['销售金额']-df_sum['退货金额']
df_sum['净销售成本额']=df_sum['销售成本']-df_sum['退货成本']
df_sum['净销售毛利额']=df_sum['净销售额']-df_sum['净销售成本额']




# 按渠道分组并计算合计
grouped = df_sum.groupby('项目', as_index=False).apply(
    lambda x: pd.concat([x, pd.DataFrame({
        '项目': ['合计'],
        '销售数量': [x['销售数量'].sum()],
        '实发数量': [x['实发数量'].sum()],
        '实发金额': [x['实发金额'].sum()],
        '销售金额': [x['销售金额'].sum()],
        '销售成本': [x['销售成本'].sum()],
        '实发成本': [x['实发成本'].sum()],
        '销售毛利': [x['销售毛利'].sum()],
        '退货数量': [x['退货数量'].sum()],
        '实退数量': [x['实退数量'].sum()],
        '退货金额': [x['退货金额'].sum()],
        '退货成本': [x['退货成本'].sum()],
        '实退成本': [x['实退成本'].sum()],
        '实退金额': [x['实退金额'].sum()],
        '净销售额': [x['净销售额'].sum()],
        '净销售成本额': [x['净销售成本额'].sum()],
        '净销售毛利额': [x['净销售毛利额'].sum()]
    })])
).reset_index(drop=True)

date_sum=grouped[['辅助列','项目','模块','销售金额','销售毛利','净销售毛利额','实退金额','实发金额','退货数量','销售数量','退货金额']]

# date_sum = date_sum.rename(columns={'销售金额':'当年销售金额','销售毛利':'当年销售毛利额','净销售毛利额':'当年净销售毛利额','实退金额':'当年实退金额','实发金额':'当年实发金额','退货数量':'当年退货数量','销售数量':'当年销售数量','退货金额':'当年退货金额'})

date_sum = date_sum.rename(columns={'净销售毛利额':'25年净销售毛利额','销售数量':'25年销售数量','实发数量':'25年实发数量','实发金额':'25年实发金额','销售金额':'25年销售金额','销售成本':'25年销售成本','实发成本':'25年实发成本','销售毛利':'25年销售毛利','退货数量':'25年退货数量','实退数量':'25年实退数量','退货金额':'25年退货金额','退货成本':'25年退货成本','实退成本':'25年实退成本','实退金额':'25年实退金额'})




date_sum['25年累计毛利率']=''
date_sum['25年累计退货率']=''
date_sum['25年累计发货前退货率']=''
date_sum['25年累计实发退货率']=''
date_sum['25年累计退款率']=''

date_sum['25年目标销额']=''
date_sum['25年累计销额占比']=''
date_sum['25年累计退款率']=''
date_sum['25年度完成进度']=''
date_sum['当月销售额占比']=''
date_sum['当月销售目标额']=''
date_sum['当月完成进度']=''
date_sum['当月进度差']=''
date_sum['当月进度差额']=''





# 当月数据
# 当前日期（含时间部分）
today = pd.to_datetime(datetime.now())  # 或 datetime.today()

today_1 = today.strftime('%Y-%m-%d')
# 昨天日期（字符串格式）
yesterday = (today - timedelta(days=1)).strftime('%Y-%m-%d')
# 当月第一天（直接基于 today 计算，避免依赖 yesterday）
this_month_start = pd.to_datetime(datetime(today.year, today.month, 1)).strftime('%Y-%m-%d')
# 前7天
yesterday_7 = (today - timedelta(days=7)).strftime('%Y-%m-%d')

# 昨天日期（保持为 datetime 对象）
yesterday_dt = today - timedelta(days=1)
# 昨天日期字符串格式
yesterday_str = yesterday_dt.strftime('%Y-%m-%d')

# 昨天所在月份的第一天
yesterday_month_start = pd.to_datetime(datetime(yesterday_dt.year, yesterday_dt.month, 1)).strftime('%Y-%m-%d')


# 当月第一天到昨天销售
yesterd_month_sael=data_1[data_1['付款日期'].between(yesterday_month_start,today_1)]
# 求和
yesterd_month_sael_sum = yesterd_month_sael.groupby(['项目', '模块'], as_index=True)[['销售数量','销售金额','销售毛利','退货金额','实发金额','实退金额']].sum().reset_index()

# # 方法1：使用Categorical分类排序（推荐）
# yesterd_month_sael_sum['渠道'] = pd.Categorical(
#     yesterd_month_sael_sum['渠道'],
#     categories=specified_order + [x for x in yesterd_month_sael_sum['渠道'].unique() if x not in specified_order],
#     ordered=True
# )
# # 先按渠道优先级排序，再按销售金额降序
# yesterd_month_sael_sum = yesterd_month_sael_sum.sort_values(['渠道', '销售金额'], ascending=[True, False])


yesterd_month_sael_sum['辅助列'] =yesterd_month_sael_sum['项目'].astype(str)+yesterd_month_sael_sum['模块'].astype(str)



# # 按渠道分组并计算合计
# yesterd_month_sael_sum_2 = yesterd_month_sael_sum.groupby('渠道', as_index=False).apply(
#     lambda x: pd.concat([x, pd.DataFrame({
#         '渠道': [x['渠道'].iloc[0]],
#         '模块': ['合计'],
#         '销售数量': [x['销售数量'].sum()],
#         '销售金额': [x['销售金额'].sum()],
#     })])
# ).reset_index(drop=True)

yesterd_month_sael_sum_2 = yesterd_month_sael_sum.rename(columns={'销售金额':'当月销售金额','销售数量':'当月销售数量','销售毛利':'当月销售毛利','退货金额':'当月退货金额','实发金额':'当月实发金额','实退金额':'当月实退金额'})

yesterd_month_sael_sum_2['当月件单价']=yesterd_month_sael_sum_2['当月销售金额']/yesterd_month_sael_sum_2['当月销售数量']

# yesterd_month_sael_sum_2 = yesterd_month_sael_sum_2.rename(columns={'模块':'模块'})

del yesterd_month_sael_sum_2['模块']
del yesterd_month_sael_sum_2['项目']





#----------------------------------------------------------------------------------------------------# 当天销售数据
# today_sale = 'D:\\数据文件\\日报\\sale\\当日销售\\'
# today_sale_mx = file('当日销售统计(明细数据)',today_sale)
#
# today_sale_mx=today_sale_mx[today_sale_mx['内部订单号']!='#']
#
#
# today_sale_mx['付款日期'] = pd.to_datetime(today_sale_mx['付款日期']).dt.strftime('%Y-%m-%d')
# today_sale_mx['达人编号'] = today_sale_mx['达人编号'].astype(str).str.replace(" ", "")
#
# # 匹配渠道
# today_sale_mx=pd.merge(today_sale_mx,QD,on='店铺',how='left')
#
# # 根据UID匹配模块
# today_sale_mx=pd.merge(today_sale_mx,MK[['UID','项目','模块']],left_on='达人编号',right_on='UID',how='left')
# # 如果项目为空则等于渠道
# today_sale_mx['项目'] = today_sale_mx['项目'].fillna(data_1['渠道'])
#
# today_sale_mx['达人编号'] = pd.to_numeric(today_sale_mx['达人编号'], errors='coerce')  # 非数字转NaN
#
# # 定义条件和对应的赋值
# conditions_today_sale = [
#     today_sale_mx['达人编号'].isna() & today_sale_mx['模块'].isna(),   # 条件1：达人编号空 & 模块空
#     today_sale_mx['达人编号'].notna() & today_sale_mx['模块'].isna()  # 条件2：达人编号非空 & 模块空
# ]
# choices_today_sale = [
#     '商品卡',   # 条件1成立时赋值
#     '其它UID'   # 条件2成立时赋值
# ]
# # 执行条件赋值
# today_sale_mx['模块'] = np.select(conditions_today_sale, choices_today_sale, default=today_sale_mx['模块'])
#
# # 求和
# today_sale_sum = today_sale_mx.groupby(['项目', '模块'], as_index=True)['销售金额'].sum().reset_index()
# # today_sale_sum['辅助列'] =today_sale_sum['项目'].astype(str)+today_sale_sum['模块'].astype(str)
# today_sale_sum['付款日期'] =today_1
# today_sale_sum=today_sale_sum[['项目', '模块','付款日期','销售金额']]

#----------------------------------------------------------------------------------------------------# 当天销售数据





# 近7天销售
yesterd_month_7_sael=data_1[data_1['付款日期'].between(yesterday_7,today_1)]
yesterd_month_7_sael_sum = yesterd_month_7_sael.groupby(['项目', '模块','付款日期'], as_index=True)[['销售金额']].sum().reset_index()

# --------------------------------------------------------------------------------------------------拼接当天的销售
# yesterd_month_7_sael_sum=pd.concat([yesterd_month_7_sael_sum,today_sale_sum],axis=0)
# --------------------------------------------------------------------------------------------------拼接当天的销售


all_dates = yesterd_month_7_sael_sum['付款日期'].unique()
# sorted_dates = sorted(all_dates)
sorted_dates = sorted(all_dates, reverse=True)# 按日期升序排列

# 创建透视表（按金额聚合）
result = yesterd_month_7_sael_sum.pivot_table(
    index=['项目', '模块'],
    columns='付款日期',
    values='销售金额',  # 改为金额列
    aggfunc='sum',     # 求和
    fill_value=0       # 缺失值填充为0
).reindex(columns=sorted_dates, fill_value=0)  # 确保日期列按升序排列


yesterd_month_7_result_reset = result.reset_index()
yesterd_month_7_result_reset['辅助列'] =yesterd_month_7_result_reset['项目'].astype(str)+yesterd_month_7_result_reset['模块'].astype(str)

del yesterd_month_7_result_reset['模块']
del yesterd_month_7_result_reset['项目']

# # 按渠道分组并计算合计
# yesterd_month_7_result_reset = yesterd_month_7_result_reset.groupby('渠道', as_index=False).apply(
#     lambda x: pd.concat([x, pd.DataFrame({
#         '渠道': [x['渠道'].iloc[0]],
#         '模块': ['合计']
#     })])
# ).reset_index(drop=True)

# yesterd_month_7_result_reset = yesterd_month_7_result_reset.rename(columns={'辅助':'模块'})


# 合并月份数据
date_sum_2=pd.merge(date_sum,yesterd_month_sael_sum_2,on=['辅助列'],how='left')

date_sum_2 = date_sum_2[['辅助列','项目', '模块', '25年目标销额','25年销售数量','25年销售金额','25年累计销额占比','25年度完成进度', '25年销售毛利', '25年净销售毛利额','25年实发金额', '25年实退金额','25年退货数量','25年退货金额','当月销售目标额','当月销售金额', '当月销售额占比', '当月销售数量', '当月件单价','当月销售毛利','当月退货金额','当月实发金额','当月实退金额', '当月完成进度', '当月进度差', '当月进度差额', '25年累计毛利率','25年累计退货率', '25年累计发货前退货率', '25年累计实发退货率', '25年累计退款率']]




# 同比数据
# 方法 1：直接指定年月日

# 24年第一天
str_date_2024 = date(2024, 1, 1).strftime('%Y-%m-%d')

# 去年当天前1天
end_date_2024 = today - timedelta(days=366)
end_date_2024 =end_date_2024.strftime('%Y-%m-%d')


# 获取去年同期数据
data_24_new_1=data_24_new[data_24_new['付款日期'].between(str_date_2024,end_date_2024)]

# 同进度
data_24_new_qj= data_24_new_1[['销售数量', '实发数量', '实发金额', '销售金额', '销售成本', '实发成本', '销售毛利', '退货数量', '实退数量', '退货金额', '退货成本', '实退成本', '实退金额']].sum().to_frame().T


# 匹配渠道
data_24_new_2=pd.merge(data_24_new_1,QD,on='店铺',how='left')

data_24_new_2['达人编号'] = data_24_new_2['达人编号'].astype(str).str.replace(" ", "")

# 根据UID匹配模块
data_24_new_3=pd.merge(data_24_new_2,MK[['UID','项目','模块']],left_on='达人编号',right_on='UID',how='left')

data_24_new_3['达人编号'] = pd.to_numeric(data_24_new_3['达人编号'], errors='coerce')  # 非数字转NaN

# 如果项目为空则等于渠道
data_24_new_3['项目'] = data_24_new_3['项目'].fillna(data_24_new_3['渠道'])


# 定义条件和对应的赋值
conditions = [
    data_24_new_3['达人编号'].isna() & data_24_new_3['模块'].isna(),   # 条件1：达人编号空 & 模块空
    data_24_new_3['达人编号'].notna() & data_24_new_3['模块'].isna()  # 条件2：达人编号非空 & 模块空
]
choices = [
    '商品卡',   # 条件1成立时赋值
    '其它UID'   # 条件2成立时赋值
]
# 执行条件赋值
data_24_new_3['模块'] = np.select(conditions, choices, default=data_24_new_3['模块'])


# 求和
data_24_new_4= data_24_new_3.groupby(['项目', '模块'], as_index=True)[['销售数量', '实发数量', '实发金额', '销售金额', '销售成本', '实发成本', '销售毛利', '退货数量', '实退数量', '退货金额', '退货成本', '实退成本', '实退金额']].sum().reset_index()
data_24_new_4['辅助列'] =data_24_new_4['项目'].astype(str)+data_24_new_4['模块'].astype(str)
data_24_new_4 = data_24_new_4.rename(columns={'销售数量':'24年销售数量','实发数量':'24年实发数量','实发金额':'24年实发金额','销售金额':'24年销售金额','销售成本':'24年销售成本','实发成本':'24年实发成本','销售毛利':'24年销售毛利','退货数量':'24年退货数量','实退数量':'24年实退数量','退货金额':'24年退货金额','退货成本':'24年退货成本','实退成本':'24年实退成本','实退金额':'24年实退金额'})



# 合并同比数据
date_sum_3=pd.merge(date_sum_2,data_24_new_4[['辅助列','24年销售金额','24年销售毛利','24年退货金额','24年实发金额','24年实退金额']],on='辅助列',how='left')


# 合并近7天数据
date_sum_4=pd.merge(date_sum_3,yesterd_month_7_result_reset,on=['辅助列'],how='left')

date_sum_4['销售金额趋势图']=''

del date_sum_4['辅助列']





app = xw.App()
book = app.books.open('D:\\桌面\\日报\\日报业绩_模板.xlsx')
# book_new_name ='D:\\数据文件\\进销存数据源\\单款进销存文件\\单款进销存'+datetime.datetime.now().strftime('%Y-%m-%d')+'.xlsx'

#
sht = book.sheets['整合']
sht.api.AutoFilterMode = False
sht.range('B3').options(index=False, header=True).value = date_sum_4
sht.range('B54').options(index=False, header=True).value = SUM_24
sht.range('B57').options(index=False, header=True).value = SUM_25
sht.range('B60').options(index=False, header=True).value = data_24_new_sum
sht.range('B63').options(index=False, header=True).value = data_25_new_sum
sht.range('B66').options(index=False, header=True).value = data_24_new_qj

sht = book.sheets['溶溶时尚旗舰店（外部合作店铺）']
sht.api.AutoFilterMode = False
sht.range('A1:G1').clear_contents()
sht.range('A1').options(index=False, header=True).value = hangzhou1

sht = book.sheets['抖音主店商品卡数据']
sht.api.AutoFilterMode = False
sht.range('A1:G1').clear_contents()
sht.range('A1').options(index=False, header=True).value = douyin1



# book.save(book_new_name)
book.save()
book.close()
app.quit()


# end_time = datetime.datetime.now()
# total_time = end_time - start_time
# mes = '共花时间： ' + str(total_time).split('.')[0]














