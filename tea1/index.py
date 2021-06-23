import pandas
import numpy
from datetime import datetime
import time

# 读入表格数据
df = pandas.read_excel('tea.xlsx')

# 打印表格数据
print(df)

allDf = None

def calTea(tea, num):
    global allDf

    # 我们计算5月红茶卖了多少斤，卖了多少钱
    weight = df[(df['类型'] == tea) & (df['日期'].dt.month == 5)]['斤'].sum()

    money = df[(df['类型'] == tea) & (df['日期'].dt.month == 5)]['金额'].sum()

    print(weight)
    print(money)

    # 生成行
    weightDf = pandas.DataFrame({
        # '日期': '',
        '': '5月' + tea + '总计卖出',
        '类型': tea,
        '编号': num,
        '斤': weight,
        '金额': money,
    }, index = [ 0 if allDf is None else len(allDf) ])
    
    if allDf is None:
        allDf = weightDf
    else:
        allDf = allDf.append(weightDf)
    

calTea('白茶', 201)
calTea('红茶', 202)
calTea('毛峰', 203)
calTea('黄金牙', 204)

# 拼接在现有数据后面
# df = df.append(weightDf)

allDf = allDf.set_index([''])

print(allDf)

# 写入result.xlsx
writer = pandas.ExcelWriter('result.xlsx')
allDf.to_excel(writer)
writer.save()











# s_date = datetime.datetime.strptime('20210501', "%Y%m%d").date()
# s_date = time.mktime(time.strptime('20210501', "%Y%m%d"))
# print(s_date)
# e_date = time.mktime(time.strptime('20210601', "%Y%m%d"))
# print(e_date)
# weight = df[(df['类型'] == '红茶') & (time.mktime(time.strptime(df['日期'].str, "%Y%m%d")) > s_date) & (time.mktime(time.strptime(df['日期'].str, "%Y%m%d")) < e_date)]['斤'].sum()
# money = df[(df['类型'] == '红茶') & (df['日期'] > s_date) & (df['日期'] < e_date)]['金额'].sum()