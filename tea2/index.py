import pandas 

df = pandas.read_excel('tea.xlsx')

print(df)

# 获取所有类型和编号，去重
categoryDf = df.loc[:, ['类型', '编号']].drop_duplicates()

print(categoryDf)

# 获取到类型、编号数组
category = categoryDf.to_dict(orient ='records' )

print(category)

allDf = None

def calTea(tea, num):
    global allDf
    # 筛选出类型匹配、月份为5月的茶叶表
    teaDf = df[(df['类型'] == tea) & (df['日期'].dt.month == 5)]
    # 计算 斤 的总和
    weight = teaDf['斤'].sum()
    # 计算 金额 的总和
    money = teaDf['金额'].sum()
    # 生成新的一行
    newDf = pandas.DataFrame({
        '类型': tea,
        '编号': num,
        '斤': weight,
        '金额': money
    }, index = [ 0 if allDf is None else len(allDf) ])
    # 如果allDf 为空未第一个表
    if allDf is None:
        allDf = newDf
    else:
    # 合并到allDf表格中
        allDf = allDf.append(newDf)

# calTea('白茶', 201)
# calTea('红茶', 202)
# calTea('毛峰', 203)
# calTea('黄金芽', 204)

# 循环获取茶叶的计算汇总值
for item in category:
    calTea(item['类型'], item['编号'])

# 设置索引为类型
allDf = allDf.set_index('类型')

print(allDf)

# 生成result.xlsx
writer = pandas.ExcelWriter('result.xlsx')
df.to_excel(writer, sheet_name = '5月明细')
allDf.to_excel(writer, sheet_name = '5月汇总')
writer.save()