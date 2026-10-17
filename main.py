

import pandas as pd

# 第一步：读取原始 Excel 文件
df = pd.read_excel("sales_raw.xlsx")

print("原始数据：")
print(df)
print("原始数据行数：", len(df))

# 第二步：删除关键字段缺失的记录
required_columns = ["日期", "销售员", "产品", "数量", "单价", "地区"]
df = df.dropna(subset=required_columns)

# 第三步：删除数量小于或等于 0 的记录
df = df[df["数量"] > 0]

# 第四步：按照全部字段删除重复记录
df = df.drop_duplicates()

# 第五步：计算每条记录的销售额
df = df.copy()
df["销售额"] = df["数量"] * df["单价"]

# 第六步：输出清洗结果
print("\n清洗后的数据：")
print(df)
print("清洗后有效数据行数：", len(df))
print("销售总额：", df["销售额"].sum(), "元")
