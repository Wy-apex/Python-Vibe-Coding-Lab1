

import pandas as pd

# 第一步：读取原始 Excel 文件

from pathlib import Path
import sys

# 检查原始销售文件是否存在
file_path = Path("sales_raw.xlsx")

if not file_path.is_file():
    print("错误：找不到 sales_raw.xlsx 文件，请检查文件是否存在！")
    sys.exit(1)

# 文件存在后，再读取 Excel
df = pd.read_excel(file_path)

# 检查 Excel 是否包含全部必要字段
required_columns = ["日期", "销售员", "产品", "数量", "单价", "地区"]

missing_columns = [
    col for col in required_columns if col not in df.columns
]

if missing_columns:
    print("错误：Excel 文件缺少必要字段：" + "、".join(missing_columns))
    sys.exit(1)

print("字段检查通过，Excel 包含全部必要字段！")

# 检查 Excel 是否没有销售数据
if df.empty:
    print("错误：Excel 文件为空，没有销售数据！")
    sys.exit(1)


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

# 检查数据清洗后是否还有有效记录
if df.empty:
    print("错误：数据清洗后没有有效销售记录！")
    sys.exit(1)

# 第五步：计算每条记录的销售额
df = df.copy()
df["销售额"] = df["数量"] * df["单价"]

# 第六步：输出清洗结果
print("\n清洗后的数据：")
print(df)
print("清洗后有效数据行数：", len(df))
print("销售总额：", df["销售额"].sum(), "元")

# 第七步：按照销售员统计销售额
salesperson_summary = (
    df.groupby("销售员")["销售额"].sum().reset_index()
)

print("\n按销售员统计：")
print(salesperson_summary)

# 第八步：按照地区统计销售额
region_summary = (
    df.groupby("地区")["销售额"].sum().reset_index()
)

print("\n按地区统计：")
print(region_summary)

# 第九步：计算月度销售总额
df["日期"] = pd.to_datetime(df["日期"])
df["月份"] = df["日期"].dt.strftime("%Y-%m")

monthly_summary = (
    df.groupby("月份")["销售额"].sum().reset_index()
)

print("\n月度销售额统计：")
print(monthly_summary)

# 第十步：生成 Excel 销售报表
with pd.ExcelWriter("report.xlsx", engine="openpyxl") as writer:

    # 工作表1：清洗后的销售明细
    df.to_excel(writer, sheet_name="明细清洗后", index=False)

    # 工作表2：按照销售员统计
    salesperson_summary.to_excel(
        writer, sheet_name="按销售员", index=False
    )

    # 工作表3：按照地区统计
    region_summary.to_excel(
        writer, sheet_name="按地区", index=False
    )

    # 工作表4：销售总览
    monthly_summary.to_excel(
        writer, sheet_name="总览", index=False
    )

print("\nExcel 报表生成成功！")
print("文件名称：report.xlsx")
