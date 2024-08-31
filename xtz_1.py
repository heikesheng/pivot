import pandas as pd
import os

def create_pivot_tables(df, years, product_filter_list):
    pivot_tables = {}
    for year in years:
        value_field = f"v{year}"
        if value_field not in df.columns:
            continue

        # 创建透视表
        pivot_table = pd.pivot_table(df,
                                     values=value_field,
                                     index='product',
                                     columns='exporter',
                                     aggfunc='sum',
                                     fill_value=0)

        # 将 vXXXX 数据转换为百分比
        pivot_table = pivot_table.div(pivot_table.sum(axis=0), axis=1) * 100

        # 过滤掉不需要的 product
        pivot_table = pivot_table[pivot_table.index.isin(product_filter_list)]
        # 设置 MultiIndex，层级分别是Year和Product
        pivot_table['Year'] = year
        pivot_table.set_index('Year', append=True, inplace=True)
        # 调整顺序，并存入字典
        pivot_tables[year] = pivot_table.reorder_levels(['Year', 'product'])

    return pivot_tables

# 设置文件夹路径
folder_path = "/Users/zxt/Desktop/trade_tmp/"  # 修改成你的文件夹路径

# 产品过滤数组
product_filter_list = ["1", "2", "5", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21",
                       "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35",
                       "36", "37", "40", "50", "74", "92", "93"]

# 指定的年份范围
years = range(1967, 2020 + 1)

# 创建临时存储所有国家数据的字典
all_countries_data = {}

# 遍历文件夹中的每个文件
for file_name in os.listdir(folder_path):
    if file_name.lower().endswith('.csv'):
        # 构建完整文件路径
        file_path = os.path.join(folder_path, file_name)
        # 读取CSV文件
        df = pd.read_csv(file_path)

        # 获取国家名（假设国家名在文件名中）
        country_name = os.path.splitext(file_name)[0]

        # 生成透视表数据
        pivot_tables = create_pivot_tables(df, years, product_filter_list)

        # 存储数据到字典中
        all_countries_data[country_name] = pivot_tables

# 保存结果到一个Excel文件中
with pd.ExcelWriter(os.path.join(folder_path, 'All_Countries_Pivot_Tables.xlsx'), engine='openpyxl') as writer:
    for country, pivot_tables in all_countries_data.items():
        combined_df = pd.DataFrame()
        for year, pt in pivot_tables.items():
            combined_df = pd.concat([combined_df, pt])
        
        # 为每个国家创建一个工作表
        combined_df.index.set_names(['Year', 'Product'], inplace=True)
        combined_df.to_excel(writer, sheet_name=country)

print("数据处理完成！🤗📊")
