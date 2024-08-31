import os
from openpyxl import load_workbook
import pandas as pd
from tqdm import tqdm

def batch_process_csv_files(folder_path):
    # 获取文件夹中所有的文件
    csv_files = [file_name for file_name in os.listdir(folder_path) if file_name.lower().endswith('.csv')]

    for file_name in tqdm(csv_files, desc="processing pivot job", unit="file"):
        # 只处理CSV文件
        if file_name.lower().endswith('.csv'):
            # 读取CSV文件为DataFrame
            file_path = os.path.join(folder_path, file_name)
            df = pd.read_csv(file_path)


            # 调用创建透视表的子程序
            pivot_tables = create_pivot_tables(df)

            # 保存工作簿为Excel文件，避免多表结构丢失
            
            output_file_path = os.path.join(output_path, file_name.replace('.csv', '.xlsx'))

            #temp_file_path = file_path.replace('.csv', '.xlsx')
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='OriginalData', index=False)
                for i, pivot_table in enumerate(pivot_tables):
                    pivot_table.to_excel(writer, sheet_name=f'PivotTable_{i+1}', index=False)

            # 将每个工作表保存为新的sheet
            wb = load_workbook(output_file_path)
            wb.save(output_file_path)

def create_pivot_tables(df):
    # 产品过滤数组
    product_filter_array = ["1", "2", "5", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", 
                            "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", 
                            "37", "40", "50", "74", "92", "93"]

    df['product'] = df['product'].astype(str)
    df_filtered = df[df['product'].str.match(r'^\d{1,2}$')].copy()
   
    # 创建一个透视表列表
    pivot_tables = []

    # 动态生成透视表从 v1967 到 v2020
    for year in range(1967, 2021):  # 从1967到2020
        value_column = f'v{year}'  # 动态生成列名

        # 确认列名存在于数据框中
        if value_column in df_filtered.columns:
            # 确保要汇总的列是数值类型，强制转换
            df_filtered[value_column] = pd.to_numeric(df_filtered[value_column], errors='coerce')

            # 创建透视表
            pivot_table = pd.pivot_table(
                df_filtered,
                index='product',  # 行索引
                columns='exporter',  # 列索引
                values=value_column,  # 数据值
                aggfunc='sum',  # 使用求和函数
                fill_value=0  # 缺失值填充为0
            ).reset_index()
            pivot_table = pivot_table.apply(pd.to_numeric, errors='coerce').fillna(0)

            # 将数据值转换为行总和的百分比
            pivot_table_percentage = pivot_table.div(pivot_table.sum(axis=1), axis=0).fillna(0) * 100
            # 重置索引以方便保存到 Excel
            pivot_table_percentage = pivot_table_percentage.reset_index()

            # 在透视表第一列添加 'Year' 列，第二列标识 'Exporter'
            pivot_table_percentage.insert(0, 'Year', value_column)  # 插入年列
            pivot_table_percentage.insert(1, 'Index Type', 'exporter')  # 插入标识列

            # 将透视表添加到列表中
            pivot_tables.append(pivot_table_percentage)

    return pivot_tables

def save_worksheet_as·_csv(df, output_path):
    df.to_csv(output_path, index=False)

# 运行批量处理CSV文件的函数
folder_path = "/Users/zxt/Desktop/trade_tmp/"  # 修改成你的文件夹路径
output_path = "/Users/zxt/Desktop/trade_out/"
batch_process_csv_files(folder_path)

