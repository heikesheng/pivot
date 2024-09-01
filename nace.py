import os
from openpyxl import load_workbook
import pandas as pd
from tqdm import tqdm

def batch_process_csv_files(folder_path):
    # 获取文件夹中所有的文件
    xlsx_files = [file_name for file_name in os.listdir(folder_path) if file_name.lower().endswith('.xlsx')]

    for file_name in tqdm(xlsx_files, desc="processing pivot job", unit="file"):
        # 只处理CSV文件
        if file_name.lower().endswith('.xlsx'):
            # 读取CSV文件为DataFrame
            file_path = os.path.join(folder_path, file_name)
            df = pd.read_excel(file_path)


            # 调用创建透视表的子程序
            pivot_tables = create_pivot_tables(df, file_name)

            # 保存工作簿为Excel文件，避免多表结构丢失
            
            output_file_path = os.path.join(output_path, file_name.replace('.xlsx', '.xlsx'))

            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='OriginalData', index=False)  # 总表命名保持不变
                for _, (sheet_name, pivot_table) in enumerate(pivot_tables):
                    pivot_table.to_excel(writer, sheet_name=sheet_name, index=False)

            # 将每个工作表保存为新的sheet
            wb = load_workbook(output_file_path)
            wb.save(output_file_path)

def save_worksheet_as_csv(df, output_path):
    df.to_csv(output_path, index=False)

def create_pivot_tables(df, file_name):
    # 产品过滤数组
    # perhaps this is nace type
    product_filter_array = ['A', 'B', 'C', 'C10-C12', 'C13-C15', 'C16-C18', 'C19', 'C20', 'C20-C21', 'C21', 'C22-C23', 'C24-C25', 'C26', 'C26-C27', 'C27', 'C28', 'C29-C30', 'C31-C33', 'D', 'D-E', 'E', 'F', 'G', 'G45', 'G46', 'G47', 'H', 'H49', 'H50', 'H51', 'H52', 'H53', 'I', 'J', 'J58-J60', 'J61', 'J62-J63', 'K', 'L', 'L68A', 'M', 'M-N', 'MARKT', 'MARKTxAG', 'N', 'O', 'O-Q', 'P', 'Q', 'Q86', 'Q87-Q88', 'R', 'R-S', 'S', 'T', 'TOT', 'TOT_IND', 'U']


    df['nace_r2_name'] = df['nace_r2_name'].astype(str)
    df_filtered = df[df['nace_r2_name'].isin(product_filter_array)].copy()
   
    # 创建一个透视表列表
    pivot_tables = []

    # 动态生成透视表从 v1967 到 v2020
    for year in range(1967, 2021):  # 从1967到2020
        value_column = f'v{year}'  # 动态生成列名
        base_name = file_name[:3].upper()
        # 确认列名存在于数据框中
        if value_column in df_filtered.columns:
            # 确保要汇总的列是数值类型，强制转换
            df_filtered[value_column] = pd.to_numeric(df_filtered[value_column], errors='coerce')

            # 创建透视表
            pivot_table = pd.pivot_table(
                df_filtered,
                index='nace_r2_name',  # 行索引
                columns='year',  # 列索引
                values='VA_Q',  # 数据值
                aggfunc='sum',  # 使用求和函数
                fill_value=0  # 缺失值填充为0
            ).reset_index()
            cols_to_normalize = pivot_table.columns.difference(['product'])
            row_sums = pivot_table[cols_to_normalize].sum(axis=1)
            
            # 处理行总和为 0 的情况
            zero_sum_rows = row_sums == 0
            pivot_table.loc[~zero_sum_rows, cols_to_normalize] = pivot_table.loc[~zero_sum_rows, cols_to_normalize].div(row_sums[~zero_sum_rows], axis=0)
            pivot_table.loc[zero_sum_rows, cols_to_normalize] = 0

            sheet_name = f"{base_name}{year}"  # 透视子表的命名变更
            pivot_tables.append((sheet_name, pivot_table))
    return pivot_tables

# 运行批量处理CSV文件的函数
folder_path = "/Users/zxt/Desktop/trade_tmp/"  # 修改成你的文件夹路径
output_path = "/Users/zxt/Desktop/trade_out/"
batch_process_csv_files(folder_path)

