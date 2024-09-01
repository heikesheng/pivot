import os
import pandas as pd
from openpyxl import load_workbook
from tqdm import tqdm

def batch_process_csv_files(folder_path, output_path):
    csv_files = [file_name for file_name in os.listdir(folder_path) if file_name.lower().endswith('.csv')]

    for file_name in tqdm(csv_files, desc="Processing pivot job", unit="file"):
        if file_name.lower().endswith('.csv') and not file_name.lower().startswith('Pivot'):
            file_path = os.path.join(folder_path, file_name)
            try:
                df = pd.read_csv(file_path)
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding='latin1')

            if 'product' not in df.columns:
                continue
            
            pivot_tables = create_pivot_tables(df, file_name)

            output_file_path = os.path.join(output_path, file_name.replace('.csv', '.xlsx'))

            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='OriginalData', index=False)  # 总表命名保持不变
                for _, (sheet_name, pivot_table) in enumerate(pivot_tables):
                    pivot_table.to_excel(writer, sheet_name=sheet_name, index=False)

            wb = load_workbook(output_file_path)
            wb.save(output_file_path)

def create_pivot_tables(df, file_name):
    product_filter_array = [
        "01", "02", "05", "10", "11", "12", "13", "14", "15", "16", "17", "18", 
        "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", 
        "31", "32", "33", "34", "35", "36", "37", "40", "50", "74", "92", "93"
    ]
    df['product'] = df['product'].astype(str)
    df_filtered = df[df['product'].isin(product_filter_array)].copy()

    pivot_tables = []
    base_name = file_name[:3].upper()

    for year in range(1967, 2021):
        value_column = f'v{year}'

        if value_column in df_filtered.columns:
            df_filtered[value_column] = pd.to_numeric(df_filtered[value_column], errors="coerce")

            pivot_table = pd.pivot_table(
                df_filtered,
                index='product',
                columns='exporter',
                values=value_column,
                aggfunc='sum',
                fill_value=0
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

# 设置文件夹路径
folder_path = "/Users/zxt/Desktop/trade_remain/"
output_path = "/Users/zxt/Desktop/trade_out_final/"
batch_process_csv_files(folder_path, output_path)
