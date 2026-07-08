# 整合文件为csv，同时去除重组项目
import pandas as pd
import os
import glob
import time


def merge_excel_to_csv(excel_folder, output_csv, encoding='utf-8-sig',
                       drop_duplicates=True, date_columns=None, numeric_columns=None):
    """
    将文件夹中的所有Excel文件合并为一个CSV文件，并进行数据清洗

    参数:
    excel_folder: 存放多个Excel文件的文件夹路径
    output_csv: 合并后的CSV文件路径
    encoding: 输出CSV的编码格式
    drop_duplicates: 是否删除重复行
    date_columns: 需要转换为日期格式的列名列表
    numeric_columns: 需要转换为数值格式的列名列表
    """
    # 验证输入文件夹是否存在
    if not os.path.exists(excel_folder):
        print(f"错误：文件夹不存在 - {excel_folder}")
        return

    # 获取并排序所有Excel文件
    excel_files = glob.glob(os.path.join(excel_folder, "*.xlsx")) + \
                  glob.glob(os.path.join(excel_folder, "*.xls"))
    excel_files.sort()

    if not excel_files:
        print("错误：未找到Excel文件！")
        return

    print(f"找到{len(excel_files)}个Excel文件，开始处理...")
    start_time = time.time()
    all_data = []  # 存储所有数据

    # 读取所有Excel文件
    # 1. 定义需要保留的列
    required_columns = ['结清日期', '租金结算日期', '合同号', '经销商名称', '起租日', '期限', '客户名称', '业务类别', '所属大区', '期号', '是否结清标志', '未偿还本金']

    for i, file in enumerate(excel_files, 1):
        try:
            with pd.ExcelFile(file, engine='openpyxl') as xls:
                df = pd.read_excel(xls, index_col=None)

            # 新增：筛选必要列（只保留required_columns中存在的列）
            available_cols = [col for col in required_columns if col in df.columns]
            if available_cols:
                df = df[available_cols]
                print(f"文件{os.path.basename(file)}已筛选必要列: {available_cols}")
            else:
                print(f"警告：文件{os.path.basename(file)}无匹配的必要列，保留所有列")

            # 跳过空文件
            if df.empty:
                print(f"警告：文件{os.path.basename(file)}为空，已跳过")
                continue

            print(f"已读取 {i}/{len(excel_files)}: {os.path.basename(file)}，行数: {len(df)}")
            all_data.append(df)

        except Exception as e:
            print(f"处理文件{os.path.basename(file)}时出错: {str(e)}，已跳过")

    # 合并所有数据
    if not all_data:
        print("错误：没有有效数据可合并！")
        return

    combined_df = pd.concat(all_data, ignore_index=True)
    original_rows = len(combined_df)
    print(f"\n合并后原始数据总行数: {original_rows}")

    # 数据清洗步骤中，新增：
    # 1. 过滤全为NaN的空行
    before_empty_filter = len(combined_df)
    combined_df = combined_df.dropna(how='all')  # 只删除“所有列都为NaN”的行
    empty_rows_removed = before_empty_filter - len(combined_df)
    if empty_rows_removed > 0:
        print(f"已过滤全为空的行: {empty_rows_removed} 行")
    # 1. 过滤合同号中包含-CZ或-RA的记录
    if '合同号' in combined_df.columns:
        # 统计需要过滤的记录数
        before_filter = len(combined_df)
        # 使用使用str.contains进行模糊匹配，case=False表示不区分大小写
        combined_df = combined_df[~combined_df['合同号'].astype(str).str.contains('-CZ|-ra', case=False, na=False)]
        filtered_rows = before_filter - len(combined_df)
        print(f"已过滤包含-CZ或-RA的合同记录: {filtered_rows} 行")
    else:
        print("警告：数据中未找到'合同号'列，无法进行合同号过滤")

    # 2. 删除完全重复的行
    if drop_duplicates:
        combined_df = combined_df.drop_duplicates()
        duplicate_rows = original_rows - len(combined_df) - (filtered_rows if '合同号' in combined_df.columns else 0)
        if duplicate_rows > 0:
            print(f"已删除重复行: {duplicate_rows} 行")

    # 3. 处理缺失值（标记而非删除，避免数据丢失）
    missing_values = combined_df.isnull().sum()
    missing_columns = [col for col, count in missing_values.items() if count > 0]
    if missing_columns:
        print("\n存在缺失值的列:")
        for col in missing_columns:
            print(f"  {col}: {missing_values[col]} 个缺失值")

    # 4. 转换日期格式
    if date_columns:
        for col in date_columns:
            if col in combined_df.columns:
                try:
                    combined_df[col] = pd.to_datetime(combined_df[col], errors='coerce')
                    print(f"已将列 '{col}' 转换为日期格式")
                except Exception as e:
                    print(f"转换列 '{col}' 为日期格式失败: {str(e)}")

    # 5. 转换数值格式
    if numeric_columns:
        for col in numeric_columns:
            if col in combined_df.columns:
                try:
                    # 先移除可能存在的千位分隔符
                    combined_df[col] = combined_df[col].replace({',': ''}, regex=True)
                    combined_df[col] = pd.to_numeric(combined_df[col], errors='coerce')
                    print(f"已将列 '{col}' 转换为数值格式")
                except Exception as e:
                    print(f"转换列 '{col}' 为数值格式失败: {str(e)}")

    # 保存清洗后的结果
    combined_df.to_csv(output_csv, index=False, encoding=encoding)

    # 输出处理结果
    elapsed_time = time.time() - start_time
    print(f"\n处理完成！清洗后总行数: {len(combined_df)}")
    print(f"文件已保存至: {output_csv}")
    print(f"总耗时: {elapsed_time:.2f}秒")

    return combined_df


if __name__ == "__main__":
    # 配置路径
    excel_folder_path = r"D:\Filelist\Python_list\File_list\系统导出的Excel文件"
    output_csv_path = r"D:\Filelist\Python_list\File_list\汇总的租金收入表.csv"

    # 根据你的实际列名修改以下配置
    # 主程序中修改：匹配required_columns中的日期/数值列
    date_cols = ['结清日期', '租金结算日期', '起租日']  # 所有日期类型列
    numeric_cols = ['未偿还本金', '期限', '期号']  # 所有数值类型列（期号、期限也是数值）

    # 执行合并和清洗
    merge_excel_to_csv(
        excel_folder=excel_folder_path,
        output_csv=output_csv_path,
        drop_duplicates=True,  # 开启去重
        date_columns=date_cols,
        numeric_columns=numeric_cols
    )
