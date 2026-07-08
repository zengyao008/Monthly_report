# 整合文件为csv，同时合并重组项目,按剩本剔除相同期数行
import pandas as pd
import os
import glob
import time
import re  # 新增：正则匹配重组后缀


def merge_restruct_contracts(df):
    """
    合并重组合同：将-CZ01/-CZ02/-RA01等后缀的重组合同映射回原始合同，期号连续顺延
    同时剔除重组节点的过渡ADV记录，保证剩余本金与结清判断准确
    """
    df = df.copy()

    # 确保期号为数值型，保证偏移计算正确
    df['期号'] = pd.to_numeric(df['期号'], errors='coerce')

    # ========== 1. 提取原始合同号与重组批次 ==========
    suffix_pattern = r'-(CZ|RA)\d+$'
    df['原始合同号'] = df['合同号'].str.replace(suffix_pattern, '', regex=True, case=False)

    batch_num = df['合同号'].str.extract(r'-(?:CZ|RA)(\d+)$', flags=re.IGNORECASE)[0]
    df['重组批次'] = batch_num.fillna(0).astype(int)

    # ========== 2. 计算每个批次的期号偏移量 ==========
    batch_max_period = df.groupby(['原始合同号', '重组批次'])['期号'].max().reset_index()
    batch_max_period = batch_max_period.sort_values(['原始合同号', '重组批次'])

    batch_max_period['期号偏移'] = (
            batch_max_period.groupby('原始合同号')['期号'].cumsum()
            - batch_max_period['期号']
    )

    df = df.merge(
        batch_max_period[['原始合同号', '重组批次', '期号偏移']],
        on=['原始合同号', '重组批次'],
        how='left'
    )

    # ========== 3. 生成连续期号 + 统一合同号 ==========
    df['期号'] = df['期号'] + df['期号偏移']
    df['合同号'] = df['原始合同号']

    # ========== 新增：剔除重组节点的过渡ADV记录 ==========
    # 步骤A：找出存在多批次的重组合同（有后续重组批次，说明ADV是过渡而非真结清）
    contract_batch_count = df.groupby('原始合同号')['重组批次'].nunique().reset_index()
    restruct_contracts = contract_batch_count[contract_batch_count['重组批次'] > 1]['原始合同号'].tolist()

    if restruct_contracts:
        # 步骤B：对重组合同，找到原始批次（批次=0）中标记为ADV的记录
        mask_restruct_transition = (
                df['原始合同号'].isin(restruct_contracts)
                & (df['重组批次'] == 0)
                & (df['是否结清标志'].astype(str).str.upper() == 'ADV')
        )
        # 步骤C：删除这些过渡ADV
        remove_count = mask_restruct_transition.sum()
        df = df[~mask_restruct_transition]
        print(f"已剔除重组过渡ADV记录：{remove_count} 条")

    # ========== 4. 统一合同级基础信息 ==========
    base_info = (
        df.sort_values('重组批次')
        .groupby('原始合同号')
        .agg(
            统一起租日=('起租日', 'first'),
            统一经销商=('经销商名称', 'first'),
            统一客户=('客户名称', 'first'),
            统一业务类别=('业务类别', 'first'),
            统一大区=('所属大区', 'first')
        )
        .reset_index()
    )
    df = df.merge(base_info, on='原始合同号', how='left')

    df['起租日'] = df['统一起租日']
    df['经销商名称'] = df['统一经销商']
    df['客户名称'] = df['统一客户']
    df['业务类别'] = df['统一业务类别']
    df['所属大区'] = df['统一大区']

    # 同母体合同+同期号，保留未偿还本金最小的最终账务分录
    df = df.sort_values(
        by=["合同号", "期号", "未偿还本金"],
        ascending=[True, True, False]
    )
    df = df.groupby(["合同号", "期号"], as_index=False).tail(1)
    print(f"同一合同同期多条分录清洗完成，清洗后行数：{len(df)}")

    # ========== 5. 清理临时列 ==========
    df = df.drop(columns=['原始合同号', '重组批次', '期号偏移',
                          '统一起租日', '统一经销商', '统一客户', '统一业务类别', '统一大区'])

    print(f"重组合同合并完成，合并后独立合同数：{df['合同号'].nunique()}")
    return df


def merge_excel_to_csv(excel_folder, output_csv, encoding='utf-8-sig',
                       drop_duplicates=True, date_columns=None, numeric_columns=None):
    """
    将文件夹中的所有Excel文件合并为一个CSV文件，并进行数据清洗
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
    required_columns = ['结清日期', '租金结算日期', '合同号', '经销商名称', '起租日', '期限', '客户名称', '业务类别',
                        '所属大区', '期号', '是否结清标志', '未偿还本金']

    for i, file in enumerate(excel_files, 1):
        try:
            with pd.ExcelFile(file, engine='openpyxl') as xls:
                df = pd.read_excel(xls, index_col=None)

            # 筛选必要列
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

    # ========== 数据清洗 ==========
    # 1. 过滤全为NaN的空行
    before_empty_filter = len(combined_df)
    combined_df = combined_df.dropna(how='all')
    empty_rows_removed = before_empty_filter - len(combined_df)
    if empty_rows_removed > 0:
        print(f"已过滤全为空的行: {empty_rows_removed} 行")

    # 2. 转换日期格式（提前转换，方便后续业务逻辑处理）
    if date_columns:
        for col in date_columns:
            if col in combined_df.columns:
                try:
                    combined_df[col] = pd.to_datetime(combined_df[col], errors='coerce')
                    print(f"已将列 '{col}' 转换为日期格式")
                except Exception as e:
                    print(f"转换列 '{col}' 为日期格式失败: {str(e)}")

    # 3. 转换数值格式（提前转换，方便期号偏移计算）
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

    # ========== 新增：合并重组合同（替换原删除逻辑） ==========
    if '合同号' in combined_df.columns and '期号' in combined_df.columns:
        combined_df = merge_restruct_contracts(combined_df)
    else:
        print("警告：缺少合同号或期号列，跳过重组合同合并")

    # 4. 删除完全重复的行
    if drop_duplicates:
        before_dedup = len(combined_df)
        combined_df = combined_df.drop_duplicates()
        duplicate_rows = before_dedup - len(combined_df)
        if duplicate_rows > 0:
            print(f"已删除重复行: {duplicate_rows} 行")

    # 5. 处理缺失值统计
    missing_values = combined_df.isnull().sum()
    missing_columns = [col for col, count in missing_values.items() if count > 0]
    if missing_columns:
        print("\n存在缺失值的列:")
        for col in missing_columns:
            print(f"  {col}: {missing_values[col]} 个缺失值")

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

    date_cols = ['结清日期', '租金结算日期', '起租日']
    numeric_cols = ['未偿还本金', '期限', '期号']

    # 执行合并和清洗
    merge_excel_to_csv(
        excel_folder=excel_folder_path,
        output_csv=output_csv_path,
        drop_duplicates=True,
        date_columns=date_cols,
        numeric_columns=numeric_cols
    )
