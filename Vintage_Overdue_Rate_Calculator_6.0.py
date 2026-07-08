# 当前代码可运行，检验基本可行，后续基于此做调整，目前使用此段代码
import pandas as pd
import warnings
import logging
from openpyxl.styles import PatternFill

# ============================================
# 0. 配置参数与日志初始化
# ============================================
CONFIG = {
    # 输入文件路径
    "input_files": {
        "rent": "D:\\Filelist\\Python_list\\File_list\\汇总的租金收入表.csv",
        "asset": "D:\\Filelist\\Python_list\\历史系统数据\\资产余额表06.30_处理后（保留格式）.xlsx"
    },
    # 输出文件路径
    "output_file": "D:\\Filelist\\Python_list\\Vintage15+.商用车.损失不担.去重组.0630.xlsx",
    # 经销商筛选条件
    "dealer_filters": [],
    # 业务参数
    "business_params": {
        "overdue_dpd_threshold": 16,
        "first_period_num": 1
    },
    # 必要列配置
    "required_columns": {
        "merge": ["合同号", "经销商名称", "客户名称", "起租日", "期号", "租金结算日期", "结清日期", "未偿还本金",
                  "放款金额", "业务类别", "业务模式"],
        "preprocess": ["合同号", "期号", "放款金额", "起租日", "逾期天数（DPD）", "未偿还本金"]
    }
}


def init_logger():
    """初始化日志配置，同时输出到控制台和文件"""
    log_format = "%(asctime)s - %(levelname)s - %(module)s - %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.FileHandler("vintage_analysis.log", encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)


# 初始化日志对象
logger = init_logger()


# ============================================
# 1. 数据提取
# ============================================
def merge_excel_files():
    """合并租金收入表和资产余额表，处理日期字段并计算逾期天数（DPD）"""
    try:
        # 读取租金收入表
        rent_df = pd.read_csv(
            CONFIG["input_files"]["rent"],
            usecols=['合同号', '经销商名称', '起租日', '期号', '租金结算日期', '结清日期', '未偿还本金', '业务类别']
        )
        logger.info(f"成功读取租金收入表，数据行数：{len(rent_df)}")

        # 读取资产余额表
        asset_df = pd.read_excel(
            CONFIG["input_files"]["asset"],
            usecols=['合同号', '客户名称', '放款金额', '业务模式', '大区', '业务来源']
        )
        logger.info(f"成功读取资产余额表，数据行数：{len(asset_df)}")

        # 合并数据
        merged_df = pd.merge(rent_df, asset_df, on='合同号', how='inner')
        logger.info(f"两表inner合并后，数据行数：{len(merged_df)}")

        # 1. 标准化日期字段（同原逻辑，不变）
        merged_df['租金结算日期'] = pd.to_datetime(merged_df['租金结算日期'], errors='coerce')
        merged_df['结清日期'] = pd.to_datetime(merged_df['结清日期'], errors='coerce')

        # 2. 获取程序运行的当前时间（仅保留日期，去掉时分秒，避免跨天误差）
        current_date = pd.Timestamp.now().normalize()  # 例如：2025-10-02 00:00:00

        # 3. 分场景计算逾期天数（DPD）
        def calculate_dpd(row, current_date):
            """
            按场景计算单条记录的逾期天数：
            - 场景1：结清日期非空 → 用结清日期 - 结算日期
            - 场景2：结清日期为空 → 若当前时间>结算日期，用当前时间-结算日期；否则未逾期（0）
            - 场景3：结算日期无效（NaT） → 无法计算（pd.NA）
            """
            # 先判断结算日期是否有效（无效则返回NA）
            if pd.isna(row['租金结算日期']):
                return pd.NA

            # 场景1：已结清（结清日期非空）- 修正：去掉.dt
            if pd.notna(row['结清日期']):
                dpd = (row['结清日期'] - row['租金结算日期']).days
                return dpd if pd.notna(dpd) else pd.NA  # 极端情况（如日期差无效）返回NA

            # 场景2：未结清（结清日期为空）
            else:
                # 2.1 已逾期：当前时间 > 结算日期
                if current_date > row['租金结算日期']:
                    dpd = (current_date - row['租金结算日期']).days  # 正确，无需修改
                    return dpd
                # 2.2 未逾期：当前时间 ≤ 结算日期（未到还款期）
                else:
                    return 0

        # 4. 应用函数计算所有记录的DPD
        merged_df['逾期天数（DPD）'] = merged_df.apply(
            lambda row: calculate_dpd(row, current_date),
            axis=1  # axis=1表示按“行”应用函数
        )

        return merged_df
    except FileNotFoundError as e:
        logger.error(f"文件未找到: {e}", exc_info=True)
    except KeyError as e:
        logger.error(f"列名错误: {e}", exc_info=True)
    except Exception as e:
        logger.error(f"数据提取未知错误: {e}", exc_info=True)
    return None

# ============================================
# 2. 数据校验
# ============================================
def validate_data(df):
    """数据合理性校验，返回校验通过的数据并记录异常"""
    df_valid = df.copy()
    invalid_records = []

    # 数值型字段校验
    if (df_valid['放款金额'] <= 0).any():
        invalid_cnt = (df_valid['放款金额'] <= 0).sum()
        invalid_records.append(f"放款金额≤0的记录数：{invalid_cnt}")
        df_valid = df_valid[df_valid['放款金额'] > 0]

    if (df_valid['期号'] <= 0).any():
        invalid_cnt = (df_valid['期号'] <= 0).sum()
        invalid_records.append(f"期号≤0的记录数：{invalid_cnt}")
        df_valid = df_valid[df_valid['期号'] > 0]

    if (df_valid['未偿还本金'] < 0).any():
        invalid_cnt = (df_valid['未偿还本金'] < 0).sum()
        invalid_records.append(f"未偿还本金<0的记录数：{invalid_cnt}")
        df_valid = df_valid[df_valid['未偿还本金'] >= 0]

    # 日期逻辑校验
    df_valid['起租日'] = pd.to_datetime(df_valid['起租日'], errors='coerce')
    df_valid['租金结算日期'] = pd.to_datetime(df_valid['租金结算日期'], errors='coerce')
    date_invalid = df_valid[(df_valid['起租日'].notna()) & (df_valid['租金结算日期'].notna()) &
                            (df_valid['起租日'] > df_valid['租金结算日期'])]
    if len(date_invalid) > 0:
        invalid_records.append(f"起租日晚于结算日期的记录数：{len(date_invalid)}")
        df_valid = df_valid.drop(date_invalid.index)

    # 输出校验结果日志
    if invalid_records:
        logger.warning("数据校验发现异常：")
        for record in invalid_records:
            logger.warning(f"  - {record}")
    else:
        logger.info("数据校验通过，无异常记录")

    return df_valid


# ============================================
# 3. 数据预处理
# ============================================
def preprocess_data(df, filter_params=None):
    """数据预处理，支持多维度筛选并计算逾期相关字段"""
    # 数据合理性校验
    df = validate_data(df)

    # 多维度筛选
    if filter_params:
        for filter_col, filter_conditions in filter_params.items():
            if filter_col not in df.columns:
                logger.warning(f"筛选列{filter_col}不存在，跳过该维度筛选")
                continue

            mask = pd.Series([True] * len(df), index=df.index)
            for condition in filter_conditions:
                if 'not ' in condition:
                    keyword = condition.replace('not ', '').strip()
                    mask &= ~df[filter_col].astype(str).str.contains(keyword, na=False, regex=False)
                    logger.info(f"筛选{filter_col}：不包含{keyword}")
                else:
                    keyword = condition.strip()
                    mask &= df[filter_col].astype(str).str.contains(keyword, na=False, regex=False)
                    logger.info(f"筛选{filter_col}：包含{keyword}")

            df = df[mask]
            logger.info(f"{filter_col}筛选后，数据行数：{len(df)}")

    # 必要列检查
    required_columns = CONFIG["required_columns"]["preprocess"]
    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"数据集中缺少必要列: {col}")

    # 数据清洗
    df = df.drop_duplicates(subset=['合同号', '期号'])
    df['起租日'] = pd.to_datetime(df['起租日'])
    df['Vintage'] = df['起租日'].dt.to_period('M').astype(str)

    # 逾期标记
    overdue_threshold = CONFIG["business_params"]["overdue_dpd_threshold"]
    df['is_overdue'] = df['逾期天数（DPD）'].apply(
        lambda x: 1 if (pd.notna(x) and x >= overdue_threshold) else 0
    )
    df_sorted = df.sort_values(['合同号', '期号'])

    # 合同级信息计算
    contract_is_bad = df_sorted.groupby('合同号')['is_overdue'].max().reset_index()
    contract_is_bad.rename(columns={'is_overdue': 'is_bad_asset'}, inplace=True)

    first_overdue = df_sorted[df_sorted['is_overdue'] == 1].groupby('合同号').first().reset_index()
    first_overdue = first_overdue[['合同号', '期号', '未偿还本金']].rename(
        columns={'期号': '首次逾期期号', '未偿还本金': '首次逾期未偿还本金'}
    )

    # 合并合同级信息
    df = pd.merge(df, contract_is_bad, on='合同号', how='left')
    df = pd.merge(df, first_overdue, on='合同号', how='left')
    df['首次逾期未偿还本金'] = df['首次逾期未偿还本金'].fillna(0)
    df['首次逾期期号'] = df['首次逾期期号'].fillna(0).astype(int)

    # 计算当期逾期本金（核心字段）
    df['当期逾期本金'] = df.apply(
        lambda row: row['首次逾期未偿还本金'] if row['期号'] >= row['首次逾期期号'] else 0,
        axis=1
    )

    # 新增：补全提前结清合同的后续期数
    df = fill_missing_periods(df)

    return df


def fill_missing_periods(df, default_total_period=48):
    """
    仅对坏资产补全后续期数：假设坏资产默认最大期数为48期，好资产即使提前结清也不补全
    适用场景：无总期数相关字段，且仅需追踪坏资产的完整逾期表现
    参数：default_total_period - 坏资产默认最大期数，默认48期
    """
    logger.warning(f"未获取到总期数相关数据，采用简便算法：仅对坏资产补全至默认{default_total_period}期")

    # 1. 计算每个合同的“实际结清期”（现有记录中的最大期号）
    contract_last_period = df.groupby('合同号')['期号'].max().reset_index()
    contract_last_period.rename(columns={'期号': '实际结清期'}, inplace=True)

    # 2. 提取合同级关键信息（是否坏资产、首次逾期本金）
    contract_info = df.groupby('合同号').agg(
        是否坏资产=('is_bad_asset', 'max'),  # 1=坏资产，0=好资产
        首次逾期未偿还本金=('首次逾期未偿还本金', 'max')
    ).reset_index()

    # 合并“实际结清期”到合同信息中
    contract_info = pd.merge(contract_info, contract_last_period, on='合同号', how='left')

    # 3. 识别需要补全的合同（核心修改：仅坏资产且实际结清期 < 默认期数）
    need_fill_contracts = contract_info[
        (contract_info['是否坏资产'] == 1) &  # 新增：仅坏资产需要补全
        (contract_info['实际结清期'] < default_total_period) &  # 提前结清
        (contract_info['实际结清期'].notna())  # 排除异常值
    ]
    logger.info(f"需补全后续期数的坏资产合同数：{len(need_fill_contracts)}（实际结清期＜{default_total_period}期）")

    # 新增：统计“好资产且提前结清”的合同（明确说明不补全）
    good_asset_early_settle = contract_info[
        (contract_info['是否坏资产'] == 0) &  # 好资产
        (contract_info['实际结清期'] < default_total_period) &  # 提前结清
        (contract_info['实际结清期'].notna())
    ]
    if len(good_asset_early_settle) > 0:
        logger.info(f"好资产提前结清的合同数：{len(good_asset_early_settle)}（不补全后续期数）")

    # 4. 统计“坏资产但已还满默认期数”的合同（无需补全）
    bad_asset_full_period = contract_info[
        (contract_info['是否坏资产'] == 1) &  # 坏资产
        (contract_info['实际结清期'] >= default_total_period) &  # 已还满默认期数
        (contract_info['实际结清期'].notna())
    ]
    if len(bad_asset_full_period) > 0:
        logger.info(f"坏资产已还满{default_total_period}期的合同数：{len(bad_asset_full_period)}（无需补全）")

    # 5. 生成补全记录（仅针对坏资产）
    filled_records = []
    for _, contract in need_fill_contracts.iterrows():
        contract_id = contract['合同号']
        last_period = int(contract['实际结清期'])
        bad_principal = contract['首次逾期未偿还本金']  # 坏资产的逾期本金基数

        # 获取合同基础信息（复制非期数相关字段）
        base_info = df[df['合同号'] == contract_id].iloc[0].to_dict()

        # 补全从“实际结清期+1”到“默认最大期数”的所有期数
        for period in range(last_period + 1, default_total_period + 1):
            filled_record = base_info.copy()
            filled_record.update({
                '期号': period,
                '结清状态': '坏资产补全（默认48期）',  # 明确标记是坏资产的补全记录
                '租金结算日期': pd.NaT,
                '结清日期': base_info.get('结清日期'),
                '当期逾期本金': bad_principal,  # 坏资产持续计逾期本金
                'is_overdue': 1,  # 坏资产补全期仍视为逾期
                '逾期天数（DPD）': 999,  # 标记持续逾期
                '未偿还本金': bad_principal  # 坏资产补全期未偿还本金不变
            })
            filled_records.append(filled_record)

    # 6. 合并补全记录到原始数据
    if filled_records:
        filled_df = pd.DataFrame(filled_records)[df.columns]
        df = pd.concat([df, filled_df], ignore_index=True)
        logger.info(f"坏资产补全完成：新增{len(filled_records)}条记录（覆盖{len(need_fill_contracts)}个合同）")
    else:
        logger.info("无需要补全的坏资产合同")

    return df



# ============================================
# 4. 构建Vintage-MOB矩阵
# ============================================
def build_vintage_matrix(df):
    """
    构建Vintage-MOB矩阵（基于合同级放款金额，确保每个合同只统计一次）
    """
    # 1. 提取合同级唯一信息：每个合同的放款金额和所属Vintage（去重，避免重复计算）
    # 每个合同只保留一条记录（因放款金额和Vintage对同一合同是固定值）
    contract_level = df.drop_duplicates(subset=['合同号'])[
        ['合同号', 'Vintage', '放款金额']
    ]
    logger.info(f"合同级去重后的数据量：{len(contract_level)}条（每个合同一条）")

    # 2. 按Vintage汇总总放款金额（每个Vintage的总放款=该组下所有合同的放款金额之和）
    vintage_total = contract_level.groupby('Vintage')['放款金额'].sum().reset_index(
        name='总放款金额'
    )
    # 处理放款金额为0的极端情况（避免后续除零）
    vintage_total['总放款金额'] = vintage_total['总放款金额'].replace(0, pd.NA)

    # 3. 按Vintage和期号聚合逾期本金
    vintage_mob = df.groupby(['Vintage', '期号']).agg(
        逾期本金=('当期逾期本金', 'sum')
    ).reset_index()

    # 4. 合并总放款金额并计算逾期率
    vintage_mob = vintage_mob.merge(vintage_total, on='Vintage', how='left')
    # 逾期率=逾期本金/总放款金额（总放款为NA时逾期率设为0）
    vintage_mob['逾期率'] = vintage_mob.apply(
        lambda row: row['逾期本金'] / row['总放款金额'] if pd.notna(row['总放款金额']) else 0,
        axis=1
    )
    # 处理异常值（如总放款为0导致的inf）
    vintage_mob['逾期率'] = vintage_mob['逾期率'].replace([float('inf'), -float('inf')], 0)

    logger.info(
        f"Vintage矩阵构建完成：{vintage_mob['Vintage'].nunique()}个Vintage × {vintage_mob['期号'].nunique()}个MOB"
    )
    return vintage_mob


# ============================================
# 5. 过滤无效MOB（未演化到的期数）
# ============================================
def filter_invalid_mob(vintage_mob_matrix, analysis_date):
    """过滤Vintage矩阵中未演化到的无效MOB期数"""
    # 修正：先处理NaT，再转换日期（避免NaT-01）
    # 1. 标记Vintage列非NaT的有效数据
    valid_mask = vintage_mob_matrix['Vintage'].notna()

    # 2. 仅对有效数据拼接-01并转日期（无效数据设为NaT）
    vintage_mob_matrix.loc[valid_mask, 'Vintage_date'] = pd.to_datetime(
        vintage_mob_matrix.loc[valid_mask, 'Vintage'].astype(str) + '-01',
        format='%Y-%m-%d',  # 明确日期格式，提高转换效率
        errors='coerce'  # 极端异常值转为NaT，不中断程序
    )
    vintage_mob_matrix.loc[~valid_mask, 'Vintage_date'] = pd.NaT  # 无效数据设为NaT

    # 3. 删除Vintage_date为空的无效行（避免后续计算出错）
    vintage_mob_matrix = vintage_mob_matrix.dropna(subset=['Vintage_date'])

    # 新增：过滤未来Vintage（Vintage日期晚于分析基准日）
    vintage_mob_matrix = vintage_mob_matrix[vintage_mob_matrix['Vintage_date'] <= analysis_date]

    # 计算每个Vintage的最大有效MOB
    vintage_mob_matrix['month_diff'] = (
            (analysis_date.year - vintage_mob_matrix['Vintage_date'].dt.year) * 12
            + (analysis_date.month - vintage_mob_matrix['Vintage_date'].dt.month)
    )

    # 修正：month_diff<0时设为0（无有效MOB），再clip（避免0）
    vintage_mob_matrix['max_valid_mob'] = vintage_mob_matrix['month_diff'].clip(lower=0).replace(0, 1)

    # 筛选有效数据
    valid_mob_matrix = vintage_mob_matrix[
        vintage_mob_matrix['期号'] <= vintage_mob_matrix['max_valid_mob']
        ].copy()

    # 清理临时列
    valid_mob_matrix = valid_mob_matrix.drop(columns=['Vintage_date', 'month_diff', 'max_valid_mob'])
    logger.info(f"有效MOB筛选完成：原始记录数{len(vintage_mob_matrix)} → 有效记录数{len(valid_mob_matrix)}")

    # 打印各Vintage的有效MOB范围
    valid_mob_summary = valid_mob_matrix.groupby('Vintage').agg(
        最小有效MOB=('期号', 'min'),
        最大有效MOB=('期号', 'max')
    ).reset_index()
    logger.info("各Vintage有效MOB范围：")
    for _, row in valid_mob_summary.iterrows():
        logger.info(f"  - {row['Vintage']}: MOB{row['最小有效MOB']}~MOB{row['最大有效MOB']}")

    return valid_mob_matrix


# ============================================
# 主程序
# ============================================
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl.styles.stylesheet')
if __name__ == "__main__":
    try:
        # 1. 数据加载与合并
        df_merged = merge_excel_files()
        if df_merged is None:
            raise ValueError("未成功加载数据")

        # 2. 数据预处理
        filter_params = {
            "经销商名称": CONFIG["dealer_filters"],
            "业务类别": ["商用车"],
            "业务模式": ["损失不担模式"]
        }
        df_clean = preprocess_data(df_merged, filter_params)

        # 3. 构建Vintage矩阵
        vintage_mob_matrix = build_vintage_matrix(df_clean)

        # 4. 确定分析基准日（上月最后一天）
        # 修正：保留pd.Timestamp类型，不转为字符串
        today = pd.Timestamp.now()
        # 计算上月最后一天（日期类型）
        last_day_of_last_month = today.replace(day=1) - pd.Timedelta(days=1)
        ANALYSIS_DATE = last_day_of_last_month  # 此时ANALYSIS_DATE是pd.Timestamp类型
        # 日志打印时再转为字符串（不影响计算）
        logger.info(f"分析基准日自动设置为：{ANALYSIS_DATE.strftime('%Y-%m-%d')}")

        # 5. 过滤无效MOB期数
        vintage_mob_valid = filter_invalid_mob(vintage_mob_matrix, ANALYSIS_DATE)

        # 6. 生成透视表
        pivot_table = vintage_mob_valid.pivot_table(
            index='Vintage',
            columns='期号',
            values='逾期率',
            aggfunc='mean'
        ).fillna('-')

        # 7. 保存结果到Excel
        output_path = CONFIG["output_file"]
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 保存Vintage矩阵透视表
            pivot_table.to_excel(writer, sheet_name='Vintage矩阵', index_label='Vintage')

            # 保存有效明细数据
            vintage_mob_valid.to_excel(writer, sheet_name='有效明细数据', index=False)

            # 保存首次逾期明细（仅保留每个合同第一次逾期的记录）
            # 1. 先筛选出所有逾期记录（is_overdue=1）
            overdue_records = df_clean[
                df_clean['is_overdue'] == 1  # 只保留逾期记录
                ][
                [
                    'Vintage', '合同号', '经销商名称', '客户名称', '大区', '业务来源', '起租日', '期号',
                    '放款金额', '未偿还本金', '逾期天数（DPD）',  '首次逾期期号', '首次逾期未偿还本金'
                ]
            ]

            # 2. 按合同号分组，每组只保留期号最小的那条记录（即首次逾期）
            first_overdue_only = overdue_records.sort_values(by=['合同号', '期号']) \
                .groupby('合同号').first().reset_index()

            # 3. 按Vintage和期号排序，确保展示有序
            first_overdue_only = first_overdue_only.sort_values(by=['Vintage', '期号', '合同号'])

            # 写入Excel（工作表名称更改为“首次逾期明细”）
            first_overdue_only.to_excel(writer, sheet_name='首次逾期明细', index=False)

            # 美化Vintage矩阵格式
            worksheet = writer.sheets['Vintage矩阵']
            valid_fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
            for row in worksheet.iter_rows(min_row=2, min_col=2):
                for cell in row:
                    if cell.value != '-':
                        cell.fill = valid_fill

        logger.info(f"分析完成，结果已保存至: {output_path}")
        print(f"分析完成，结果已保存至: {output_path}")

    except Exception as e:
        logger.error(f"主程序执行错误：{e}", exc_info=True)
