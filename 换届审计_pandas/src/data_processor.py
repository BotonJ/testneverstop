# /src/data_processor.py
import pandas as pd
from typing import Tuple
from src.utils.logger_config import logger
from modules.utils import normalize_name

def pivot_and_clean_data(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    【最终版】
    1. 将数据进行透视和清理。
    2. 【新增】实现“数据自动衔接”逻辑。
    """
    logger.info("开始进行数据透视和清理...")
    if '科目类型' not in df.columns:
        return pd.DataFrame(), pd.DataFrame()
    
    # --- 1. 数据透视 (与之前版本相同) ---
    normal_subjects_df = df[df['科目类型'] == '普通'].copy()
    total_subjects_df = df[df['科目类型'] == '合计'].copy()
    def _pivot(input_df, name):
        if input_df.empty:
            return pd.DataFrame()
        bs_df = input_df[input_df['报表类型'] == '资产负债表'][['年份', '项目', '期末金额']]
        bs_pivot = bs_df.pivot_table(index='项目', columns='年份', values='期末金额') if not bs_df.empty else pd.DataFrame()
        is_df = input_df[input_df['报表类型'] == '业务活动表'][['年份', '项目', '本期金额']]
        is_pivot = is_df.pivot_table(index='项目', columns='年份', values='本期金额') if not is_df.empty else pd.DataFrame()
        final_pivot = pd.concat([bs_pivot, is_pivot], axis=0).fillna(0)
        if not final_pivot.empty:
            final_pivot = final_pivot.reindex(sorted(final_pivot.columns), axis=1)
        logger.info(f"{name}数据透视完成。")
        return final_pivot
    pivoted_normal = _pivot(normal_subjects_df, "普通科目")
    pivoted_total = _pivot(total_subjects_df, "合计科目")

    # --- 2. 【新增】数据自动衔接逻辑 (在业务活动表上) ---
    logger.info("开始执行业务活动表数据自动衔接...")
    # 我们需要在原始的、未透视的数据上操作，以获取期初和期末
    biz_df = df[df['报表类型'] == '业务活动表'].copy()
    years = sorted(biz_df['年份'].unique())
    
    for i, year in enumerate(years):
        if i == 0: continue # 第一年没有前一年，跳过
        prev_year = years[i-1]
        
        # 找到当前年份和前一年的数据
        current_year_data = biz_df[biz_df['年份'] == year].set_index('项目')
        prev_year_data = biz_df[biz_df['年份'] == prev_year].set_index('项目')
        
        # 遍历当前年份的数据，检查上期金额是否为空
        for project, row in current_year_data.iterrows():
            if pd.isna(row['上期金额']) or row['上期金额'] == 0:
                # 如果为空，则查找前一年的本期金额来填充
                if project in prev_year_data.index:
                    bridged_value = prev_year_data.loc[project, '本期金额']
                    # 更新我们主数据源 raw_df 中的值
                    df.loc[(df['项目'] == project) & (df['年份'] == year), '上期金额'] = bridged_value
                    logger.debug(f"数据衔接: 已将'{project}'项目{year}年的上期金额更新为{prev_year}年的本期金额({bridged_value})。")

    logger.info("数据透视和清理完成 (已包含数据衔接)。")
    return pivoted_normal, pivoted_total

# calculate_summary_values 函数保持不变
def calculate_summary_values(pivoted_total_df: pd.DataFrame, raw_df: pd.DataFrame) -> dict:
    # ... 此函数内容与上一版完全相同 ...
    logger.info("开始计算最终汇总指标...")
    summary, years = {}, sorted([col for col in pivoted_total_df.columns if str(col).isdigit()])
    if not years: return {}
    start_year, end_year = years[0], years[-1]
    summary.update({'起始年份': start_year, '终止年份': end_year})
    def _get_raw(item, yr, col):
        try:
            return raw_df[(raw_df['项目'] == item) & (raw_df['年份'] == yr)][col].iloc[0]
        except (KeyError, IndexError): return 0
    def _get_pivot(item, yr_or_yrs):
        try:
            return pivoted_total_df.loc[item, yr_or_yrs].sum() if isinstance(yr_or_yrs, list) else pivoted_total_df.loc[item, yr_or_yrs]
        except KeyError: return 0
    summary['期初资产总额'] = _get_raw(normalize_name('资产总计'), start_year, '期初金额')
    summary['期末资产总额'] = _get_raw(normalize_name('资产总计'), end_year, '期末金额')
    summary['期初负债总额'] = _get_raw(normalize_name('负债合计'), start_year, '期初金额')
    summary['期末负债总额'] = _get_raw(normalize_name('负债合计'), end_year, '期末金额')
    summary['期初净资产总额'] = _get_raw(normalize_name('净资产合计'), start_year, '期初金额')
    summary['期末净资产总额'] = _get_raw(normalize_name('净资产合计'), end_year, '期末金额')
    summary['资产总额增减'] = summary['期末资产总额'] - summary['期初资产总额']
    summary['负债总额增减'] = summary['期末负债总额'] - summary['期初负债总额']
    summary['净资产总额增减'] = summary['期末净资产总额'] - summary['期初净资产总额']
    summary['审计期间收入总额'] = _get_pivot(normalize_name('收入合计'), years)
    summary['审计期间费用总额'] = _get_pivot(normalize_name('费用合计'), years)
    summary['审计期间净结余'] = summary['审计期间收入总额'] - summary['审计期间费用总额']
    logger.info("所有汇总指标计算完成。")
    return summary