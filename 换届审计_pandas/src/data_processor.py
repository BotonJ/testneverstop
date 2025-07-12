# /src/data_processor.py
import pandas as pd
from typing import Tuple
from src.utils.logger_config import logger

def pivot_and_clean_data(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # ... 此函数保持不变，此处省略 ...
    logger.info("开始进行数据透视和清理...")
    if '科目类型' not in df.columns:
        logger.error("输入的DataFrame缺少'科目类型'列，无法进行分类处理。")
        return pd.DataFrame(), pd.DataFrame()
    normal_subjects_df = df[df['科目类型'] == '普通'].copy()
    total_subjects_df = df[df['科目类型'] == '合计'].copy()
    def _pivot(input_df, name):
        if input_df.empty:
            logger.info(f"{name}数据为空，跳过透视。")
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
    logger.info("数据透视和清理完成。")
    return pivoted_normal, pivoted_total


def calculate_summary_values(pivoted_total_df: pd.DataFrame, raw_df: pd.DataFrame) -> dict:
    """
    【最终修复版】
    整合了“期初的期初，期末的期末”精确取值逻辑。
    """
    logger.info("开始计算最终汇总指标...")
    summary = {}
    
    # 优先检查透视合计表
    if pivoted_total_df.empty:
        logger.error("传入的合计科目透视表(pivoted_total_df)为空，无法计算汇总指标。")
        return summary

    # 从透视表中获取年份范围
    years = sorted([col for col in pivoted_total_df.columns if str(col).isdigit()])
    if not years:
        logger.error("无法从透视表中确定年份范围。")
        return summary
    start_year = years[0]
    end_year = years[-1]
    
    summary['起始年份'] = start_year
    summary['终止年份'] = end_year
    logger.info(f"数据期间为: {start_year} 年至 {end_year} 年。")

    # --- vvvvvvvv 这是您新增的、现在被正确缩进到函数内部的代码块 vvvvvvvv ---

    # 1. 定义一个嵌套的辅助函数，它只能在当前函数内部被调用
    def _get_value_from_raw(item_name, year, col_name):
        """一个专门从raw_df中精确查找单一值的辅助函数"""
        try:
            # 筛选出特定年份和特定项目的那一行，然后取指定列的值
            value = raw_df[(raw_df['项目'] == item_name) & (raw_df['年份'] == year)][col_name].iloc[0]
            return value
        except (KeyError, IndexError):
            logger.warning(f"在原始数据中未能找到项目'{item_name}'的{year}年'{col_name}'，将使用0代替。")
            return 0

    # 2. 使用这个辅助函数来精确获取期初和期末的值
    summary['期初资产总额'] = _get_value_from_raw('资产总计', start_year, '期初金额')
    summary['期末资产总额'] = _get_value_from_raw('资产总计', end_year, '期末金额')

    summary['期初负债总额'] = _get_value_from_raw('负债合计', start_year, '期初金额')
    summary['期末负债总额'] = _get_value_from_raw('负债合计', end_year, '期末金额')

    summary['期初净资产总额'] = _get_value_from_raw('净资产合计', start_year, '期初金额')
    summary['期末净资产总额'] = _get_value_from_raw('净资产合计', end_year, '期末金额')

    # --- ^^^^^^^^ 新代码块结束 ^^^^^^^^ ---
    
    # --- 收入和费用的计算逻辑保持不变，但使用一个新的辅助函数 ---
    def _get_value_from_pivoted(item_name, year_or_years):
        """专门从合计透视表中取值的辅助函数"""
        try:
            if isinstance(year_or_years, list):
                return pivoted_total_df.loc[item_name, year_or_years].sum()
            else:
                return pivoted_total_df.loc[item_name, year_or_years]
        except KeyError:
            logger.warning(f"在合计透视表中未能找到项目'{item_name}'的数据，将使用0代替。")
            return 0

    summary['资产总额增减'] = summary['期末资产总额'] - summary['期初资产总额']
    summary['负债总额增减'] = summary['期末负债总额'] - summary['期初负债总额']
    summary['净资产总额增减'] = summary['期末净资产总额'] - summary['期初净资产总额']
    logger.info("资产、负债、净资产指标计算完成。")

    summary['审计期间收入总额'] = _get_value_from_pivoted('收入合计', years)
    summary['审计期间费用总额'] = _get_value_from_pivoted('费用合计', years)
    summary['审计期间净结余'] = summary['审计期间收入总额'] - summary['审计期间费用总额']
    logger.info("收入、费用、结余指标计算完成。")
    
    logger.info("所有汇总指标计算完成。")
    return summary