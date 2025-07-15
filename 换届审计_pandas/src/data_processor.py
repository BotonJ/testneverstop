# /src/data_processor.py

import pandas as pd
import logging
from typing import Dict, Tuple

logger = logging.getLogger(__name__)

def pivot_and_clean_data(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    logger.info("开始进行数据透视和清理...")
    
    # 确保年份列是数值类型，便于排序
    df['年份'] = pd.to_numeric(df['年份'], errors='coerce')
    df.dropna(subset=['年份'], inplace=True)
    df['年份'] = df['年份'].astype(int)

    # 分离普通科目和合计科目
    normal_df = df[df['科目类型'] == '普通'].copy()
    total_df = df[df['科目类型'] == '合计'].copy()

    # --- 透视普通科目 ---
    pivoted_normal = normal_df.pivot_table(
        index='项目',
        columns='年份',
        values=['期初金额', '期末金额', '本期金额', '上期金额'],
        aggfunc='sum'
    ).fillna(0)
    logger.info("普通科目数据透视完成。")

    # --- 透视合计科目 ---
    pivoted_total = total_df.pivot_table(
        index='项目',
        columns='年份',
        values=['期初金额', '期末金额', '本期金额', '上期金额'],
        aggfunc='sum'
    ).fillna(0)
    logger.info("合计科目数据透视完成。")
    
    # 将多层列名简化
    if not pivoted_total.empty:
        pivoted_total.columns = [col[1] for col in pivoted_total.columns]
        pivoted_total = pivoted_total.reindex(sorted(pivoted_total.columns), axis=1)

    logger.info("数据透视和清理完成。")
    return pivoted_normal, pivoted_total


def calculate_summary_values(pivoted_total_df: pd.DataFrame, raw_df: pd.DataFrame) -> Dict:
    logger.info("开始计算最终汇总指标...")
    summary = {}
    
    if pivoted_total_df.empty:
        logger.error("传入的合计科目透视表为空，无法计算汇总指标。")
        return summary

    # --- [修复] 改进年份和期初/期末值的提取逻辑 ---
    years = sorted(pivoted_total_df.columns)
    start_year = years[0]
    end_year = years[-1]
    
    summary['起始年份'] = str(start_year)
    summary['终止年份'] = str(end_year)

    def _get_value(item, year, df):
        try:
            return df.loc[item, year]
        except KeyError:
            logger.debug(f"在透视表中查找 '{item}' ({year}年) 失败，返回 0。")
            return 0
    
    def _get_raw_value(item, year, col, source_df):
        try:
            val = source_df[(source_df['项目'] == item) & (source_df['年份'] == year)][col].iloc[0]
            return val if pd.notna(val) else 0
        except (KeyError, IndexError):
            logger.debug(f"在原始数据中查找 '{item}' ({year}年, {col}列) 失败，返回 0。")
            return 0

    # --- 计算资产、负债、净资产相关指标 ---
    try:
        # 期初值 = 最早一年的期初金额
        summary['期初资产总额'] = _get_raw_value('资产总计', start_year, '期初金额', raw_df)
        summary['期末资产总额'] = _get_value('资产总计', end_year, pivoted_total_df)
        
        summary['期初负债总额'] = _get_raw_value('负债合计', start_year, '期初金额', raw_df)
        summary['期末负债总额'] = _get_value('负债合计', end_year, pivoted_total_df)

        summary['期初净资产总额'] = _get_raw_value('净资产合计', start_year, '期初金额', raw_df)
        summary['期末净资产总额'] = _get_value('净资产合计', end_year, pivoted_total_df)

        summary['资产总额增减'] = summary['期末资产总额'] - summary['期初资产总额']
        summary['负债总额增减'] = summary['期末负债总额'] - summary['期初负债总额']
        summary['净资产总额增减'] = summary['期末净资产总额'] - summary['期初净资产总额']
        
    except KeyError as e:
        logger.error(f"计算资产负债指标时出错：找不到关键项目 {e}。")

    # --- 计算总收入、总支出、总结余 ---
    try:
        income_items = raw_df[raw_df['类型'] == '收入']['项目'].unique()
        expense_items = raw_df[raw_df['类型'] == '费用']['项目'].unique()

        total_income = pivoted_total_df.loc[pivoted_total_df.index.isin(income_items), years].sum().sum()
        total_expense = pivoted_total_df.loc[pivoted_total_df.index.isin(expense_items), years].sum().sum()
        
        summary['审计期间收入总额'] = total_income
        summary['审计期间费用总额'] = total_expense
        summary['审计期间净结余'] = total_income - total_expense
        
    except KeyError as e:
        logger.error(f"计算收支指标时出错：找不到关键项目 {e}。")

    logger.info("所有汇总指标计算完成。")
    return summary
