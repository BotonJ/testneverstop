# /src/data_processor.py
import pandas as pd
from typing import Tuple # <--- 导入Tuple类型
from src.utils.logger_config import logger

def pivot_and_clean_data(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]: # <--- 使用规范的类型提示
    """
    【新版】
    将数据进行透视和清理。
    返回两个DataFrame: 一个包含普通科目，一个包含合计科目。
    """
    logger.info("开始进行数据透视和清理...")
    
    # 确保'科目类型'列存在
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

# calculate_summary_values 函数保持不变，无需修改
def calculate_summary_values(pivoted_normal_df: pd.DataFrame, pivoted_total_df: pd.DataFrame) -> dict:
    # ... 此函数内容与上一版完全相同，此处省略 ...
    logger.info("开始计算最终汇总指标...")
    summary = {}
    
    df_to_use = pivoted_total_df
    
    if df_to_use.empty:
        logger.error("传入的合计科目DataFrame为空，无法计算汇总指标。")
        return summary

    years = sorted([col for col in df_to_use.columns if str(col).isdigit()])
    start_year = years[0]
    end_year = years[-1]
    
    summary['起始年份'] = start_year
    summary['终止年份'] = end_year
    logger.info(f"数据期间为: {start_year} 年至 {end_year} 年。")

    def _get_value(item_name, year_or_years):
        try:
            # 判断是单个年份还是年份列表
            if isinstance(year_or_years, list):
                # 如果是列表，执行求和
                return df_to_use.loc[item_name, year_or_years].sum()
            else:
                # 否则，取单个值
                return df_to_use.loc[item_name, year_or_years]
        except KeyError:
            logger.warning(f"在合计表中未能找到项目'{item_name}'的数据，将使用0代替。")
            return 0

    summary['期初资产总额'] = _get_value('资产总计', start_year)
    summary['期末资产总额'] = _get_value('资产总计', end_year)
    summary['期初负债总额'] = _get_value('负债合计', start_year)
    summary['期末负债总额'] = _get_value('负债合计', end_year)
    summary['期初净资产总额'] = _get_value('净资产合计', start_year)
    summary['期末净资产总额'] = _get_value('净资产合计', end_year)

    summary['资产总额增减'] = summary['期末资产总额'] - summary['期初资产总额']
    summary['负债总额增减'] = summary['期末负债总额'] - summary['期初负债总额']
    summary['净资产总额增减'] = summary['期末净资产总额'] - summary['期初净资产总额']
    
    logger.info("资产、负债、净资产指标计算完成。")

    summary['审计期间收入总额'] = _get_value('收入合计', years)
    summary['审计期间费用总额'] = _get_value('费用合计', years)
    summary['审计期间净结余'] = summary['审计期间收入总额'] - summary['审计期间费用总额']
    logger.info("收入、费用、结余指标计算完成。")
        
    logger.info("所有汇总指标计算完成。")
    return summary