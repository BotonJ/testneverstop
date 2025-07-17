# /src/data_processor.py

import pandas as pd
import logging

logger = logging.getLogger(__name__)

def calculate_summary_values(raw_df: pd.DataFrame):
    """
    【V4.0】从raw_df高效计算所有核心汇总值，为报告生成器提供原始数据。
    """
    logger.info("  -> 从 raw_df 计算核心 summary_values...")
    summary = {}
    
    df = raw_df.copy()
    df['年份'] = pd.to_numeric(df['年份'], errors='coerce')
    df.dropna(subset=['年份'], inplace=True)
    df['年份'] = df['年份'].astype(int)

    years = sorted(df['年份'].unique())
    if not years: return {}
        
    start_year, end_year = years[0], years[-1]
    
    def _get_val(item, year, col):
        val_series = df[(df['项目'] == item) & (df['年份'] == year)][col]
        return val_series.iloc[0] if not val_series.empty else 0

    # 计算每个年度的净资产变动额
    for year in years:
        start_net_asset = _get_val('净资产合计', year, '期初金额')
        end_net_asset = _get_val('净资产合计', year, '期末金额')
        summary[f"{year}_净资产变动额"] = end_net_asset - start_net_asset
            
    # ... 其他汇总计算 ...
    return summary