# /src/data_processor.py

import pandas as pd
from src.utils.logger_config import logger



def pivot_and_clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    【核心函数1】
    将从legacy_runner提取出来的长格式DataFrame进行透视和清理。
    转换成以“项目”为索引，以“年份”和“金额类型”为列的宽格式表。

    Args:
        df (pd.DataFrame): 从legacy_runner.py的run_legacy_extraction函数获取的原始DataFrame。

    Returns:
        pd.DataFrame: 一个经过透视和初步处理的、更适合分析的DataFrame。
    """
    logger.info("开始进行数据透视和清理...")
    
    # 筛选出资产负债表和业务活动表的数据
    balance_sheet_df = df[df['报表类型'] == '资产负债表'].copy()
    income_statement_df = df[df['报表类型'] == '业务活动表'].copy()

    # --- 处理资产负债表 ---
    # 我们只需要'期末金额'，因为下一年的期初就是上一年的期末
    balance_sheet_df = balance_sheet_df[['年份', '项目', '期末金额']]
    # 使用pivot_table进行透视，将年份作为列
    bs_pivot = balance_sheet_df.pivot_table(
        index='项目', 
        columns='年份', 
        values='期末金额'
    ).sort_index()
    logger.info("资产负债表数据透视完成。")

    # --- 处理业务活动表 ---
    # 业务活动表记录的是当期发生额，所以我们使用'本期金额'
    income_statement_df = income_statement_df[['年份', '项目', '本期金额']]
    is_pivot = income_statement_df.pivot_table(
        index='项目', 
        columns='年份', 
        values='本期金额'
    ).sort_index()
    logger.info("业务活动表数据透视完成。")

    # --- 合并两张透视表 ---
    # 使用pd.concat进行合并，相同的项目会自动对齐
    final_pivot_df = pd.concat([bs_pivot, is_pivot], axis=0)
    
    # 清理工作
    final_pivot_df = final_pivot_df.fillna(0) # 将所有NaN值填充为0
    # 将列名（年份）按数字顺序排序
    final_pivot_df = final_pivot_df.reindex(sorted(final_pivot_df.columns), axis=1)

    logger.info("数据透视和清理完成。")
    return final_pivot_df


def calculate_summary_values(pivoted_df: pd.DataFrame) -> dict:
    """
    【核心函数2】
    从已经透视好的DataFrame中，计算最终报告所需的各项核心指标。
    例如：期初总资产、期末总资产、净资产变动额等。

    Args:
        pivoted_df (pd.DataFrame): 经过pivot_and_clean_data函数处理后的宽格式DataFrame。

    Returns:
        dict: 一个包含所有最终计算指标的字典，可用于注入报告模板。
    """
    logger.info("开始计算最终汇总指标...")
    
    summary = {}
    
    if pivoted_df.empty:
        logger.error("传入的DataFrame为空，无法计算汇总指标。")
        return summary

    # 获取时间范围
    years = sorted([col for col in pivoted_df.columns if str(col).isdigit()])
    start_year = years[0]
    end_year = years[-1]
    
    summary['起始年份'] = start_year
    summary['终止年份'] = end_year
    logger.info(f"数据期间为: {start_year} 年至 {end_year} 年。")

    # --- 计算资产、负债、净资产相关指标 ---
    # 注意：这里的'资产总计'等字符串需要和mapping_file中的'标准科目名'完全一致
    try:
        # 获取期初、期末的各项总额
        # 期初 = 最早一年的期初，但因为资产负债表我们只用了期末值，所以需要找到 start_year-1 的期末值
        # 为了简化，我们暂时将最早一年的值作为期初（可以在后续流程中传入更早一年的数据来修复）
        summary['期初资产总额'] = pivoted_df.loc['资产总计', start_year]
        summary['期末资产总额'] = pivoted_df.loc['资产总计', end_year]
        
        summary['期初负债总额'] = pivoted_df.loc['负债合计', start_year]
        summary['期末负债总额'] = pivoted_df.loc['负债合计', end_year]

        summary['期初净资产总额'] = pivoted_df.loc['净资产合计', start_year]
        summary['期末净资产总额'] = pivoted_df.loc['净资产合计', end_year]

        # 计算增减
        summary['资产总额增减'] = summary['期末资产总额'] - summary['期初资产总额']
        summary['负债总额增减'] = summary['期末负债总额'] - summary['期初负债总额']
        summary['净资产总额增减'] = summary['期末净资产总额'] - summary['期初净资产总额']
        
        logger.info("资产、负债、净资产指标计算完成。")

    except KeyError as e:
        logger.error(f"计算汇总指标时出错：找不到关键项目 '{e}'。请检查mapping_file中的标准科目名是否正确。")

    # --- 计算总收入、总支出、总结余 ---
    try:
        # 对所有年份的收入和支出进行求和
        total_income = pivoted_df.loc['收入合计', years].sum()
        total_expense = pivoted_df.loc['费用合计', years].sum()
        
        summary['审计期间收入总额'] = total_income
        summary['审计期间费用总额'] = total_expense
        summary['审计期间净结余'] = total_income - total_expense
        logger.info("收入、费用、结余指标计算完成。")
        
    except KeyError as e:
        logger.error(f"计算收支指标时出错：找不到关键项目 '{e}'。")

    logger.info("所有汇总指标计算完成。")
    return summary


# --- 测试入口 ---
if __name__ == '__main__':
    # 模拟一个从legacy_runner获取的DataFrame
    mock_data = {
        '来源Sheet': ['2021资产负债表', '2021资产负债表', '2022资产负债表', '2022资产负债表', '2021业务活动表', '2022业务活动表'],
        '报表类型': ['资产负债表', '资产负债表', '资产负债表', '资产负债表', '业务活动表', '业务活动表'],
        '年份': [2021, 2021, 2022, 2022, 2021, 2022],
        '项目': ['资产总计', '负债合计', '资产总计', '负债合计', '收入合计', '收入合计'],
        '期初金额': [0,0,0,0,0,0],
        '期末金额': [1000, 800, 1200, 900, 0, 0],
        '本期金额': [0, 0, 0, 0, 500, 600],
        '上期金额': [0,0,0,0,0,0]
    }
    mock_df = pd.DataFrame(mock_data)
    # 手动计算净资产
    mock_df['净资产合计'] = mock_df['期末金额'] - mock_df.get('负债合计', 0) # 简化的计算
    
    print("--- 测试 data_processor.py ---")
    print("\n【输入】模拟的原始DataFrame:")
    print(mock_df)
    
    # 1. 测试透视功能
    pivoted = pivot_and_clean_data(mock_df)
    print("\n【步骤1输出】透视后的DataFrame:")
    print(pivoted)
    
    # 2. 测试计算功能
    # 为了测试，需要手动添加'费用合计'和'净资产合计'到透视表
    pivoted.loc['费用合计'] = [400, 550]
    pivoted.loc['净资产合计'] = [200, 300]
    
    final_summary = calculate_summary_values(pivoted)
    print("\n【步骤2输出】最终计算的汇总指标字典:")
    import json
    print(json.dumps(final_summary, indent=4, ensure_ascii=False))