# /src/data_validator.py
import pandas as pd
from src.utils.logger_config import logger

def run_all_checks(pivoted_normal_df, pivoted_total_df, raw_df, mapping):
    logger.info("--- [复核机制] 开始执行所有数据检查... ---")
    results = []
    
    if pivoted_total_df.empty:
        results.append("❌ 错误: 合计项数据表为空，无法执行复核。")
        return results

    years = sorted([col for col in pivoted_total_df.columns if str(col).isdigit()])
    if not years:
        results.append("❌ 错误: 无法确定复核年份。")
        return results

    # --- 检查 1: 业务活动表内部平衡 ---
    logger.info("  -> 正在执行: 业务活动表内部分项核对...")
    type_to_total_map = {'收入': '收入合计', '费用': '费用合计'}
    yewu_subtotal_config = mapping.get("yewu_subtotal_config", {})
    if yewu_subtotal_config:
        for config_type, sub_items_list in yewu_subtotal_config.items():
            standard_total_name = type_to_total_map.get(config_type)
            if standard_total_name:
                results.extend(
                    _check_subtotal(pivoted_normal_df, pivoted_total_df, sub_items_list, standard_total_name, years)
                )

    # --- 检查 2: 核心勾稽关系 ---
    logger.info("  -> 正在执行: 核心勾稽关系检查...")
    results.extend(_check_core_equalities(pivoted_total_df, years))
    
    logger.info("--- [复核机制] 所有数据检查执行完毕。 ---")
    return results

def _check_subtotal(normal_df, total_df, sub_items_list, total_item_name, years):
    check_results = []
    if total_item_name not in total_df.index:
        check_results.append(f"❌ 复核失败: 关键合计项 '{total_item_name}' 未能成功提取。")
        return check_results

    calculated_totals = normal_df[normal_df.index.isin(sub_items_list)].sum()

    for year in years:
        report_total = total_df.loc[total_item_name, year]
        calculated_total = calculated_totals.get(year, 0)
        diff = calculated_total - report_total
        if abs(diff) < 0.01:
            msg = f"✅ {year}年'{total_item_name}'内部分项核对平衡 (计算值 {calculated_total:,.2f})"
            check_results.append(msg)
        else:
            msg = f"❌ {year}年'{total_item_name}'内部分项核对**不平**: 计算值 {calculated_total:,.2f} vs 报表值 {report_total:,.2f} (差异: {diff:,.2f})"
            check_results.append(msg)
    return check_results

def _check_core_equalities(total_df, years):
    results = []
    required_totals = ['资产总计', '负债合计', '净资产合计', '收入合计', '费用合计']
    missing_totals = [t for t in required_totals if t not in total_df.index]
    if missing_totals:
        results.append(f"❌ 核心勾稽关系检查失败: 缺少关键合计项 {missing_totals}")
        return results

    # ... 此函数其余部分保持不变 ...
    start_year, end_year = years[0], years[-1]
    for year in years:
        asset, lia, equity = total_df.loc['资产总计', year], total_df.loc['负债合计', year], total_df.loc['净资产合计', year]
        diff = asset - (lia + equity)
        if abs(diff) < 0.01:
            results.append(f"✅ {year}年资产负债表内部平衡")
        else:
            results.append(f"❌ {year}年资产负债表内部**不平** (差异: {diff:,.2f})")
    net_asset_change = total_df.loc['净资产合计', end_year] - total_df.loc['净资产合计', start_year]
    income = total_df.loc['收入合计', years].sum()
    expense = total_df.loc['费用合计', years].sum()
    net_profit = income - expense
    diff = net_asset_change - net_profit
    if abs(diff) < 0.01:
        results.append(f"✅ 跨期核心勾稽关系平衡")
    else:
        results.append(f"❌ 跨期核心勾稽关系**不平** (差异: {diff:,.2f})")
    return results