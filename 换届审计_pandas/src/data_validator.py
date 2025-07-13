# /src/data_validator.py
import pandas as pd
from src.utils.logger_config import logger
from modules.utils import normalize_name

def run_all_checks(pivoted_normal_df, pivoted_total_df, raw_df, mapping):
    """【复核机制主函数】"""
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
    yewu_subtotal_config = mapping.get("yewu_subtotal_config", {})
    if yewu_subtotal_config:
        type_to_total_map = {
            '收入': normalize_name('收入合计'), 
            '费用': normalize_name('费用合计')
        }
        for config_type, sub_items_list in yewu_subtotal_config.items():
            standard_total_name = type_to_total_map.get(config_type)
            if standard_total_name:
                results.extend(
                    _check_subtotal_biz(pivoted_normal_df, pivoted_total_df, sub_items_list, standard_total_name, years)
                )

    # --- 检查 2: 资产负债表内部平衡 (分项 vs 合计) ---
    logger.info("  -> 正在执行: 资产负债表内部分项核对...")
    blocks_df = mapping.get("blocks_df")
    normal_bs_raw_df = raw_df[(raw_df['报表类型'] == '资产负债表') & (raw_df['科目类型'] == '普通')].copy()
    if blocks_df is not None and not blocks_df.empty and not normal_bs_raw_df.empty:
        results.extend(
            _check_balance_sheet_subtotals(normal_bs_raw_df, pivoted_total_df, blocks_df, years)
        )
    
    # --- 检查 3: 核心勾稽关系 ---
    logger.info("  -> 正在执行: 核心勾稽关系检查...")
    results.extend(_check_core_equalities(pivoted_total_df, raw_df, years))
    
    logger.info("--- [复核机制] 所有数据检查执行完毕。 ---")
    return results

def _check_subtotal_biz(normal_df, total_df, sub_items_list, total_item_name, years):
    """业务活动表分项与合计交叉验证"""
    check_results = []
    if total_item_name not in total_df.index.map(normalize_name):
        check_results.append(f"❌ 复核失败: 关键合计项 '{total_item_name}' 未能成功提取。")
        return check_results

    normal_df.index = normal_df.index.map(normalize_name)
    calculated_totals = normal_df[normal_df.index.isin(sub_items_list)].sum()

    for year in years:
        report_total = total_df.loc[total_item_name, year]
        calculated_total = calculated_totals.get(year, 0)
        diff = calculated_total - report_total
        if abs(diff) < 0.01:
            msg = f"✅ {year}年'{total_item_name}'内部分项核对平衡 (计算值 {calculated_total:,.2f})"
            check_results.append(msg)
        else:
            msg = f"❌ {year}年'{total_item_name}'内部分项核对**不平** (差异: {diff:,.2f})"
            check_results.append(msg)
    return check_results

def _check_balance_sheet_subtotals(normal_raw_df, total_df, blocks_df, years):
    """【最终修复版】根据“资产负债表区块”配置，进行交叉验证。"""
    check_results = []
    if '所属区块' not in normal_raw_df.columns:
        check_results.append("❌ 资产负债表复核失败: 缺少'所属区块'信息。")
        return check_results

    clean_total_df_index = total_df.index.map(normalize_name)

    for _, block_row in blocks_df.iterrows():
        total_item_name = block_row.get('区块名称')
        if not total_item_name or '合计' not in total_item_name:
            continue
        
        total_item_name_clean = normalize_name(total_item_name)
        if total_item_name_clean not in clean_total_df_index:
            check_results.append(f"❌ 复核失败: 关键合计项 '{total_item_name}' 未能成功提取。")
            continue

        for year in years:
            sub_items_df = normal_raw_df[
                (normal_raw_df['所属区块'] == total_item_name) &
                (normal_raw_df['年份'] == year)
            ]
            
            calculated_total = sub_items_df['期末金额'].sum()
            report_total = total_df.loc[total_item_name_clean, year]
            diff = calculated_total - report_total
            if abs(diff) < 0.01:
                msg = f"✅ {year}年'{total_item_name}'内部分项核对平衡 (计算值 {calculated_total:,.2f})"
                check_results.append(msg)
            else:
                msg = f"❌ {year}年'{total_item_name}'内部分项核对**不平**: 计算值 {calculated_total:,.2f} vs 报表值 {report_total:,.2f} (差异: {diff:,.2f})"
                check_results.append(msg)
    return check_results

def _check_core_equalities(total_df, raw_df, years):
    """核心勾稽关系检查函数"""
    results = []
    required = ['资产总计', '负债合计', '净资产合计', '收入合计', '费用合计']
    clean_required = {normalize_name(s) for s in required}
    
    clean_total_df_index = total_df.index.map(normalize_name)
    missing = [item for item in clean_required if item not in clean_total_df_index]
    if missing:
        results.append(f"❌ 核心勾稽关系检查失败: 缺少关键合计项 {missing}")
        return results

    start_year, end_year = years[0], years[-1]
    
    for year in years:
        asset = total_df.loc[normalize_name('资产总计'), year]
        lia = total_df.loc[normalize_name('负债合计'), year]
        equity = total_df.loc[normalize_name('净资产合计'), year]
        diff = asset - (lia + equity)
        if abs(diff) < 0.01:
            results.append(f"✅ {year}年资产负债表内部平衡")
        else:
            results.append(f"❌ {year}年资产负债表内部**不平** (差异: {diff:,.2f})")

    start_equity = raw_df[(normalize_name(raw_df['项目']) == normalize_name('净资产合计')) & (raw_df['年份'] == start_year)]['期初金额'].sum()
    end_equity = raw_df[(normalize_name(raw_df['项目']) == normalize_name('净资产合计')) & (raw_df['年份'] == end_year)]['期末金额'].sum()
    net_asset_change = end_equity - start_equity

    income = total_df.loc[normalize_name('收入合计'), years].sum()
    expense = total_df.loc[normalize_name('费用合计'), years].sum()
    net_profit = income - expense
    diff = net_asset_change - net_profit
    if abs(diff) < 0.01:
        results.append(f"✅ 跨期核心勾稽关系平衡: 净资产变动 {net_asset_change:,.2f} ≈ 收支总差额 {net_profit:,.2f}")
    else:
        results.append(f"❌ 跨期核心勾稽关系**不平**: 净资产变动 {net_asset_change:,.2f} vs 收支总差额 {net_profit:,.2f} (差异: {diff:,.2f})")
    return results