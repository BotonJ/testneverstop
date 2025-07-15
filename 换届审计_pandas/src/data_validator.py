# /src/data_validator.py

import pandas as pd
import numpy as np
import logging
from typing import List, Dict, Tuple

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def _check_subtotals(
    raw_df: pd.DataFrame,
    pivoted_total_df: pd.DataFrame,
    config_df: pd.DataFrame,
    report_type: str,
    category_col: str,
    item_col_name: str # 新增参数，指定哪个列包含要检查的科目名
) -> List[str]:
    """
    【适配版】
    通用的分项与合计交叉验证函数。
    通过 item_col_name 参数，使其能够使用您配置文件中已有的列名。
    """
    results = []
    
    if item_col_name not in config_df.columns:
        logger.error(f"配置错误: 用于 '{report_type}' 的配置Sheet缺少必需的 '{item_col_name}' 列。跳过此项核对。")
        results.append(f"❌ 配置错误: 无法执行'{report_type}'的内部分项核对，因配置缺少'{item_col_name}'列。")
        return results
    
    total_items_to_check = config_df[item_col_name].unique()
    normal_items_df = raw_df[raw_df['科目类型'] == '普通'].copy()

    for total_item in total_items_to_check:
        if pd.isna(total_item): continue
        
        sub_items_df = normal_items_df[normal_items_df[category_col] == total_item].copy()
        value_col = '期末金额' if report_type == '资产负债表' else '本期金额'
        calculated_totals = sub_items_df.groupby('年份')[value_col].sum()

        for year, calculated_sum in calculated_totals.items():
            try:
                reported_total = pivoted_total_df.loc[total_item, year]
                
                if not np.isclose(calculated_sum, reported_total, atol=0.01):
                    diff = calculated_sum - reported_total
                    results.append(
                        f"❌ {year}年'{total_item}'内部分项核对**不平**: "
                        f"计算值 {calculated_sum:,.2f} vs 报表值 {reported_total:,.2f} (差异: {diff:,.2f})"
                    )
                else:
                     results.append(
                        f"✅ {year}年'{total_item}'内部分项核对平衡 (计算值 {calculated_sum:,.2f})"
                    )
            except KeyError:
                results.append(f"⚠️ {year}年'{total_item}'的报表值无法在合计表中找到，跳过核对。")

    return results

def _check_core_equalities(pivoted_total_df: pd.DataFrame, summary_values: Dict) -> List[str]:
    results = []
    years = sorted([col for col in pivoted_total_df.columns if isinstance(col, int) or str(col).isdigit()])

    for year in years:
        try:
            assets = pivoted_total_df.loc['资产总计', year]
            liabilities = pivoted_total_df.loc['负债合计', year]
            net_assets = pivoted_total_df.loc['净资产合计', year]
            
            if not np.isclose(assets, liabilities + net_assets, atol=0.01):
                diff = assets - (liabilities + net_assets)
                results.append(f"❌ {year}年资产负债表**不平衡**: 资产 {assets:,.2f} != 负债+净资产 {liabilities + net_assets:,.2f} (差异: {diff:,.2f})")
            else:
                results.append(f"✅ {year}年资产负债表内部平衡")
        except KeyError as e:
            results.append(f"⚠️ 无法检查{year}年资产负债表平衡，缺少关键科目: {e}")

    try:
        net_asset_change = summary_values.get('净资产总额增减', 0)
        net_surplus = summary_values.get('审计期间净结余', 0)
        if not np.isclose(net_asset_change, net_surplus, atol=0.01):
            diff = net_asset_change - net_surplus
            results.append(f"❌ 跨期核心勾稽关系**不平**: 净资产变动 {net_asset_change:,.2f} vs 收支总差额 {net_surplus:,.2f} (差异: {diff:,.2f})")
        else:
            results.append("✅ 跨期核心勾稽关系平衡")
    except (KeyError, TypeError):
        results.append("⚠️ 无法检查跨期核心勾稽关系，缺少必要的汇总数据。")
        
    return results

def run_all_checks(
    raw_df: pd.DataFrame,
    pivoted_total_df: pd.DataFrame,
    summary_values: Dict,
    mapping_configs: Dict[str, pd.DataFrame]
) -> Tuple[List[str], pd.DataFrame]:
    logger.info("--- [复核机制] 开始执行所有数据检查... ---")
    final_results = []
    
    # --- [修复] 适配业务活动表配置 ---
    logger.info("   -> 正在执行: 业务活动表内部分项核对...")
    income_config = mapping_configs.get('业务活动表汇总注入配置')
    if income_config is not None and not income_config.empty:
        income_df = raw_df[raw_df['报表类型'] == '业务活动表'].copy()
        # 告诉函数，使用您已有的 '科目名称' 列
        income_results = _check_subtotals(income_df, pivoted_total_df, income_config, '业务活动表', '类型', item_col_name='科目名称')
        final_results.extend(income_results)
    else:
        logger.warning("未找到'业务活动表汇总注入配置'，跳过业务活动表内部分项核对。")

    # --- [修复] 适配资产负债表配置 ---
    logger.info("   -> 正在执行: 资产负债表内部分项核对...")
    # 不再寻找不存在的Sheet，而是智能地使用 '科目等价映射'
    alias_config = mapping_configs.get('科目等价映射')
    if alias_config is not None and not alias_config.empty and '科目类型' in alias_config.columns:
        # 从科目映射表中筛选出所有合计项目
        bs_total_items_config = alias_config[alias_config['科目类型'] == '合计'].copy()
        bs_df = raw_df[raw_df['报表类型'] == '资产负债表'].copy()
        # 告诉函数，使用 '标准科目名' 列
        bs_results = _check_subtotals(bs_df, pivoted_total_df, bs_total_items_config, '资产负债表', '所属区块', item_col_name='标准科目名')
        final_results.extend(bs_results)
    else:
        logger.warning("在'科目等价映射'中未找到'科目类型'列或配置为空，跳过资产负债表内部分项核对。")

    # --- 核心勾稽关系检查（逻辑不变） ---
    logger.info("   -> 正在执行: 核心勾稽关系检查...")
    core_results = _check_core_equalities(pivoted_total_df, summary_values)
    final_results.extend(core_results)

    logger.info("--- [复核机制] 所有数据检查执行完毕。 ---")
    return final_results, pd.DataFrame()
