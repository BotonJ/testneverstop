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
    category_col: str, # '所属区块' or '类型'
    item_col_name: str # '标准科目名' or '科目名称'
) -> List[str]:
    """
    【V2 - 稳定版】
    统一使用 '期末金额' 作为本期/期末值的检查列。
    能正确处理资产负债表和业务活动表的内部分项核对。
    """
    results = []
    
    if config_df is None or item_col_name not in config_df.columns:
        logger.error(f"配置错误: 用于 '{report_type}' 的配置Sheet为空或缺少 '{item_col_name}' 列。跳过核对。")
        return [f"❌ 配置错误: 无法执行'{report_type}'内部分项核对，配置缺失。"]
    
    # 获取所有需要被核对的合计项
    total_items_to_check = config_df[item_col_name].unique()
    normal_items_df = raw_df[raw_df['科目类型'] == '普通'].copy()

    for total_item in total_items_to_check:
        if pd.isna(total_item): continue
        
        # 关键步骤：根据合计项名称，筛选出所有属于它的普通子项
        # 例如，筛选出所有 category_col ('所属区块') == '流动资产合计' 的普通科目
        sub_items_df = normal_items_df[normal_items_df[category_col] == total_item]
        
        # 如果找不到任何子项，说明可能配置或提取有误，跳过此合计项的检查
        if sub_items_df.empty:
            logger.debug(f"对于合计项'{total_item}'，未能在普通科目中找到任何归属于它的子项，跳过该项核对。")
            continue

        # 按年份对子项求和
        calculated_totals = sub_items_df.groupby('年份')['期末金额'].sum()

        for year, calculated_sum in calculated_totals.items():
            try:
                # 从合计透视表中找到报表上报告的官方总数
                reported_total = pivoted_total_df.loc[total_item, ('期末金额', year)]
                
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
                # 这是您遇到的警告的来源
                results.append(f"⚠️ {year}年'{total_item}'的报表值无法在合计表中找到，跳过核对。")

    return results

def _check_core_equalities(pivoted_total_df: pd.DataFrame, summary_values: Dict) -> List[str]:
    results = []
    if pivoted_total_df.empty:
        results.append("⚠️ 无法检查核心勾稽关系，因为合计科目表为空。")
        return results

    years = sorted(pivoted_total_df.columns.get_level_values('年份').unique())

    for year in years:
        try:
            assets = pivoted_total_df.loc['资产总计', ('期末金额', year)]
            liabilities = pivoted_total_df.loc['负债合计', ('期末金额', year)]
            net_assets = pivoted_total_df.loc['净资产合计', ('期末金额', year)]
            
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
    logger.info("  -> 正在执行所有数据复核检查...")
    final_results = []
    
    # 业务活动表内部分项核对
    income_config = mapping_configs.get('业务活动表汇总注入配置')
    if income_config is not None:
        income_df = raw_df[raw_df['报表类型'] == '业务活动表'].copy()
        # '类型'列(收入/费用)作为分类依据, '科目名称'是合计项本身
        income_results = _check_subtotals(income_df, pivoted_total_df, income_config, '业务活动表', '类型', item_col_name='科目名称')
        final_results.extend(income_results)
    
    # 资产负债表内部分项核对
    alias_config = mapping_configs.get('科目等价映射')
    if alias_config is not None and '科目类型' in alias_config.columns:
        # 筛选出所有合计科目作为核对目标
        bs_total_items_config = alias_config[alias_config['科目类型'] == '合计'].copy()
        bs_df = raw_df[raw_df['报表类型'] == '资产负债表'].copy()
        # '所属区块'列作为分类依据, '标准科目名'是合计项本身
        bs_results = _check_subtotals(bs_df, pivoted_total_df, bs_total_items_config, '资产负-债表', '所属区块', item_col_name='标准科目名')
        final_results.extend(bs_results)

    # 核心勾稽关系检查
    core_results = _check_core_equalities(pivoted_total_df, summary_values)
    final_results.extend(core_results)

    logger.info("  -> 数据复核检查完毕。")
    return final_results, pd.DataFrame()