import re

def _test_header_logic_v2(audit_period_str: str):
    """
    V2版本：增加了对“完整年度”的判断，以简化表头。
    - 如果起始月是1月，期初表头简化为标准的 "上年累计数"。
    - 如果终止月是12月，期末表头简化为标准的 "本年累计数"。
    """
    print(f"\n--- 开始测试审计期间: '{audit_period_str}' ---")
    
    match = re.match(r'(\d{4})年(\d{1,2})月[-至](\d{4})年(\d{1,2})月', audit_period_str.replace(" ", ""))
    if not match:
        print("错误：无法解析审计期间格式。")
        return

    start_year, start_month, end_year, end_month = map(int, match.groups())
    
    print(f"解析结果: start={start_year}年{start_month}月, end={end_year}年{end_month}月")
    print("-" * 20)

    years_to_check = range(start_year, end_year + 1)

    for year in years_to_check:
        header_qichu = f"{year - 1}年累计数"
        header_qimo = f"{year}年累计数"

        # --- 应用规则 ---
        
        # 规则A: 当前年份是【审计起始年】
        if year == start_year:
            # A.1: 处理【期初】表头
            # [V2 新增判断] 只有当起始月不是1月时，才需要特殊格式
            if start_month > 1:
                header_qichu = f"{year}年1-{start_month - 1}月累计数"
            
            # A.2: 处理【期末】表头
            # [V2 新增判断] 只有当终止月不是12月时，才需要特殊格式
            # (这个判断主要影响同一年结束的场景)
            if end_month < 12:
                 header_qimo = f"{year}年{start_month}-{end_month}月累计数"
            else: # 如果是12月结束
                 header_qimo = f"{year}年{start_month}-12月累计数"
                 # [V2 核心简化] 如果恰好是1月开始，12月结束，就简化
                 if start_month == 1:
                     header_qimo = f"{year}年累计数"


        # 规则B: 当前年份是【审计终止年】 (且不是起始年)
        if year == end_year and start_year != end_year:
             # [V2 新增判断] 只有当终止月不是12月时，才需要特殊格式
            if end_month < 12:
                header_qimo = f"{year}年1-{end_month}月累计数"

        # 规则C: 【最优先】如果审计期间在同一年内
        if start_year == end_year:
            # 此时 year == start_year == end_year
            # 期初逻辑不变 (来自规则A.1)
            # 期末逻辑需要被覆盖为 "[起始月]-[终止月]"
            # [V2 简化] 仅当不是完整年度时才使用详细格式
            if start_month > 1 or end_month < 12:
                header_qimo = f"{year}年{start_month}-{end_month}月累计数"
            else: # 如果是1月到12月，简化
                header_qimo = f"{year}年累计数"

        print(f"年份: {year} -> 计算结果: 期初='{header_qichu}', 期末='{header_qimo}'")

    print("--- 测试结束 ---\n")

# --- 运行所有测试用例 ---
if __name__ == '__main__':
    # 旧的测试用例
    _test_header_logic_v2("2019年3月至2019年11月")
    _test_header_logic_v2("2020年5月至2022年4月")
    _test_header_logic_v2("2021年1月至2021年8月")
    
    # 您新增的测试用例
    _test_header_logic_v2("2011年11月至2019年12月")
    _test_header_logic_v2("2020年1月至2022年12月")