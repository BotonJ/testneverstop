# test_runner.py
import json
from pathlib import Path
from inject_modules.inject import run_full_injection
from inject_modules.biz import inject_income_expense_all

base_dir = Path(__file__).resolve().parent
input_dir = base_dir / "data"
output_dir = base_dir / "outputs"
output_dir.mkdir(exist_ok=True)

mapping_file = input_dir / "mapping_file.xlsx"
source_file = input_dir / "source_file.xlsx"
template_file = input_dir / "template_file.xlsx"
final_output_file = output_dir / "output.xlsx"
log_file = output_dir / "inject_log.txt"

log = []
run_full_injection(
    mapping_file=mapping_file,
    source_file=source_file,
    template_file=template_file,
    output_file=final_output_file,
    log=log
)

from openpyxl import load_workbook
wb_tgt = load_workbook(final_output_file)
inject_income_expense_all(
    mapping_file=mapping_file,
    source_file=source_file,
    wb_tgt=wb_tgt
)
wb_tgt.save(final_output_file)

with open(log_file, "w", encoding="utf-8") as f:
    f.write("\n".join(log))

print(f"✅ 日志已保存至：{log_file}")
