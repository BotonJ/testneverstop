import os
from openpyxl import load_workbook
import xlrd

# 设置目标路径
target_dir = r"C:\Users\Administrator\Desktop\省特种设备与节能会离任审计群\促进会审计250613\新建文件夹"

def check_xlsx(path):
    try:
        wb = load_workbook(path, read_only=True)
        visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        if visible:
            print(f"✅ [{os.path.basename(path)}] 可见工作表: {visible}")
        else:
            print(f"⚠️ [{os.path.basename(path)}] 无可见工作表")
    except Exception as e:
        print(f"❌ [{os.path.basename(path)}] 读取失败: {e}")

def check_xls(path):
    try:
        book = xlrd.open_workbook(path)
        sheets = book.sheet_names()
        if sheets:
            print(f"⚠️ [{os.path.basename(path)}] 是旧版 .xls 格式，请手动另存为 .xlsx")
        else:
            print(f"❌ [{os.path.basename(path)}] 无可读 sheet")
    except Exception as e:
        print(f"❌ [{os.path.basename(path)}] 打开失败: {e}")

# 遍历文件夹
for fname in os.listdir(target_dir):
    if fname.startswith("~$"):
        continue  # 忽略临时文件
    fpath = os.path.join(target_dir, fname)
    if fname.lower().endswith(".xlsx"):
        check_xlsx(fpath)
    elif fname.lower().endswith(".xls"):
        check_xls(fpath)
